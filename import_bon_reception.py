#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Import d'un Bon de Reception a partir d'un fichier Excel fournisseur.

Principe
--------
Au lieu de saisir les bons de reception a la main, le fournisseur exporte
un fichier Excel (liste d'articles avec leur prix de vente). Ce script
importe ce fichier directement dans la base Firebird de la gestion
commerciale sous forme d'un *Bon de reception* (type de piece PC_AC_B) :

  * Les articles qui n'existent pas encore sont crees automatiquement.
  * Le PRIX D'ACHAT HT de l'article = le prix indique sur l'Excel du
    fournisseur (c'est le prix que vous payez l'article).
  * Le PRIX DE VENTE est laisse VIDE : vous fixerez vous-meme le prix de
    vente plus tard dans le logiciel.
  * Les lignes du bon de reception alimentent le stock (QTE) avec le meme
    prix d'achat, comme une reception saisie manuellement.

Le script reproduit exactement la logique interne du logiciel :
  - NOPIECE / NOITEM sont tires des generateurs NEXTPIECE / NEXTITEM ;
  - REF_PIECE est attribue automatiquement par le trigger INSERT_PIECE ;
  - COEFF = 1 et ANNULEE = 1 sur chaque ligne  -> le stock est bien
    incremente (verifie via la procedure SPSTOCK du logiciel) ;
  - PRIXHT de la ligne = prix d'achat  -> le PUMP (cout moyen pondere)
    est calcule correctement par le logiciel.

Usage
-----
    python import_bon_reception.py --config config.json --excel fournisseur.xlsx
    python import_bon_reception.py --config config.json --excel f.xlsx --dry-run

Dependances : openpyxl, fdb   (pip install -r requirements.txt)
"""

import argparse
import json
import os
import sys
import datetime
import unicodedata

try:
    import openpyxl
except ImportError:
    sys.exit("Module manquant : openpyxl  ->  pip install openpyxl")

try:
    import fdb
except ImportError:
    sys.exit("Module manquant : fdb  ->  pip install fdb")


# --------------------------------------------------------------------------- #
#  Configuration par defaut (surchargeable par le fichier --config)
# --------------------------------------------------------------------------- #
DEFAULT_CONFIG = {
    # --- Connexion Firebird --------------------------------------------------
    "host": "localhost",            # "" ou null = acces local direct au fichier
    "port": 3050,
    "database": "C:\\\\PRIME\\\\PR22.FDB",  # chemin du fichier .FDB
    "user": "SYSDBA",
    "password": "masterkey",
    "charset": "WIN1252",           # la base est en charset NONE (octets bruts)

    # --- Type de piece -------------------------------------------------------
    "code_type_piece": "PC_AC_B",   # PC_AC_B = "Bons de reception"
    "etat": "VA",                   # etat de la piece (validee)

    # --- Tiers / depot (optionnels) -----------------------------------------
    "code_tiers": "",               # code fournisseur ; "" = aucun
    "raison_sociale": "",           # nom du fournisseur (si creation)
    "code_depot": "",               # code depot ; "" = aucun
    "create_missing_tiers": True,   # creer le fournisseur/depot s'ils manquent

    # --- Valeurs par defaut pour les nouveaux articles ----------------------
    "default_famille": "DIVERS",    # code famille par defaut si non trouvee
    "default_famille_intitule": "Divers",
    "default_unite": "U",           # code unite de base par defaut
    "default_unite_intitule": "Unite",
    "default_tva": 19,              # TVA par defaut si absente de l'Excel
    "match_famille_par_intitule": True,  # associer la colonne "Famille" a une famille existante
    "calc_prix_achat_ttc": True,    # renseigner aussi PRIXACHATTTC (= HT * (1+TVA/100))

    # --- Mapping des colonnes Excel -> alias acceptes (insensible accents/casse)
    "colonnes": {
        "ref_art":     ["ref. art.", "ref art", "reference", "ref", "code article", "code"],
        "designation": ["designation", "designation + qte", "libelle", "intitule", "article"],
        "qte":         ["qte", "quantite", "quantity", "qty"],
        "prix":        ["prix ht", "prix vente ht", "pu ht", "prix unitaire ht", "prix"],
        "tva":         ["tva", "taux tva", "taux de tva"],
        "famille":     ["famille", "rayon", "categorie"],
        "code_barres": ["code barres", "code barre", "code-barres", "codebarre",
                        "code a barre", "code a barres", "ean", "ean13", "gencode",
                        "barcode", "code barre article"],
    },
    # Code-barres : si l'Excel n'a pas de colonne code-barres, on utilise la
    # reference article (Ref. Art.) comme code-barres -> scan direct, pas de
    # ressaisie.
    "barcode_depuis_ref": True,
    # Colonne a utiliser comme prix d'achat (= prix de vente du fournisseur)
    "colonne_prix": "prix",
}


# --------------------------------------------------------------------------- #
#  Utilitaires
# --------------------------------------------------------------------------- #
def norm(s):
    """Normalise un libelle : minuscule, sans accents, espaces compactes."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def to_float(v, default=0.0):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    txt = str(v).strip().replace(" ", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return default


def load_config(path):
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # copie profonde
    if path:
        with open(path, "r", encoding="utf-8") as fh:
            user_cfg = json.load(fh)
        for k, v in user_cfg.items():
            if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                cfg[k].update(v)
            else:
                cfg[k] = v
    return cfg


# --------------------------------------------------------------------------- #
#  Lecture de l'Excel fournisseur
# --------------------------------------------------------------------------- #
def find_header_row(ws, aliases):
    """Trouve la ligne d'en-tete : celle qui contient la colonne 'ref_art'."""
    wanted = set(aliases["ref_art"])
    for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
        cells = {norm(c) for c in row if c is not None}
        if cells & wanted:
            return r, list(row)
    return None, None


def map_columns(header, aliases):
    """Associe chaque champ logique a son index de colonne dans l'Excel."""
    index = {}
    normalized = [norm(h) for h in header]
    for field, names in aliases.items():
        for i, h in enumerate(normalized):
            if h in names:
                index[field] = i
                break
    return index


def read_excel(path, cfg):
    aliases = cfg["colonnes"]
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    hrow, header = find_header_row(ws, aliases)
    if hrow is None:
        raise ValueError(
            "Impossible de trouver la ligne d'en-tete (colonne 'Ref. Art.' "
            "introuvable). Verifiez le fichier Excel.")
    colmap = map_columns(header, aliases)
    for req in ("ref_art", "qte"):
        if req not in colmap:
            raise ValueError("Colonne obligatoire manquante dans l'Excel : %s" % req)

    price_field = cfg.get("colonne_prix", "prix")
    if price_field not in colmap:
        raise ValueError(
            "Colonne de prix introuvable dans l'Excel (cherche : %s)."
            % ", ".join(aliases.get(price_field, [price_field])))

    lines = []
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        ref = row[colmap["ref_art"]] if colmap.get("ref_art") is not None else None
        if ref is None or str(ref).strip() == "":
            continue
        line = {
            "ref_art":     str(ref).strip(),
            "designation": (str(row[colmap["designation"]]).strip()
                            if "designation" in colmap and row[colmap["designation"]] is not None
                            else str(ref).strip()),
            "qte":   to_float(row[colmap["qte"]], 0.0),
            "prix":  to_float(row[colmap[price_field]], 0.0),
            "tva":   (to_float(row[colmap["tva"]], cfg["default_tva"])
                      if "tva" in colmap else cfg["default_tva"]),
            "famille": (str(row[colmap["famille"]]).strip()
                        if "famille" in colmap and row[colmap["famille"]] is not None
                        else ""),
            "code_barres": (str(row[colmap["code_barres"]]).strip()
                            if "code_barres" in colmap and row[colmap["code_barres"]] is not None
                            else ""),
        }
        lines.append(line)
    return lines


# --------------------------------------------------------------------------- #
#  Operations base de donnees
# --------------------------------------------------------------------------- #
class Importer:
    def __init__(self, con, cfg):
        self.con = con
        self.cfg = cfg
        self.cur = con.cursor()
        self._famille_cache = None

    # -- helpers ----------------------------------------------------------
    def gen_id(self, generator):
        self.cur.execute("SELECT GEN_ID(%s, 1) FROM RDB$DATABASE" % generator)
        return str(self.cur.fetchone()[0])

    def exists(self, sql, params):
        self.cur.execute(sql, params)
        return self.cur.fetchone() is not None

    # -- referentiels (familles / unites / tiers) -------------------------
    def ensure_famille(self, code, intitule):
        if not self.exists("SELECT 1 FROM FAMILLE WHERE CODEFAMILLE = ?", (code,)):
            self.cur.execute(
                "INSERT INTO FAMILLE (CODEFAMILLE, INTITULE, TAUX_TVA) VALUES (?, ?, ?)",
                (code, intitule[:50], self.cfg["default_tva"]))

    def ensure_unite(self, code, intitule):
        if not self.exists("SELECT 1 FROM UNITE WHERE CODE_UNITE = ?", (code,)):
            self.cur.execute(
                "INSERT INTO UNITE (CODE_UNITE, INTITULE, FACTEUR) VALUES (?, ?, 1)",
                (code, intitule[:30]))

    def ensure_tiers(self, code, raison):
        if code and not self.exists("SELECT 1 FROM TIERS WHERE CODE_TIERS = ?", (code,)):
            self.cur.execute(
                "INSERT INTO TIERS (CODE_TIERS, RAISON_SOCIALE, DATE_CREATION) "
                "VALUES (?, ?, ?)",
                (code, (raison or code)[:200], datetime.datetime.now()))

    def resolve_famille(self, label):
        """Associe le libelle 'Famille' de l'Excel a un code famille existant,
        sinon retourne la famille par defaut."""
        cfg = self.cfg
        if cfg.get("match_famille_par_intitule") and label:
            if self._famille_cache is None:
                self.cur.execute("SELECT CODEFAMILLE, INTITULE FROM FAMILLE")
                self._famille_cache = {norm(i): c for c, i in self.cur.fetchall()}
            code = self._famille_cache.get(norm(label))
            if code:
                return code
        return cfg["default_famille"]

    # -- articles ---------------------------------------------------------
    def upsert_article(self, line):
        """Cree l'article s'il n'existe pas. Retourne 'created' ou 'exists'."""
        ref = line["ref_art"]
        if self.exists("SELECT 1 FROM ARTICLE WHERE REF_ART = ?", (ref,)):
            return "exists"

        cfg = self.cfg
        codefamille = self.resolve_famille(line["famille"])
        prix_achat_ht = line["prix"]              # prix de vente fournisseur = notre prix d'achat
        tva = line["tva"]
        prix_achat_ttc = (round(prix_achat_ht * (1 + tva / 100.0), 4)
                          if cfg.get("calc_prix_achat_ttc") else None)

        # Code-barres : colonne de l'Excel si presente, sinon la reference
        # article (pour scanner directement sans ressaisie).
        barcode = line.get("code_barres") or ""
        if not barcode and cfg.get("barcode_depuis_ref"):
            barcode = ref
        code_barres = barcode[:60] or None        # CODE_BARRES VARCHAR(60)
        code_barre = (barcode[:35] or None) if barcode else None  # CODE_BARRE VARCHAR(35)

        self.cur.execute(
            "INSERT INTO ARTICLE "
            "(REF_ART, CODEFAMILLE, DESIGNATION, CODE_BARRES, CODE_BARRE, "
            " PRIXACHATHT, PRIXACHATTTC, PRIXVENTEHT, PRIXVENTETTC, TAUX_TVA, "
            " CODE_UNITE_BASE, CODE_UNITE_AC, CODE_UNITE_VE, EN_SOMMEIL, DATE_CREATION) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?, 0, ?)",
            (ref, codefamille, line["designation"][:100], code_barres, code_barre,
             prix_achat_ht, prix_achat_ttc, tva,
             cfg["default_unite"], cfg["default_unite"], cfg["default_unite"],
             datetime.datetime.now()))
        return "created"

    # -- piece + lignes ---------------------------------------------------
    def create_piece(self, date_piece):
        cfg = self.cfg
        nopiece = self.gen_id("NEXTPIECE")
        self.cur.execute(
            "INSERT INTO PIECE "
            "(NOPIECE, CODE_TYPE_PIECE, CODE_TIERS, CODE_DEPOT, DATEPIECE, "
            " ETAT, ANNULEE, COEFF, COEFF_TR, MONTANT, MONTANTHT, MONTANTTTC, TVA) "
            "VALUES (?, ?, ?, ?, ?, ?, 1, 1, 0, 0, 0, 0, 0)",
            (nopiece, cfg["code_type_piece"],
             cfg["code_tiers"] or None, cfg["code_depot"] or None,
             date_piece, cfg["etat"]))
        return nopiece

    def add_item(self, nopiece, line, date_piece):
        cfg = self.cfg
        noitem = self.gen_id("NEXTITEM")
        self.cur.execute(
            "INSERT INTO ITEM "
            "(NOITEM, NOPIECE, REF_ART, QTE, PRIXHT, TVA, COEFF, COEFF_TR, "
            " CODE_UNITE, CODE_TIERS, CODE_DEPOT, DATEPIECE) "
            "VALUES (?, ?, ?, ?, ?, ?, 1, 0, ?, ?, ?, ?)",
            (noitem, nopiece, line["ref_art"], line["qte"], line["prix"],
             line["tva"], cfg["default_unite"],
             cfg["code_tiers"] or None, cfg["code_depot"] or None, date_piece))
        return noitem

    def update_totaux(self, nopiece, montant_ht, tva, montant_ttc):
        self.cur.execute(
            "UPDATE PIECE SET MONTANTHT = ?, TVA = ?, MONTANTTTC = ?, MONTANT = ? "
            "WHERE NOPIECE = ?",
            (montant_ht, tva, montant_ttc, montant_ttc, nopiece))

    def get_ref_piece(self, nopiece):
        self.cur.execute("SELECT REF_PIECE FROM PIECE WHERE NOPIECE = ?", (nopiece,))
        row = self.cur.fetchone()
        return row[0] if row else None


# --------------------------------------------------------------------------- #
#  Programme principal
# --------------------------------------------------------------------------- #
def connect(cfg):
    # Permet de pointer explicitement vers la librairie cliente Firebird
    # (ex: fbclient.dll sur Windows, ou libfbembed pour un acces embedded).
    lib = cfg.get("fb_client_library")
    if lib:
        fdb.load_api(lib)
    kwargs = dict(database=cfg["database"], user=cfg["user"],
                  password=cfg["password"], charset=cfg.get("charset", "WIN1252"))
    host = cfg.get("host")
    if host:
        kwargs["host"] = host
        if cfg.get("port"):
            kwargs["port"] = int(cfg["port"])
    return fdb.connect(**kwargs)


def main():
    ap = argparse.ArgumentParser(
        description="Importe un Excel fournisseur comme Bon de reception dans la base Firebird.")
    ap.add_argument("--config", help="Fichier de configuration JSON")
    ap.add_argument("--excel", required=True, help="Fichier Excel du fournisseur (.xlsx)")
    ap.add_argument("--db", help="Chemin du .FDB (surcharge la config)")
    ap.add_argument("--date", help="Date du bon (AAAA-MM-JJ). Defaut : aujourd'hui.")
    ap.add_argument("--code-tiers", help="Code fournisseur (surcharge la config)")
    ap.add_argument("--dry-run", action="store_true",
                    help="Simulation : affiche le resultat sans rien ecrire en base.")
    args = ap.parse_args()

    cfg = load_config(args.config)
    if args.db:
        cfg["database"] = args.db
    if args.code_tiers is not None:
        cfg["code_tiers"] = args.code_tiers

    if args.date:
        date_piece = datetime.datetime.strptime(args.date, "%Y-%m-%d")
    else:
        date_piece = datetime.datetime.now()

    # 1) Lecture de l'Excel
    if not os.path.isfile(args.excel):
        sys.exit("Fichier Excel introuvable : %s" % args.excel)
    lines = read_excel(args.excel, cfg)
    if not lines:
        sys.exit("Aucune ligne d'article trouvee dans l'Excel.")
    print("Lignes lues dans l'Excel : %d" % len(lines))

    # 2) Connexion + import (transaction unique)
    con = connect(cfg)
    imp = Importer(con, cfg)
    try:
        # referentiels minimaux
        imp.ensure_famille(cfg["default_famille"], cfg["default_famille_intitule"])
        imp.ensure_unite(cfg["default_unite"], cfg["default_unite_intitule"])
        if cfg.get("create_missing_tiers"):
            imp.ensure_tiers(cfg.get("code_tiers"), cfg.get("raison_sociale"))
            imp.ensure_tiers(cfg.get("code_depot"), cfg.get("code_depot"))

        nopiece = imp.create_piece(date_piece)

        created, existing = [], []
        montant_ht = tva_tot = 0.0
        for line in lines:
            state = imp.upsert_article(line)
            (created if state == "created" else existing).append(line["ref_art"])
            imp.add_item(nopiece, line, date_piece)
            ht = line["qte"] * line["prix"]
            montant_ht += ht
            tva_tot += ht * line["tva"] / 100.0

        montant_ttc = montant_ht + tva_tot
        imp.update_totaux(nopiece, round(montant_ht, 4),
                          round(tva_tot, 4), round(montant_ttc, 4))
        ref_piece = imp.get_ref_piece(nopiece)

        # 3) Resume
        print("-" * 60)
        print("Bon de reception : NOPIECE=%s  REF_PIECE=%s  (type %s)"
              % (nopiece, ref_piece, cfg["code_type_piece"]))
        print("Fournisseur : %s    Depot : %s"
              % (cfg.get("code_tiers") or "(aucun)", cfg.get("code_depot") or "(aucun)"))
        print("Articles crees   : %d  %s"
              % (len(created), created if len(created) <= 20 else created[:20] + ["..."]))
        print("Articles existants : %d" % len(existing))
        print("Total HT  : %.2f" % montant_ht)
        print("Total TVA : %.2f" % tva_tot)
        print("Total TTC : %.2f" % montant_ttc)

        if args.dry_run:
            con.rollback()
            print("\n[DRY-RUN] Aucune modification enregistree (rollback).")
        else:
            con.commit()
            print("\nImport termine et enregistre (commit).")
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


if __name__ == "__main__":
    main()
