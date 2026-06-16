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
    # Charset de connexion (base en charset NONE = octets bruts). Le logiciel
    # ecrit selon la page de code ANSI de Windows : WIN1256 pour l'arabe (+
    # francais accentue en minuscules), WIN1252 pour le francais seul.
    "charset": "WIN1256",

    # --- Type de piece -------------------------------------------------------
    "code_type_piece": "PC_AC_B",   # PC_AC_B = "Bons de reception"
    "etat": None,                   # etat de la piece : le logiciel laisse NULL

    # --- Tiers / depot (optionnels) -----------------------------------------
    "code_tiers": "",               # code fournisseur ; "" = aucun
    "raison_sociale": "",           # nom du fournisseur (si creation)
    "code_depot": "",               # code depot ; "" = aucun
    "create_missing_tiers": True,   # creer le fournisseur/depot s'ils manquent

    # --- Valeurs par defaut pour les nouveaux articles ----------------------
    "default_famille": "TOUS",      # famille racine / de repli (code d'une famille existante)
    "default_famille_intitule": "Tous",
    "default_unite": "",            # code unite de base ; "" = aucune (comme le logiciel)
    "default_unite_intitule": "Unite",
    "default_tva": 19,              # TVA par defaut si absente de l'Excel

    # --- Prix de vente automatique (nouveaux articles seulement) ------------
    "prix_vente_auto": True,        # calculer un prix de vente = prix achat + marge
    "marge_pct": 50,                # marge appliquee au prix d'achat (50 = +50%)
    # Arrondi VERS LE HAUT par paliers : [seuil_max, pas]. null = au-dela.
    #   < 200 -> au 5 superieur ; 200-999 -> au 10 superieur ; >= 1000 -> au 50 superieur.
    "arrondi_paliers": [[200, 5], [1000, 10], [None, 50]],
    "match_famille_par_intitule": True,  # associer la colonne "Famille" a une famille existante
    "create_missing_familles": True,     # creer la famille (par son NOM) si aucune ne correspond
    "calc_prix_achat_ttc": True,    # renseigner PRIXACHATTTC (= PRIXACHATHT, identiques)

    # --- Numerotation NOPIECE / NOITEM --------------------------------------
    # Le logiciel numerote en MAX(NOPIECE)+1 (le generateur peut etre obsolete).
    # On calcule donc le prochain numero a partir du MAX existant ET du
    # generateur, en ignorant les plages reservees (ex: 8000000 pour les
    # inventaires) au-dela de ce seuil.
    "reserved_id_threshold": 1000000,

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
    # Code-barres : dans le logiciel, c'est la REFERENCE ARTICLE (Ref. Art.)
    # qui sert de code scanne (CODE_BARRES reste vide, et il porte un index
    # UNIQUE). On laisse donc CODE_BARRES vide par defaut : scanner le
    # code-barres retrouve l'article par sa reference, sans ressaisie.
    # Mettre True pour recopier malgre tout la reference dans CODE_BARRES,
    # ou fournir une colonne code-barres distincte dans l'Excel.
    "barcode_depuis_ref": False,
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


# Correspondance charset Firebird -> codec Python (pour assainir le texte).
_CHARSET_CODEC = {
    "WIN1252": "cp1252", "WIN1256": "cp1256", "WIN1250": "cp1250",
    "ISO8859_1": "latin-1", "ISO8859_15": "iso8859-15",
    "UTF8": "utf-8", "UNICODE_FSS": "utf-8", "ASCII": "ascii", "NONE": "cp1252",
}


def codec_for(charset):
    return _CHARSET_CODEC.get((charset or "WIN1252").upper(), "cp1252")


def safe_text(s, codec):
    """Rend une chaine ecrivable dans le charset de la base.
    Conserve les accents (presents dans le charset) et ne remplace que les
    caracteres impossibles a encoder (translitteration ASCII, sinon '?'),
    afin d'eviter l'echec total de l'ecriture (probleme d'encodage)."""
    if s is None:
        return None
    s = str(s)
    try:
        s.encode(codec)
        return s
    except (UnicodeEncodeError, LookupError):
        out = []
        for ch in s:
            try:
                ch.encode(codec)
                out.append(ch)
            except UnicodeEncodeError:
                repl = unicodedata.normalize("NFKD", ch).encode("ascii", "ignore").decode("ascii")
                out.append(repl if repl else "?")
        return "".join(out)


def round_price_up(value, tiers):
    """Arrondit un prix VERS LE HAUT par paliers selon sa grandeur.
    'tiers' = liste de [seuil_max, pas] ; un seuil_max null/None = au-dela.
    Ex. [[200, 5], [None, 10]] : < 200 -> au 5 superieur ; >= 200 -> au 10 superieur."""
    import math
    if value is None or value <= 0:
        return value
    for seuil, pas in tiers:
        if seuil is None or value < seuil:
            return int(math.ceil(value / pas) * pas)
    return value


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

    codec = codec_for(cfg.get("charset"))
    txt = lambda v: safe_text(v, codec)
    lines = []
    for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
        ref = row[colmap["ref_art"]] if colmap.get("ref_art") is not None else None
        if ref is None or str(ref).strip() == "":
            continue
        line = {
            "ref_art":     txt(str(ref).strip()),
            "designation": txt(str(row[colmap["designation"]]).strip()
                               if "designation" in colmap and row[colmap["designation"]] is not None
                               else str(ref).strip()),
            "qte":   to_float(row[colmap["qte"]], 0.0),
            "prix":  to_float(row[colmap[price_field]], 0.0),
            "tva":   (to_float(row[colmap["tva"]], cfg["default_tva"])
                      if "tva" in colmap else cfg["default_tva"]),
            "famille": txt(str(row[colmap["famille"]]).strip()
                           if "famille" in colmap and row[colmap["famille"]] is not None
                           else ""),
            "code_barres": txt(str(row[colmap["code_barres"]]).strip()
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
        self._fam_code_next = None
        self.codec = codec_for(cfg.get("charset"))
        # coefficients reels du type de piece (PIECE et ITEM)
        self.coeff_piece, self.coeff_piece_tr, \
            self.coeff_item, self.coeff_item_tr = self._load_type_coeffs()

    # -- helpers ----------------------------------------------------------
    def _load_type_coeffs(self):
        """Lit les coefficients du type de piece (defaut : reception +1)."""
        self.cur.execute(
            "SELECT COEFF_PIECE, COEFF_PIECE_TR, COEFF_ITEM, COEFF_ITEM_TR "
            "FROM LOCAL_TYPE_PIECE WHERE CODE_TYPE_PIECE = ?",
            (self.cfg["code_type_piece"],))
        row = self.cur.fetchone()
        if not row:
            return 1, 0, 1, 0
        return (row[0] if row[0] is not None else 1,
                row[1] if row[1] is not None else 0,
                row[2] if row[2] is not None else 1,
                row[3] if row[3] is not None else 0)

    def gen_value(self, generator):
        self.cur.execute("SELECT GEN_ID(%s, 0) FROM RDB$DATABASE" % generator)
        return self.cur.fetchone()[0] or 0

    def advance_generator(self, generator, target):
        """Avance le generateur jusqu'a 'target' (jamais en arriere)."""
        cur = self.gen_value(generator)
        if target > cur:
            self.cur.execute("SELECT GEN_ID(%s, %d) FROM RDB$DATABASE"
                             % (generator, target - cur))
            self.cur.fetchone()

    def next_base(self, generator, table, col):
        """Base de numerotation = max(MAX numerique existant, generateur).
        Reproduit le 'MAX+1' du logiciel tout en restant au-dessus du
        generateur, en ignorant les plages reservees (> seuil)."""
        threshold = int(self.cfg.get("reserved_id_threshold", 1000000))
        self.cur.execute(
            "SELECT MAX(CAST(%s AS BIGINT)) FROM %s "
            "WHERE %s SIMILAR TO '[0-9]+' AND CHAR_LENGTH(%s) <= 15 "
            "  AND CAST(%s AS BIGINT) < ?" % (col, table, col, col, col),
            (threshold,))
        mx = self.cur.fetchone()[0] or 0
        return max(int(mx), int(self.gen_value(generator)))

    def exists(self, sql, params):
        self.cur.execute(sql, params)
        return self.cur.fetchone() is not None

    # -- referentiels (familles / unites / tiers) -------------------------
    def ensure_famille(self, code, intitule):
        if not self.exists("SELECT 1 FROM FAMILLE WHERE CODEFAMILLE = ?", (code,)):
            self.cur.execute(
                "INSERT INTO FAMILLE (CODEFAMILLE, INTITULE, TAUX_TVA) VALUES (?, ?, ?)",
                (code, safe_text(intitule, self.codec)[:50], self.cfg["default_tva"]))

    def ensure_unite(self, code, intitule):
        if code and not self.exists("SELECT 1 FROM UNITE WHERE CODE_UNITE = ?", (code,)):
            self.cur.execute(
                "INSERT INTO UNITE (CODE_UNITE, INTITULE, FACTEUR) VALUES (?, ?, 1)",
                (code, safe_text(intitule, self.codec)[:30]))

    def ensure_tiers(self, code, raison, categ=None):
        """Cree le tiers s'il manque. 'categ' = 'F' (fournisseur) ou 'D' (depot)."""
        if code and not self.exists("SELECT 1 FROM TIERS WHERE CODE_TIERS = ?", (code,)):
            self.cur.execute(
                "INSERT INTO TIERS (CODE_TIERS, RAISON_SOCIALE, CATEGS, DATE_CREATION) "
                "VALUES (?, ?, ?, ?)",
                (code, safe_text(raison or code, self.codec)[:200], categ,
                 datetime.datetime.now()))

    def _load_famille_cache(self):
        if self._famille_cache is None:
            self.cur.execute("SELECT CODEFAMILLE, INTITULE FROM FAMILLE")
            self._famille_cache = {norm(i): c for c, i in self.cur.fetchall()}

    def _next_famille_code(self):
        """Genere un code famille numerique unique (max numerique + 1)."""
        if self._fam_code_next is None:
            self.cur.execute(
                "SELECT MAX(CAST(CODEFAMILLE AS INTEGER)) FROM FAMILLE "
                "WHERE CODEFAMILLE SIMILAR TO '[0-9]+' AND CHAR_LENGTH(CODEFAMILLE) <= 9")
            self._fam_code_next = int(self.cur.fetchone()[0] or 0)
        while True:
            self._fam_code_next += 1
            code = str(self._fam_code_next)
            if not self.exists("SELECT 1 FROM FAMILLE WHERE CODEFAMILLE = ?", (code,)):
                return code

    def create_famille(self, label):
        """Cree une famille portant le NOM (intitule) du libelle Excel, avec un
        code genere automatiquement, rattachee a la famille par defaut."""
        cfg = self.cfg
        code = self._next_famille_code()
        parent = cfg["default_famille"] if self.exists(
            "SELECT 1 FROM FAMILLE WHERE CODEFAMILLE = ?", (cfg["default_famille"],)) else None
        self.cur.execute(
            "INSERT INTO FAMILLE (CODEFAMILLE, CODEFAMILLE_M, INTITULE, TAUX_TVA) "
            "VALUES (?, ?, ?, ?)",
            (code, parent, safe_text(label, self.codec)[:50], cfg["default_tva"]))
        self._famille_cache[norm(label)] = code     # eviter les doublons
        return code

    def resolve_famille(self, label):
        """Associe le libelle 'Famille' de l'Excel a une famille EXISTANTE par
        son NOM (intitule). Si aucune ne correspond : cree la famille par son
        nom (si create_missing_familles), sinon retourne la famille par defaut."""
        cfg = self.cfg
        if not label:
            return cfg["default_famille"]
        if cfg.get("match_famille_par_intitule"):
            self._load_famille_cache()
            code = self._famille_cache.get(norm(label))
            if code:
                return code
        if cfg.get("create_missing_familles"):
            self._load_famille_cache()
            code = self._famille_cache.get(norm(label))
            return code if code else self.create_famille(label)
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
        # Prix d'achat TTC = HT (identiques). Renseigne par defaut.
        prix_achat_ttc = prix_achat_ht if cfg.get("calc_prix_achat_ttc", True) else None

        # Prix de vente automatique = prix achat + marge, arrondi vers le haut.
        # Laisse vide si desactive (vous fixez le prix vous-meme).
        prix_vente_ht = prix_vente_ttc = None
        if cfg.get("prix_vente_auto"):
            brut = prix_achat_ht * (1 + cfg.get("marge_pct", 50) / 100.0)
            prix_vente_ht = round_price_up(brut, cfg.get("arrondi_paliers",
                                                          [[200, 5], [1000, 10], [None, 50]]))
            prix_vente_ttc = (round(prix_vente_ht * (1 + tva / 100.0), 4)
                              if prix_vente_ht is not None else None)

        # Code-barres : le logiciel utilise la REFERENCE comme code scanne et
        # laisse CODE_BARRES vide (index UNIQUE). On ne le renseigne donc que
        # si l'Excel fournit une colonne code-barres distincte, ou si
        # barcode_depuis_ref est explicitement active.
        barcode = line.get("code_barres") or ""
        if not barcode and cfg.get("barcode_depuis_ref"):
            barcode = ref
        code_barres = barcode[:60] or None        # CODE_BARRES VARCHAR(60)
        code_barre = (barcode[:35] or None) if barcode else None  # CODE_BARRE VARCHAR(35)

        # Unite de base : vide par defaut (comme le logiciel)
        unite = cfg.get("default_unite") or None

        self.cur.execute(
            "INSERT INTO ARTICLE "
            "(REF_ART, CODEFAMILLE, DESIGNATION, CODE_BARRES, CODE_BARRE, "
            " PRIXACHATHT, PRIXACHATTTC, PRIXVENTEHT, PRIXVENTETTC, TAUX_TVA, "
            " CODE_UNITE_BASE, CODE_UNITE_AC, CODE_UNITE_VE, DATE_CREATION) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (ref, codefamille, line["designation"][:100], code_barres, code_barre,
             prix_achat_ht, prix_achat_ttc, prix_vente_ht, prix_vente_ttc, tva,
             unite, unite, unite, datetime.datetime.now()))
        return "created"

    # -- piece + lignes ---------------------------------------------------
    def create_piece(self, nopiece, date_piece):
        cfg = self.cfg
        self.cur.execute(
            "INSERT INTO PIECE "
            "(NOPIECE, CODE_TYPE_PIECE, CODE_TIERS, CODE_DEPOT, DATEPIECE, "
            " ETAT, ANNULEE, COEFF, COEFF_TR, MONTANT, MONTANTHT, MONTANTTTC, TVA, "
            " USERNAME) "
            "VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?, 0, 0, 0, 0, ?)",
            (nopiece, cfg["code_type_piece"],
             cfg["code_tiers"] or None, cfg["code_depot"] or None,
             date_piece, cfg.get("etat") or None,
             self.coeff_piece, self.coeff_piece_tr, cfg.get("user")))
        return nopiece

    def add_item(self, noitem, nopiece, line, date_piece):
        cfg = self.cfg
        unite = cfg.get("default_unite") or None
        self.cur.execute(
            "INSERT INTO ITEM "
            "(NOITEM, NOPIECE, REF_ART, QTE, PRIXHT, TVA, COEFF, COEFF_TR, "
            " CODE_UNITE, CODE_TIERS, CODE_DEPOT, DATEPIECE) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (noitem, nopiece, line["ref_art"], line["qte"], line["prix"],
             line["tva"], self.coeff_item, self.coeff_item_tr, unite,
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
                  password=cfg["password"], charset=cfg.get("charset", "WIN1256"))
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
        # referentiels minimaux  (PIECE.USERNAME = utilisateur de connexion)
        imp.ensure_famille(cfg["default_famille"], cfg["default_famille_intitule"])
        imp.ensure_unite(cfg.get("default_unite"), cfg["default_unite_intitule"])
        if cfg.get("create_missing_tiers"):
            imp.ensure_tiers(cfg.get("code_tiers"), cfg.get("raison_sociale"), "F")
            imp.ensure_tiers(cfg.get("code_depot"), cfg.get("code_depot"), "D")

        # numerotation collision-safe (MAX existant ou generateur)
        nopiece = str(imp.next_base("NEXTPIECE", "PIECE", "NOPIECE") + 1)
        item_no = imp.next_base("NEXTITEM", "ITEM", "NOITEM")

        imp.create_piece(nopiece, date_piece)

        created, existing = [], []
        montant_ht = tva_tot = 0.0
        for line in lines:
            state = imp.upsert_article(line)
            (created if state == "created" else existing).append(line["ref_art"])
            item_no += 1
            imp.add_item(str(item_no), nopiece, line, date_piece)
            ht = line["qte"] * line["prix"]
            montant_ht += ht
            tva_tot += ht * line["tva"] / 100.0

        montant_ttc = montant_ht + tva_tot
        imp.update_totaux(nopiece, round(montant_ht, 4),
                          round(tva_tot, 4), round(montant_ttc, 4))
        # tenir les generateurs a jour (au cas ou l'appli les utilise)
        imp.advance_generator("NEXTPIECE", int(nopiece))
        imp.advance_generator("NEXTITEM", item_no)
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
    except Exception as exc:
        con.rollback()
        msg = str(exc)
        print("\nECHEC : aucune ecriture (transaction annulee).", file=sys.stderr)
        print("Detail : %s" % msg, file=sys.stderr)
        if ("transliterate" in msg.lower() or "encod" in msg.lower()
                or "malformed" in msg.lower() or "character set" in msg.lower()):
            print("-> Probleme d'encodage. Le texte est pourtant assaini vers le "
                  "charset configure (%s). Verifiez la cle \"charset\" du config "
                  "(WIN1252 recommande pour ces bases ; sinon NONE)."
                  % cfg.get("charset"), file=sys.stderr)
        sys.exit(1)
    finally:
        con.close()


if __name__ == "__main__":
    main()
