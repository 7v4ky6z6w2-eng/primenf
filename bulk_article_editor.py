#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bulk_article_editor.py
======================

Editeur EN MASSE des articles de la base PRIME (Firebird), avec interface
graphique (Tkinter). Permet de modifier rapidement, pour plusieurs articles
a la fois :

    * le PRIX DE VENTE  (PRIXVENTEHT / PRIXVENTETTC)
    * la REFERENCE      (REF_ART)
    * le CODE-BARRES    (CODE_BARRES / CODE_BARRE)

Fonctions principales
---------------------
  - Connexion a la base via config.json (meme format que votre import).
  - Recherche / filtre des articles (reference, code-barres, designation).
  - Edition d'une cellule par double-clic (valeur tracee "en attente").
  - Operations en masse sur la SELECTION :
        * Prix : fixer / +%, -% / +montant, -montant / arrondi (.99, .95, 0,50...)
                 au choix sur le HT ou le TTC, avec recalcul automatique de
                 l'autre via le taux de TVA.
        * Code-barres : fixer / vider / recopier la reference.
        * Reference : prefixe / suffixe / chercher-remplacer.
  - Recap des modifications en attente, puis ENREGISTRER (commit) ou ANNULER
    (rollback) — le tout dans une seule transaction.
  - Import d'un fichier Excel/CSV pour mettre a jour prix & code-barres par
    reference, et export de la vue courante en CSV.

Lancement
---------
    python bulk_article_editor.py                 # demande le config.json
    python bulk_article_editor.py --config config.json
    python bulk_article_editor.py --demo          # donnees de demo, sans base

Dependances : fdb (Firebird), openpyxl (Excel). Tkinter est livre avec Python.
"""

from __future__ import annotations

import argparse
import csv
import os
import sys

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from editor_logic import (Cols, PriceOp, TextOp, fmt_price, parse_number,
                          encoded_len)
import article_db
from article_db import ArticleRepository, DemoRepository, DBError, load_config, save_config


APP_TITLE = "PRIME — Editeur en masse des articles"
ROUND_CHOICES = [("(aucun)", None), (".99", "0.99"), (".95", "0.95"),
                 (".90", "0.90"), ("0,50", "0.50"), ("0,10", "0.10"),
                 ("0,05", "0.05"), ("entier", "1")]
CHARSETS = ["WIN1252", "ISO8859_1", "ISO8859_15", "UTF8", "NONE", "DOS850"]


def friendly_error(raw):
    """Traduit les pannes Firebird courantes en francais lisible."""
    t = (raw or "").lower()
    rules = [
        (("no module named 'fdb'", "module 'fdb'", "import fdb"),
         "Le pilote Firebird « fdb » n'est pas installe (py -m pip install fdb)."),
        (("fbclient", "client library", "load_api", "libfbclient",
          "valid win32", "specified module", "dll load"),
         "Librairie cliente Firebird introuvable ou de mauvaise architecture "
         "(fbclient.dll 32/64 bits). Indiquez son chemin dans « fbclient »."),
        (("no such file", "cannot open", "unable to open", "i/o error",
          "error while trying to open file", "no permission",
          "unavailable database", "file is not a valid database"),
         "Base introuvable ou inaccessible : verifiez le chemin du fichier .FDB "
         "(ex : C:/oransoft/netfact2/data/primeoffice2026.fdb)."),
        (("user name", "login", "incorrect user", "install_check",
          "your user name and password", "password"),
         "Identifiant ou mot de passe incorrect (par defaut SYSDBA / masterkey)."),
        (("connection refused", "rejected", "network", "failed to establish",
          "unable to complete network"),
         "Serveur Firebird injoignable : verifiez l'hote / le port, ou laissez "
         "l'hote VIDE pour un acces local au fichier (embedded)."),
        (("malformed string", "transliteration", "charset", "character set"),
         "Probleme d'encodage : essayez un autre charset (WIN1252 ou NONE)."),
        (("unsupported on-disk structure", "ods"),
         "Version de base (ODS) non supportee par la librairie cliente "
         "installee : utilisez le fbclient.dll fourni avec votre logiciel PRIME."),
        (("table", "introuvable dans la base"),
         "Table des articles introuvable : verifiez le nom de la table "
         "(par defaut ARTICLE)."),
    ]
    for keys, msg in rules:
        if any(k in t for k in keys):
            return msg
    s = (raw or "").strip()
    return s.splitlines()[-1][:300] if s else "Erreur inconnue."


# --------------------------------------------------------------------------- #
#  Boite de dialogue de connexion
# --------------------------------------------------------------------------- #
class ConnectDialog(tk.Toplevel):
    """Ecran de connexion : choisir la base, tester la connexion, se connecter.

    A la fermeture :
      * self.result = dict de config si l'utilisateur s'est connecte (None sinon)
      * self.repo   = le repository deja connecte (pour eviter une reconnexion)
      * self.demo   = True si l'utilisateur a choisi le mode demonstration
    """

    def __init__(self, master, cfg, allow_demo=True):
        super().__init__(master)
        self.title("Connexion a la base PRIME (Firebird)")
        self.resizable(False, False)
        self.result = None
        self.repo = None
        self.demo = False
        self.cfg = dict(cfg)
        self.vars = {}

        frm = ttk.Frame(self, padding=14)
        frm.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frm, text="Base de donnees PRIME (Firebird)",
                  font=("", 11, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, foreground="#666",
                  text="Choisissez le fichier .FDB puis testez la connexion avant d'ouvrir."
                  ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        rows = [
            ("database", "Fichier .FDB", 46, "file"),
            ("host", "Hote (vide = acces local au fichier)", 30, None),
            ("port", "Port", 30, None),
            ("user", "Utilisateur", 30, None),
            ("password", "Mot de passe", 30, None),
            ("charset", "Charset", 30, "charset"),
            ("table", "Table des articles", 30, None),
            ("fb_client_library", "fbclient.dll (optionnel)", 46, "file"),
        ]
        r = 2
        for key, label, width, kind in rows:
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=2, padx=(0, 8))
            var = tk.StringVar(value=str(self.cfg.get(key, "") or ""))
            self.vars[key] = var
            if kind == "charset":
                w = ttk.Combobox(frm, textvariable=var, values=CHARSETS, width=width - 3)
            else:
                show = "*" if key == "password" else ""
                w = ttk.Entry(frm, textvariable=var, width=width, show=show)
            w.grid(row=r, column=1, sticky="we", pady=2)
            if kind == "file":
                ttk.Button(frm, text="Parcourir…", width=11,
                           command=lambda v=var, k=key: self._browse(v, k)
                           ).grid(row=r, column=2, padx=4)
            r += 1

        # ligne test connexion
        test_row = ttk.Frame(frm)
        test_row.grid(row=r, column=0, columnspan=3, sticky="we", pady=(10, 2))
        ttk.Button(test_row, text="Tester la connexion", command=self._test).pack(side="left")
        self.conn_status = tk.StringVar(value="")
        self.status_lbl = tk.Label(test_row, textvariable=self.conn_status, anchor="w",
                                   justify="left", wraplength=420)
        self.status_lbl.pack(side="left", padx=10, fill="x", expand=True)
        r += 1

        # boutons
        btns = ttk.Frame(frm)
        btns.grid(row=r, column=0, columnspan=3, pady=(12, 0), sticky="e")
        if allow_demo:
            ttk.Button(btns, text="Mode demo", command=self._demo).pack(side="left", padx=4)
        ttk.Button(btns, text="Se connecter", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="Annuler", command=self.destroy).pack(side="left")

        frm.columnconfigure(1, weight=1)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(master)
        self.grab_set()
        self.update_idletasks()
        self._center(master)

    def _center(self, master):
        try:
            self.geometry("+%d+%d" % (master.winfo_rootx() + 60, master.winfo_rooty() + 60))
        except tk.TclError:
            pass

    def _browse(self, var, key):
        if key == "fb_client_library":
            path = filedialog.askopenfilename(
                parent=self, title="Choisir fbclient",
                filetypes=[("Librairies", "*.dll *.so *.dylib"), ("Tous", "*.*")])
        else:
            path = filedialog.askopenfilename(
                parent=self, title="Choisir le fichier de base Firebird",
                filetypes=[("Base Firebird", "*.fdb *.FDB *.gdb"), ("Tous", "*.*")])
        if path:
            var.set(path)

    def _collect(self):
        cfg = dict(self.cfg)
        for key, var in self.vars.items():
            cfg[key] = var.get().strip()
        return cfg

    def _set_status(self, state, msg):
        color = {True: "#1f7a33", False: "#a11", None: "#555"}[state]
        mark = {True: "✓ ", False: "✗ ", None: "" }[state]
        self.conn_status.set(mark + msg)
        self.status_lbl.configure(fg=color)
        self.update_idletasks()

    def _test(self):
        cfg = self._collect()
        if not cfg.get("database"):
            self._set_status(False, "Indiquez d'abord le fichier .FDB.")
            return
        self._set_status(None, "Connexion en cours…")
        repo = None
        try:
            repo = ArticleRepository(cfg).connect()
            n = repo.count()
            fam = len(repo.list_familles()) if repo.has_famille_ref() else 0
            self._set_status(True, "Connexion reussie : %d article(s), %d famille(s) "
                             "(table %s)." % (n, fam, repo.table))
        except DBError as exc:
            self._set_status(False, friendly_error(str(exc)))
        except Exception as exc:                       # noqa: BLE001
            self._set_status(False, friendly_error(str(exc)))
        finally:
            if repo is not None:
                repo.close()

    def _ok(self):
        cfg = self._collect()
        if not cfg.get("database"):
            self._set_status(False, "Indiquez le fichier .FDB.")
            return
        self._set_status(None, "Connexion en cours…")
        try:
            self.repo = ArticleRepository(cfg).connect()
        except DBError as exc:
            self._set_status(False, friendly_error(str(exc)))
            return
        except Exception as exc:                       # noqa: BLE001
            self._set_status(False, friendly_error(str(exc)))
            return
        self.result = cfg
        self.destroy()

    def _demo(self):
        self.demo = True
        self.result = None
        self.destroy()


# --------------------------------------------------------------------------- #
#  Application principale
# --------------------------------------------------------------------------- #
class BulkEditorApp(ttk.Frame):
    def __init__(self, master, repo, config_path=None):
        super().__init__(master, padding=6)
        self.master = master
        self.repo = repo
        self.config_path = config_path
        self.pack(fill="both", expand=True)

        self.rows = []            # liste de dicts (donnees affichees)
        self.row_by_iid = {}      # iid Treeview -> dict
        self.pending = {}         # ref0 -> {colonne_logique: nouvelle_valeur}
        self.new_familles = {}    # code -> (intitule, tva) familles a creer au commit
        self.columns = repo.display_columns()

        self._build_toolbar()
        self._build_table()
        self._build_bulk_panel()
        self._build_statusbar()
        self.reload()

    # -- changer de base --------------------------------------------------
    def change_database(self):
        """Rouvre l'ecran de connexion et reconstruit l'application sur la base
        choisie (le schema peut differer : on recree toute l'interface)."""
        if (self.pending or self.new_familles) and not messagebox.askyesno(
                APP_TITLE, "Des modifications ne sont pas enregistrees. "
                "Changer de base et les abandonner ?"):
            return
        start_cfg = getattr(self.repo, "cfg", None) or dict(article_db.DEFAULT_CONFIG)
        repo, cfg = prompt_connection(self.master, start_cfg)
        if repo is None:
            return
        try:
            self.repo.close()
        except Exception:                              # noqa: BLE001
            pass
        if not isinstance(repo, DemoRepository):
            try:
                save_config(self.config_path or _default_config_path(), cfg)
            except OSError:
                pass
        new_app = BulkEditorApp(self.master, repo, self.config_path)
        self.master._app = new_app                     # pour la fermeture propre
        self.destroy()

    # -- construction de l'interface -------------------------------------
    def _build_toolbar(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 6))

        ttk.Label(bar, text="Recherche :").pack(side="left")
        self.search_var = tk.StringVar()
        ent = ttk.Entry(bar, textvariable=self.search_var, width=28)
        ent.pack(side="left", padx=(4, 4))
        ent.bind("<Return>", lambda e: self.reload())
        ttk.Button(bar, text="Filtrer", command=self.reload).pack(side="left")
        ttk.Button(bar, text="Tout afficher",
                   command=lambda: (self.search_var.set(""), self.reload())).pack(side="left", padx=4)

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="Changer de base…", command=self.change_database).pack(side="left", padx=2)
        ttk.Button(bar, text="Importer Excel/CSV...", command=self.import_file).pack(side="left", padx=2)
        ttk.Button(bar, text="Exporter CSV...", command=self.export_csv).pack(side="left", padx=2)

        self.save_btn = ttk.Button(bar, text="Enregistrer (0)", command=self.commit_changes)
        self.save_btn.pack(side="right", padx=2)
        self.cancel_btn = ttk.Button(bar, text="Annuler les modifs", command=self.discard_changes)
        self.cancel_btn.pack(side="right", padx=2)

    def _build_table(self):
        wrap = ttk.Frame(self)
        wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(wrap, columns=self.columns, show="headings",
                                 selectmode="extended")
        widths = {Cols.REF: 90, Cols.DESIGNATION: 230, Cols.CODE_BARRES: 130,
                  Cols.CODE_BARRE: 110, Cols.PV_HT: 90, Cols.PV_TTC: 90,
                  Cols.TVA: 60, Cols.PA_HT: 90, Cols.FAMILLE: 110}
        for c in self.columns:
            self.tree.heading(c, text=Cols.LABELS.get(c, c))
            anchor = "e" if c in Cols.NUMERIC else "w"
            self.tree.column(c, width=widths.get(c, 100), anchor=anchor, stretch=False)
        self.tree.tag_configure("modif", background="#fff3bf")     # jaune = modifie
        self.tree.tag_configure("editable_hint", background="#ffffff")

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_double_click)

    def _build_bulk_panel(self):
        panel = ttk.LabelFrame(self, text="Operations en masse (sur la selection)", padding=8)
        panel.pack(fill="x", pady=(6, 0))

        # --- Prix de vente ---
        price = ttk.Frame(panel)
        price.grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Label(price, text="PRIX DE VENTE", font=("", 9, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(price, text="(HT et TTC = le meme prix ; la TVA n'est pas modifiee)",
                  foreground="#888").grid(row=1, column=0, columnspan=4, sticky="w")
        self.price_mode = tk.StringVar(value="set")
        modes = [("Fixer a", "set"), ("+ %", "inc_pct"), ("- %", "dec_pct"),
                 ("+ montant", "inc_amount"), ("- montant", "dec_amount"),
                 ("Arrondir seul.", "round")]
        self.price_combo = ttk.Combobox(price, state="readonly", width=14,
                                        values=[m[0] for m in modes])
        self.price_combo.current(0)
        self.price_combo.grid(row=2, column=0, columnspan=2, padx=2, pady=(4, 0), sticky="w")
        self._price_modes = dict(modes)
        self.price_value = tk.StringVar()
        ttk.Entry(price, textvariable=self.price_value, width=10).grid(
            row=2, column=2, padx=4, pady=(4, 0))
        ttk.Label(price, text="Arrondi :").grid(row=3, column=0, sticky="e", pady=(4, 0))
        self.round_combo = ttk.Combobox(price, state="readonly", width=10,
                                       values=[r[0] for r in ROUND_CHOICES])
        self.round_combo.current(0)
        self.round_combo.grid(row=3, column=1, sticky="w", pady=(4, 0))
        ttk.Button(price, text="Appliquer prix", command=self.apply_price).grid(
            row=3, column=2, columnspan=2, sticky="e", pady=(4, 0))

        # --- TVA (taux), modifiable independamment du prix ---
        if self.repo.has(Cols.TVA):
            ttk.Separator(price, orient="horizontal").grid(
                row=4, column=0, columnspan=4, sticky="ew", pady=6)
            ttk.Label(price, text="TVA %").grid(row=5, column=0, sticky="e")
            self.tva_value = tk.StringVar()
            ttk.Entry(price, textvariable=self.tva_value, width=8).grid(
                row=5, column=1, sticky="w")
            ttk.Button(price, text="Fixer la TVA", command=self.apply_tva).grid(
                row=5, column=2, columnspan=2, sticky="e")

        ttk.Separator(panel, orient="vertical").grid(row=0, column=1, sticky="ns", padx=8)

        # --- Code-barres ---
        cb = ttk.Frame(panel)
        cb.grid(row=0, column=2, sticky="w", padx=(0, 16))
        ttk.Label(cb, text="CODE-BARRES", font=("", 9, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        self.cb_value = tk.StringVar()
        ttk.Entry(cb, textvariable=self.cb_value, width=18).grid(row=1, column=0, columnspan=2, pady=2)
        ttk.Button(cb, text="Fixer", width=8,
                   command=lambda: self.apply_text(Cols.CODE_BARRES, TextOp("set", self.cb_value.get()))
                   ).grid(row=1, column=2, padx=2)
        ttk.Button(cb, text="Recopier la reference", width=20,
                   command=self.copy_ref_to_barcode).grid(row=2, column=0, columnspan=2, pady=2, sticky="w")
        ttk.Button(cb, text="Vider", width=8,
                   command=lambda: self.apply_text(Cols.CODE_BARRES, TextOp("clear"))
                   ).grid(row=2, column=2, padx=2)

        ttk.Separator(panel, orient="vertical").grid(row=0, column=3, sticky="ns", padx=8)

        # --- Reference ---
        ref = ttk.Frame(panel)
        ref.grid(row=0, column=4, sticky="w")
        ttk.Label(ref, text="REFERENCE", font=("", 9, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(ref, text="Chercher").grid(row=1, column=0, sticky="e")
        self.ref_find = tk.StringVar()
        ttk.Entry(ref, textvariable=self.ref_find, width=12).grid(row=1, column=1, pady=2)
        ttk.Label(ref, text="Remplacer").grid(row=2, column=0, sticky="e")
        self.ref_repl = tk.StringVar()
        ttk.Entry(ref, textvariable=self.ref_repl, width=12).grid(row=2, column=1, pady=2)
        ttk.Button(ref, text="Remplacer", command=self.apply_ref_replace).grid(
            row=1, column=2, rowspan=2, padx=4)
        ttk.Label(ref, text="(la reference est aussi le code scanne)",
                  foreground="#888").grid(row=3, column=0, columnspan=3, sticky="w")

        if not self.repo.has(Cols.REF):
            ref.grid_remove()

        # --- Famille ---
        if self.repo.has(Cols.FAMILLE) and getattr(self.repo, "has_famille_ref", lambda: False)():
            ttk.Separator(panel, orient="vertical").grid(row=0, column=5, sticky="ns", padx=8)
            fam = ttk.Frame(panel)
            fam.grid(row=0, column=6, sticky="w")
            ttk.Label(fam, text="FAMILLE", font=("", 9, "bold")).grid(
                row=0, column=0, columnspan=3, sticky="w")
            ttk.Label(fam, text="Existante :").grid(row=1, column=0, sticky="e")
            self.fam_combo = ttk.Combobox(fam, state="readonly", width=22)
            self.fam_combo.grid(row=1, column=1, pady=2)
            ttk.Button(fam, text="Affecter", width=9,
                       command=self.apply_famille_existing).grid(row=1, column=2, padx=2)
            ttk.Label(fam, text="Nouvelle :").grid(row=2, column=0, sticky="e")
            self.fam_new_code = tk.StringVar()
            self.fam_new_name = tk.StringVar()
            nf = ttk.Frame(fam)
            nf.grid(row=2, column=1, sticky="w", pady=2)
            ttk.Entry(nf, textvariable=self.fam_new_code, width=8).pack(side="left")
            ttk.Label(nf, text="code").pack(side="left", padx=(2, 6))
            ttk.Entry(nf, textvariable=self.fam_new_name, width=14).pack(side="left")
            ttk.Label(nf, text="nom").pack(side="left", padx=2)
            ttk.Button(fam, text="Creer & affecter", width=15,
                       command=self.apply_famille_new).grid(row=2, column=2, padx=2)
            self._refresh_famille_combo()

    def _build_statusbar(self):
        self.status = tk.StringVar()
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(6, 0))
        ttk.Label(bar, textvariable=self.status, anchor="w").pack(side="left")

    # -- chargement / affichage ------------------------------------------
    def reload(self):
        if (self.pending or self.new_familles) and not messagebox.askyesno(
                APP_TITLE,
                "Des modifications ne sont pas enregistrees. Les abandonner et recharger ?"):
            return
        self.pending.clear()
        self.new_familles.clear()
        try:
            self.repo.rollback()                 # annule toute ecriture non validee
        except Exception:                        # noqa: BLE001
            pass
        try:
            self.rows = self.repo.load(self.search_var.get())
        except DBError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self._populate()
        self._refresh_famille_combo()
        self._update_save_button()

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        self.row_by_iid.clear()
        for rec in self.rows:
            iid = self._insert_row(rec)
            self.row_by_iid[iid] = rec
        try:
            total = self.repo.count()
        except Exception:                               # noqa: BLE001
            total = len(self.rows)
        self.status.set("%d article(s) affiche(s) sur %d  —  table %s"
                        % (len(self.rows), total, self.repo.table))

    def _row_values(self, rec):
        vals = []
        for c in self.columns:
            v = rec.get(c)
            vals.append(fmt_price(v) if c in Cols.NUMERIC else ("" if v is None else str(v)))
        return vals

    def _insert_row(self, rec):
        return self.tree.insert("", "end", values=self._row_values(rec))

    def _refresh_row(self, iid, rec):
        self.tree.item(iid, values=self._row_values(rec))
        ref0 = rec.get("__ref0__")
        self.tree.item(iid, tags=("modif",) if ref0 in self.pending else ())

    # -- edition d'une cellule (double-clic) -----------------------------
    def _on_double_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not iid or not col_id:
            return
        col_idx = int(col_id[1:]) - 1
        logical = self.columns[col_idx]
        if logical not in Cols.EDITABLE:
            messagebox.showinfo(APP_TITLE, "Cette colonne n'est pas modifiable ici.")
            return
        rec = self.row_by_iid[iid]
        x, y, w, h = self.tree.bbox(iid, col_id)
        old = rec.get(logical)
        edit = tk.Entry(self.tree)
        edit.insert(0, "" if old is None else (fmt_price(old) if logical in Cols.NUMERIC else str(old)))
        edit.select_range(0, "end")
        edit.focus_set()
        edit.place(x=x, y=y, width=w, height=h)

        def commit(_=None):
            new = edit.get()
            edit.destroy()
            if logical == Cols.FAMILLE:
                self._edit_famille_inline(iid, rec, new)
            else:
                self._set_cell(iid, rec, logical, new)

        edit.bind("<Return>", commit)
        edit.bind("<Escape>", lambda e: edit.destroy())
        edit.bind("<FocusOut>", commit)

    def _set_cell(self, iid, rec, logical, new_value):
        """Valide et applique une nouvelle valeur de cellule (en attente)."""
        if logical in Cols.NUMERIC:
            num = parse_number(new_value)
            if new_value.strip() != "" and num is None:
                messagebox.showwarning(APP_TITLE, "Valeur numerique invalide : %r" % new_value)
                return
            value = num
            # le prix de vente HT et TTC restent identiques (pas la TVA)
            if logical in Cols.PRICE_FIELDS:
                self._sync_price(rec, logical, value)
        else:
            value = new_value.strip() or None
            real = self.repo.real(logical)
            maxb = self.repo.maxlen.get(real, Cols.DEFAULT_MAX_LEN.get(logical, 255))
            if value and encoded_len(value, self.repo.codec) > maxb:
                messagebox.showwarning(
                    APP_TITLE, "Trop long pour %s (max %d caracteres). Sera tronque."
                    % (Cols.LABELS.get(logical, logical), maxb))
            if logical == Cols.REF and value and value != rec.get(Cols.REF):
                if not self._warn_ref_rename():
                    return
        rec[logical] = value
        self._stage(rec, logical, value)
        self._refresh_row(iid, rec)
        self._update_save_button()

    def _sync_price(self, rec, logical, value):
        """Maintient PRIXVENTEHT et PRIXVENTETTC IDENTIQUES.

        Dans la base PRIME, le prix de vente HT et le prix de vente TTC
        contiennent la MEME valeur (le prix saisi) ; la TVA est portee a part
        par TAUX_TVA. On recopie donc simplement la valeur sur l'autre champ,
        sans aucun calcul base sur la TVA (ce qui evitait de fausser le taux)."""
        for other in Cols.PRICE_FIELDS:
            if other != logical and self.repo.has(other):
                rec[other] = value
                self._stage(rec, other, value)

    _ref_warned = False

    def _warn_ref_rename(self):
        if not BulkEditorApp._ref_warned:
            BulkEditorApp._ref_warned = True
            return messagebox.askyesno(
                APP_TITLE,
                "Vous modifiez une REFERENCE article (REF_ART).\n\n"
                "La reference est la cle de l'article et peut etre utilisee par "
                "les lignes de pieces (stock, ventes). Selon votre base, la "
                "renommer peut etre refuse ou necessiter une mise a jour liee.\n\n"
                "Continuer quand meme ?")
        return True

    # -- gestion des modifications en attente ----------------------------
    def _stage(self, rec, logical, value):
        ref0 = rec.get("__ref0__")
        self.pending.setdefault(ref0, {})[logical] = value

    def _update_save_button(self):
        extra = (" +%d fam." % len(self.new_familles)) if self.new_familles else ""
        self.save_btn.config(text="Enregistrer (%d)%s" % (len(self.pending), extra))

    def discard_changes(self):
        if not self.pending and not self.new_familles:
            return
        if messagebox.askyesno(APP_TITLE, "Abandonner les %d modification(s) en attente ?"
                               % len(self.pending)):
            self.reload()

    def commit_changes(self):
        if not self.pending and not self.new_familles:
            messagebox.showinfo(APP_TITLE, "Aucune modification a enregistrer.")
            return
        changes = [{"ref0": ref0, "values": vals} for ref0, vals in self.pending.items()]
        new_fam = [(code, name, tva) for code, (name, tva) in self.new_familles.items()]
        detail = self._summary(changes)
        if new_fam:
            detail += "\n\nNouvelles familles : " + ", ".join(
                "%s (%s)" % (c, n) for c, n, _ in new_fam)
        if not messagebox.askyesno(APP_TITLE,
                                   "Enregistrer %d article(s) modifie(s) ?\n\n%s"
                                   % (len(changes), detail)):
            return
        try:
            n = self.repo.update_rows(changes, new_familles=new_fam)
            self.repo.commit()
        except DBError as exc:
            self.repo.rollback()
            messagebox.showerror(APP_TITLE, "Echec — rien n'a ete enregistre.\n\n%s" % exc)
            return
        messagebox.showinfo(APP_TITLE, "%d article(s) enregistre(s)%s."
                            % (n, (" + %d famille(s)" % len(new_fam)) if new_fam else ""))
        self.reload()

    def _summary(self, changes, limit=12):
        lines = []
        for ch in changes[:limit]:
            parts = ", ".join("%s=%s" % (Cols.LABELS.get(k, k),
                                         fmt_price(v) if k in Cols.NUMERIC else v)
                              for k, v in ch["values"].items())
            lines.append("  %s : %s" % (ch["ref0"], parts))
        if len(changes) > limit:
            lines.append("  ... (%d de plus)" % (len(changes) - limit))
        return "\n".join(lines)

    # -- operations en masse ---------------------------------------------
    def _selected_recs(self):
        recs = [self.row_by_iid[i] for i in self.tree.selection()]
        if not recs:
            messagebox.showinfo(APP_TITLE, "Selectionnez d'abord une ou plusieurs lignes "
                                "(Ctrl+clic / Maj+clic).")
        return recs

    def apply_price(self):
        recs = self._selected_recs()
        if not recs:
            return
        mode = self._price_modes[self.price_combo.get()]
        round_to = dict(ROUND_CHOICES)[self.round_combo.get()]
        val = parse_number(self.price_value.get(), 0.0) or 0.0
        if mode != "round" and self.price_value.get().strip() == "" and mode == "set":
            messagebox.showwarning(APP_TITLE, "Indiquez la valeur du prix.")
            return
        op = PriceOp(mode=mode, value=val, round_to=round_to)
        # le prix de vente est unique : on prend un champ de reference present
        col = Cols.PV_HT if self.repo.has(Cols.PV_HT) else Cols.PV_TTC
        n = 0
        for rec in recs:
            new = op.apply(rec.get(col))
            if new is None:
                continue
            rec[col] = new
            self._stage(rec, col, new)
            self._sync_price(rec, col, new)   # recopie sur l'autre champ (HT=TTC)
            n += 1
        self._refresh_all_selected(recs)
        self.status.set("Prix de vente applique a %d article(s)." % n)

    def apply_tva(self):
        recs = self._selected_recs()
        if not recs:
            return
        tva = parse_number(self.tva_value.get())
        if tva is None:
            messagebox.showwarning(APP_TITLE, "Indiquez un taux de TVA (ex : 19).")
            return
        for rec in recs:
            rec[Cols.TVA] = tva
            self._stage(rec, Cols.TVA, tva)
        self._refresh_all_selected(recs)
        self.status.set("TVA %s%% appliquee a %d article(s)." % (fmt_price(tva), len(recs)))

    def copy_ref_to_barcode(self):
        recs = self._selected_recs()
        if not recs:
            return
        for rec in recs:
            self._apply_text_to_rec(rec, Cols.CODE_BARRES,
                                    TextOp("copy_from"), source=rec.get(Cols.REF))
        self._refresh_all_selected(recs)
        self.status.set("Reference recopiee dans le code-barres pour %d article(s)." % len(recs))

    def apply_text(self, logical, op):
        recs = self._selected_recs()
        if not recs:
            return
        for rec in recs:
            self._apply_text_to_rec(rec, logical, op)
        self._refresh_all_selected(recs)
        self.status.set("Operation '%s' appliquee a %d article(s)." % (op.mode, len(recs)))

    def apply_ref_replace(self):
        if not self.repo.has(Cols.REF):
            return
        recs = self._selected_recs()
        if not recs:
            return
        find = self.ref_find.get()
        if not find:
            messagebox.showwarning(APP_TITLE, "Indiquez le texte a chercher dans la reference.")
            return
        if not self._warn_ref_rename():
            return
        op = TextOp("replace", value=self.ref_repl.get(), find=find)
        changed = 0
        for rec in recs:
            old = rec.get(Cols.REF)
            new = op.apply(old)
            if new and new != old:
                if self.repo.ref_exists(new) and new not in (r.get(Cols.REF) for r in recs):
                    messagebox.showwarning(APP_TITLE,
                                           "La reference '%s' existe deja. Operation interrompue." % new)
                    return
                rec[Cols.REF] = new
                self._stage(rec, Cols.REF, new)
                changed += 1
        self._refresh_all_selected(recs)
        self.status.set("Reference modifiee pour %d article(s)." % changed)

    # -- famille ----------------------------------------------------------
    def _famille_list(self):
        """Familles existantes en base + celles en attente de creation."""
        items = list(self.repo.list_familles())
        existing = {c for c, _ in items}
        for code, (name, _tva) in self.new_familles.items():
            if code not in existing:
                items.append((code, (name or "") + "  (nouvelle)"))
        return sorted(items, key=lambda t: str(t[0]))

    def _refresh_famille_combo(self):
        if not hasattr(self, "fam_combo"):
            return
        self._fam_values = self._famille_list()
        self.fam_combo["values"] = ["%s — %s" % (c, n) for c, n in self._fam_values]

    def _assign_famille(self, code, recs):
        for rec in recs:
            rec[Cols.FAMILLE] = code
            self._stage(rec, Cols.FAMILLE, code)
        self._refresh_all_selected(recs)

    def apply_famille_existing(self):
        recs = self._selected_recs()
        if not recs:
            return
        sel = self.fam_combo.current()
        if sel < 0:
            messagebox.showwarning(APP_TITLE, "Choisissez une famille dans la liste.")
            return
        code = self._fam_values[sel][0]
        self._assign_famille(code, recs)
        self.status.set("Famille '%s' affectee a %d article(s)." % (code, len(recs)))

    def apply_famille_new(self):
        recs = self._selected_recs()
        if not recs:
            return
        code = self.fam_new_code.get().strip()
        name = self.fam_new_name.get().strip()
        if not code or not name:
            messagebox.showwarning(APP_TITLE,
                                   "Indiquez le CODE et le NOM de la nouvelle famille.")
            return
        if self.repo.famille_exists(code):
            if not messagebox.askyesno(
                    APP_TITLE, "La famille '%s' existe deja. L'affecter quand meme ?" % code):
                return
        else:
            self.new_familles[code] = (name, None)
            self._refresh_famille_combo()
        self._assign_famille(code, recs)
        self.status.set("Famille '%s' (%s) creee et affectee a %d article(s) "
                        "— a confirmer par Enregistrer." % (code, name, len(recs)))

    def _edit_famille_inline(self, iid, rec, new_code):
        """Edition de CODEFAMILLE par double-clic : verifie l'existence, propose
        de creer la famille (code + nom) si elle n'existe pas."""
        new_code = (new_code or "").strip()
        if not new_code:
            rec[Cols.FAMILLE] = None
            self._stage(rec, Cols.FAMILLE, None)
            self._refresh_row(iid, rec)
            self._update_save_button()
            return
        if not (self.repo.famille_exists(new_code) or new_code in self.new_familles):
            name = simpledialog.askstring(
                APP_TITLE, "La famille '%s' n'existe pas.\nNom de la nouvelle famille "
                "(laisser vide pour annuler) :" % new_code, parent=self.master)
            if not name:
                return
            self.new_familles[new_code] = (name.strip(), None)
            self._refresh_famille_combo()
        rec[Cols.FAMILLE] = new_code
        self._stage(rec, Cols.FAMILLE, new_code)
        self._refresh_row(iid, rec)
        self._update_save_button()

    def _apply_text_to_rec(self, rec, logical, op, source=None):
        new = op.apply(rec.get(logical), source_value=source)
        rec[logical] = new
        self._stage(rec, logical, new)

    def _refresh_all_selected(self, recs):
        for iid in self.tree.selection():
            self._refresh_row(iid, self.row_by_iid[iid])
        self._update_save_button()

    # -- import / export --------------------------------------------------
    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Exporter la vue en CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh, delimiter=";")
            w.writerow([Cols.LABELS.get(c, c) for c in self.columns])
            for rec in self.rows:
                w.writerow([("" if rec.get(c) is None else
                             (fmt_price(rec.get(c)) if c in Cols.NUMERIC else rec.get(c)))
                            for c in self.columns])
        messagebox.showinfo(APP_TITLE, "Export termine :\n%s" % path)

    def import_file(self):
        """Importe un fichier Excel/CSV pour mettre a jour prix/code-barres par reference."""
        path = filedialog.askopenfilename(
            title="Fichier de mise a jour (Excel/CSV)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")])
        if not path:
            return
        try:
            rows = _read_tabular(path)
        except Exception as exc:                        # noqa: BLE001
            messagebox.showerror(APP_TITLE, "Lecture impossible : %s" % exc)
            return
        if not rows:
            messagebox.showwarning(APP_TITLE, "Fichier vide.")
            return
        dlg = ImportMappingDialog(self.master, rows[0].keys())
        self.master.wait_window(dlg)
        if not dlg.result:
            return
        m = dlg.result    # {"ref": col, "pv_ht": col|None, "pv_ttc": col|None, "cb": col|None}
        index = {r.get("__ref0__"): r for r in self.rows}
        # construit aussi un index global (la vue peut etre filtree)
        applied, missing = 0, 0
        updates = {}
        for row in rows:
            ref = str(row.get(m["ref"], "")).strip()
            if not ref:
                continue
            vals = {}
            if m.get("pv_ht") and row.get(m["pv_ht"]) not in (None, ""):
                vals[Cols.PV_HT] = parse_number(row[m["pv_ht"]])
            if m.get("pv_ttc") and row.get(m["pv_ttc"]) not in (None, ""):
                vals[Cols.PV_TTC] = parse_number(row[m["pv_ttc"]])
            if m.get("cb") and row.get(m["cb"]) not in (None, ""):
                vals[Cols.CODE_BARRES] = str(row[m["cb"]]).strip()
            if vals:
                updates[ref] = vals
        # Applique aux lignes visibles ; pour les autres, stage direct par ref.
        for ref, vals in updates.items():
            rec = index.get(ref)
            if rec is not None:
                for k, v in vals.items():
                    rec[k] = v
                    self._stage(rec, k, v)
                applied += 1
            else:
                # pas dans la vue courante : on programme quand meme la modif
                self.pending.setdefault(ref, {}).update(vals)
                missing += 1
        self._populate_keep_pending()
        self._update_save_button()
        messagebox.showinfo(
            APP_TITLE,
            "Import preparé : %d ligne(s) sur des articles visibles, %d sur "
            "des articles hors vue.\nVerifiez puis cliquez 'Enregistrer'."
            % (applied, missing))

    def _populate_keep_pending(self):
        self._populate()
        # re-applique le surlignage des lignes modifiees
        for iid, rec in self.row_by_iid.items():
            if rec.get("__ref0__") in self.pending:
                self._refresh_row(iid, rec)


# --------------------------------------------------------------------------- #
#  Dialogue de correspondance des colonnes pour l'import
# --------------------------------------------------------------------------- #
class ImportMappingDialog(tk.Toplevel):
    def __init__(self, master, columns):
        super().__init__(master)
        self.title("Correspondance des colonnes")
        self.result = None
        cols = ["(aucune)"] + list(columns)
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Indiquez quelle colonne du fichier correspond a chaque champ :"
                  ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.vars = {}
        rows = [("ref", "Reference (obligatoire)"), ("pv_ht", "Prix vente HT"),
                ("pv_ttc", "Prix vente TTC"), ("cb", "Code-barres")]
        from editor_logic import norm
        for i, (key, label) in enumerate(rows, start=1):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=2)
            var = tk.StringVar()
            # pre-selection automatique par nom
            guess = "(aucune)"
            for c in columns:
                nc = norm(c)
                if key == "ref" and ("ref" in nc or "code article" in nc):
                    guess = c
                elif key == "pv_ht" and "ht" in nc and ("vente" in nc or "pv" in nc or "prix" in nc):
                    guess = c
                elif key == "pv_ttc" and "ttc" in nc:
                    guess = c
                elif key == "cb" and ("barre" in nc or "ean" in nc or "gencode" in nc):
                    guess = c
            var.set(guess)
            ttk.Combobox(frm, textvariable=var, values=cols, state="readonly",
                         width=30).grid(row=i, column=1, sticky="we", pady=2)
            self.vars[key] = var
        btns = ttk.Frame(frm)
        btns.grid(row=len(rows) + 1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Importer", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="Annuler", command=self.destroy).pack(side="left")
        self.transient(master)
        self.grab_set()

    def _ok(self):
        res = {k: (v.get() if v.get() != "(aucune)" else None) for k, v in self.vars.items()}
        if not res.get("ref"):
            messagebox.showwarning(APP_TITLE, "La colonne 'Reference' est obligatoire.", parent=self)
            return
        self.result = res
        self.destroy()


# --------------------------------------------------------------------------- #
#  Lecture tabulaire Excel/CSV -> liste de dicts
# --------------------------------------------------------------------------- #
def _read_tabular(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = [str(c).strip() if c is not None else "" for c in rows[0]]
        out = []
        for r in rows[1:]:
            out.append({header[i]: r[i] for i in range(len(header))})
        return out
    # CSV
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        sample = fh.read(2048)
        fh.seek(0)
        delim = ";" if sample.count(";") >= sample.count(",") else ","
        return list(csv.DictReader(fh, delimiter=delim))


# --------------------------------------------------------------------------- #
#  Point d'entree
# --------------------------------------------------------------------------- #
def open_repository(root, args):
    """Etablit la connexion (ou le mode demo) et renvoie un repository pret."""
    if args.demo:
        return DemoRepository().connect()

    cfg_path = args.config or _default_config_path()
    try:
        cfg = load_config(cfg_path)
        config_ok = True
    except DBError as exc:
        # config.json illisible : on previent et on ouvre la fenetre de connexion
        messagebox.showerror(APP_TITLE, str(exc))
        cfg = dict(article_db.DEFAULT_CONFIG)
        config_ok = False
    if args.db:
        cfg["database"] = args.db

    # Connexion directe et silencieuse si une config valide existe deja
    # (et qu'on ne force pas le dialogue). Sinon : ecran de connexion.
    if config_ok and os.path.isfile(cfg_path) and not args.ask:
        try:
            return ArticleRepository(cfg).connect()
        except DBError:
            pass   # on bascule sur l'ecran de connexion ci-dessous

    repo, cfg = prompt_connection(root, cfg)
    if repo is not None:
        try:
            save_config(cfg_path, cfg)     # memorise les parametres qui marchent
        except OSError:
            pass
    return repo


def prompt_connection(root, cfg, allow_demo=True):
    """Affiche l'ecran de connexion en boucle. Renvoie (repo, cfg).

    repo est deja connecte (ArticleRepository ou DemoRepository), ou None si
    l'utilisateur annule.
    """
    dlg = ConnectDialog(root, cfg, allow_demo=allow_demo)
    root.wait_window(dlg)
    if dlg.demo:
        return DemoRepository().connect(), cfg
    if dlg.repo is not None:
        return dlg.repo, dlg.result
    return None, cfg


def _default_config_path():
    """Chemin de config.json, A COTE de l'exe une fois empaquete (PyInstaller).

    En .exe --onefile, __file__ pointe vers un dossier temporaire (_MEIPASS)
    efface a la fermeture : on prend alors le dossier de l'executable pour que
    config.json soit persistant et editable a cote du programme."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Editeur en masse des articles PRIME.")
    ap.add_argument("--config", help="Fichier de configuration JSON (connexion).")
    ap.add_argument("--db", help="Chemin du .FDB (surcharge la config).")
    ap.add_argument("--ask", action="store_true",
                    help="Toujours afficher la boite de connexion au demarrage.")
    ap.add_argument("--demo", action="store_true",
                    help="Mode demonstration : donnees en memoire, sans Firebird.")
    args = ap.parse_args(argv)

    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("1120x680")
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass

    repo = open_repository(root, args)
    if repo is None:
        root.destroy()
        return 0

    root._app = BulkEditorApp(root, repo, config_path=args.config)
    root.protocol("WM_DELETE_WINDOW", lambda: _on_close(root))
    root.mainloop()
    return 0


def _on_close(root):
    app = getattr(root, "_app", None)
    if app is not None and (app.pending or app.new_familles) and not messagebox.askyesno(
            APP_TITLE, "Des modifications ne sont pas enregistrees. Quitter quand meme ?"):
        return
    try:
        if app is not None:
            app.repo.close()
    finally:
        root.destroy()


if __name__ == "__main__":
    sys.exit(main())
