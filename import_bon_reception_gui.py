#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface graphique pour import_bon_reception.py
================================================

GUI de bureau (PySide6) qui pilote l'outil `import_bon_reception.py`.

  1. choisir le fichier Excel du fournisseur,
  2. choisir la base Firebird (.FDB) + fournisseur + depot,
  3. APERCU (--dry-run) qui n'ecrit rien,
  4. IMPORT reel avec resume.

Conception
----------
La GUI ne reimplemente AUCUNE logique d'import :
  * l'apercu des lignes utilise `read_excel` de l'outil (hors-ligne, sans DB) ;
  * l'execution relance le programme lui-meme en mode "--run-cli", qui appelle
    le `main()` de l'outil d'origine. Le comportement est donc identique a la
    ligne de commande, et ce mecanisme fonctionne aussi bien en script qu'une
    fois empaquete en .exe (PyInstaller --onefile).

Lancement
---------
    pip install PySide6 openpyxl fdb
    python import_bon_reception_gui.py

Empaquetage .exe (sur Windows) :
    pyinstaller --onefile --windowed --name ImportBonReception \\
        --add-data "import_bon_reception.py;." import_bon_reception_gui.py
"""

import importlib.util
import json
import os
import re
import sys
import tempfile
import traceback

SCRIPT_NAME = "import_bon_reception.py"
CHARSETS = ["WIN1256", "WIN1252", "ISO8859_1", "UTF8", "NONE", "DOS850"]
ORG, APP = "PrimeOffice", "ImportBonReception"

try:
    GUI_SCRIPT = os.path.abspath(__file__)
except NameError:  # pragma: no cover
    GUI_SCRIPT = os.path.abspath(sys.argv[0])


# --------------------------------------------------------------------------- #
#  Localisation / chargement de l'outil d'origine
# --------------------------------------------------------------------------- #
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(GUI_SCRIPT)


def resolve_script_path():
    """Trouve import_bon_reception.py : dans le bundle PyInstaller (_MEIPASS)
    si gele, sinon a cote de ce fichier."""
    candidates = []
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(os.path.join(meipass, SCRIPT_NAME))
    candidates.append(os.path.join(app_dir(), SCRIPT_NAME))
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def load_tool_module(script_path):
    """Importe l'outil par chemin. Renvoie (module, erreur). L'outil fait
    `import fdb` / `import openpyxl` au chargement (sys.exit si absent) :
    on intercepte SystemExit pour ne pas tuer l'appelant."""
    if not script_path or not os.path.isfile(script_path):
        return None, "Script %s introuvable." % SCRIPT_NAME
    try:
        spec = importlib.util.spec_from_file_location("import_bon_reception", script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod, None
    except SystemExit as exc:
        return None, str(exc.code or "Dependance manquante (openpyxl/fdb).")
    except Exception as exc:  # noqa: BLE001
        return None, "%s: %s" % (type(exc).__name__, exc)


def friendly_error(raw):
    """Traduit les pannes courantes en francais lisible pour un debutant."""
    t = (raw or "").lower()
    rules = [
        (("module manquant : fdb", "no module named 'fdb'", "import fdb"),
         "Le pilote Firebird « fdb » n'est pas installe sur ce poste "
         "(pip install fdb)."),
        (("module manquant : openpyxl", "no module named 'openpyxl'"),
         "Le module « openpyxl » (lecture Excel) n'est pas installe "
         "(pip install openpyxl)."),
        (("fbclient", "client library", "load_api", "libfbclient"),
         "Librairie cliente Firebird introuvable (fbclient.dll). Indiquez son "
         "chemin dans les options avancees."),
        (("password", "user name", "login", "-902", "incorrect user"),
         "Identifiant ou mot de passe incorrect."),
        (("i/o error", "no such file", "cannot open", "unable to open",
          "error while trying to open file", "not found"),
         "Base de donnees introuvable : verifiez le chemin du fichier .FDB."),
        (("unavailable database", "connection refused", "rejected", "network",
          "failed to establish"),
         "Serveur Firebird injoignable : verifiez l'hote / le port, ou laissez "
         "l'hote vide pour un acces local au fichier."),
        (("malformed string", "transliteration", "charset"),
         "Probleme d'encodage : essayez un autre charset (WIN1252 / NONE)."),
        (("excel introuvable", "fichier excel"),
         "Fichier Excel introuvable."),
        (("ligne d'en-tete", "ref. art", "colonne obligatoire", "colonne de prix"),
         "Format Excel non reconnu : il faut une colonne « Ref. Art. », une "
         "quantite et un prix."),
    ]
    for keys, msg in rules:
        if any(k in t for k in keys):
            return msg
    s = (raw or "").strip()
    if not s:
        return "Erreur inconnue."
    return s.splitlines()[-1][:300]


# --------------------------------------------------------------------------- #
#  Mode "--run-cli" : lance le main() de l'outil dans ce meme binaire
# --------------------------------------------------------------------------- #
def run_cli(argv):
    """Execute import_bon_reception.main() avec argv. Utilise par la GUI via
    un sous-processus (fonctionne en script ET en .exe gele)."""
    script = resolve_script_path()
    mod, err = load_tool_module(script)
    if not mod:
        print(err)
        return 1
    sys.argv = ["import_bon_reception"] + list(argv)
    try:
        mod.main()
        return 0
    except SystemExit as exc:
        code = exc.code
        if isinstance(code, str):
            print(code)
            return 1
        return int(code) if code else 0
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        return 1


# --------------------------------------------------------------------------- #
#  GUI
# --------------------------------------------------------------------------- #
def _window_class():
    from PySide6.QtCore import Qt, QProcess, QSettings
    from PySide6.QtGui import QFont, QTextCursor
    from PySide6.QtWidgets import (
        QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
        QCheckBox, QSpinBox, QDoubleSpinBox, QComboBox, QPlainTextEdit,
        QTableWidget, QTableWidgetItem, QGroupBox, QFormLayout, QGridLayout,
        QVBoxLayout, QHBoxLayout, QMessageBox, QTabWidget, QSizePolicy,
        QHeaderView, QAbstractItemView, QProgressBar,
    )

    class MainWindow(QWidget):
        FIELD_KEYS = ("database", "code_tiers", "raison_sociale", "code_depot",
                      "host", "port", "user", "charset", "fb_client_library",
                      "code_type_piece", "default_famille", "default_famille_intitule",
                      "default_tva", "default_unite", "default_unite_intitule",
                      "colonne_prix", "reserved_id_threshold")

        def __init__(self):
            super().__init__()
            self.setWindowTitle("Import Bon de reception — Prime Office")
            self.resize(1060, 800)

            self.script_path = resolve_script_path()
            self.tool_mod, self.tool_err = load_tool_module(self.script_path)
            self.proc = None
            self.tmp_config_path = None
            self.current_mode = None

            self._create_fields()
            self._build_ui()
            self._apply_defaults_from_tool()
            self._load_settings()
            self._refresh_script_banner()

        # ---- creation des widgets de saisie (noms stables) -------------- #
        def _create_fields(self):
            self.excel_edit = QLineEdit(readOnly=True)
            self.excel_edit.setPlaceholderText("Fichier .xlsx du fournisseur…")

            self.f_database = QLineEdit()
            self.f_database.setPlaceholderText("C:\\PRIME\\PR22.FDB")
            self.f_code_tiers = QLineEdit()
            self.f_code_tiers.setPlaceholderText("ex. F001 — vide = aucun")
            self.f_raison = QLineEdit()
            self.f_raison.setPlaceholderText("Nom du fournisseur (si creation)")
            self.f_code_depot = QLineEdit()
            self.f_code_depot.setPlaceholderText("Code depot — vide = aucun")

            self.f_host = QLineEdit()
            self.f_host.setPlaceholderText("localhost — vide = acces local au fichier")
            self.f_port = QSpinBox(); self.f_port.setRange(0, 65535); self.f_port.setValue(3050)
            self.f_user = QLineEdit()
            self.f_password = QLineEdit(); self.f_password.setEchoMode(QLineEdit.EchoMode.Password)
            self.f_charset = QComboBox(); self.f_charset.addItems(CHARSETS); self.f_charset.setEditable(True)
            self.f_fbclient = QLineEdit()
            self.f_fbclient.setPlaceholderText("Chemin de fbclient.dll (optionnel)")

            self.f_code_type_piece = QLineEdit("PC_AC_B")
            self.f_date = QLineEdit(); self.f_date.setPlaceholderText("AAAA-MM-JJ — vide = aujourd'hui")
            self.f_create_tiers = QCheckBox("Creer le fournisseur / depot s'ils manquent")
            self.f_create_tiers.setChecked(True)

            self.f_default_famille = QLineEdit("TOUS")
            self.f_default_famille_int = QLineEdit("Tous")
            self.f_match_famille = QCheckBox("Rapprocher la colonne « Famille » d'une famille existante (par nom)")
            self.f_match_famille.setChecked(True)
            self.f_create_familles = QCheckBox("Creer la famille (par son nom) si aucune ne correspond")
            self.f_create_familles.setChecked(True)
            self.f_default_tva = QDoubleSpinBox(); self.f_default_tva.setRange(0, 100)
            self.f_default_tva.setDecimals(2); self.f_default_tva.setValue(19); self.f_default_tva.setSuffix(" %")
            self.f_calc_ttc = QCheckBox("Renseigner aussi le prix d'achat TTC")
            self.f_calc_ttc.setChecked(True)
            self.f_colonne_prix = QComboBox(); self.f_colonne_prix.addItems(["prix"]); self.f_colonne_prix.setEditable(True)
            self.f_default_unite = QLineEdit(); self.f_default_unite.setPlaceholderText("Vide = aucune unite")
            self.f_default_unite_int = QLineEdit("Unite")
            self.f_barcode_ref = QCheckBox("Recopier la reference dans CODE_BARRES")
            self.f_reserved = QSpinBox(); self.f_reserved.setRange(0, 2_000_000_000); self.f_reserved.setValue(1_000_000)

        # ---- assemblage ------------------------------------------------- #
        def _build_ui(self):
            root = QVBoxLayout(self)
            root.setContentsMargins(14, 12, 14, 12); root.setSpacing(10)

            title = QLabel("Import d'un Bon de reception depuis un Excel fournisseur")
            tf = QFont(); tf.setPointSize(14); tf.setBold(True); title.setFont(tf)
            root.addWidget(title)

            self.script_banner = QLabel(wordWrap=True)
            self.script_banner.setTextFormat(Qt.TextFormat.RichText)
            self.script_banner.linkActivated.connect(self.locate_script)
            root.addWidget(self.script_banner)

            # --- panneau ESSENTIEL ---
            ess = QGroupBox("L'essentiel")
            g = QGridLayout(ess); g.setColumnStretch(1, 1)

            g.addWidget(QLabel("<b>Fichier Excel</b>"), 0, 0, Qt.AlignmentFlag.AlignRight)
            g.addWidget(self.excel_edit, 0, 1)
            b1 = QPushButton("Parcourir…"); b1.clicked.connect(self.pick_excel)
            g.addWidget(b1, 0, 2)

            g.addWidget(QLabel("<b>Base Firebird (.FDB)</b>"), 1, 0, Qt.AlignmentFlag.AlignRight)
            g.addWidget(self.f_database, 1, 1)
            b2 = QPushButton("Parcourir…"); b2.clicked.connect(self.pick_fdb)
            g.addWidget(b2, 1, 2)

            test_row = QHBoxLayout()
            self.btn_test = QPushButton("Tester la connexion"); self.btn_test.clicked.connect(self.test_connection)
            self.conn_result = QLabel("")
            self.conn_result.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            test_row.addWidget(self.btn_test); test_row.addWidget(self.conn_result, 1)
            tw = QWidget(); tw.setLayout(test_row)
            g.addWidget(tw, 2, 1, 1, 2)

            g.addWidget(QLabel("Fournisseur"), 3, 0, Qt.AlignmentFlag.AlignRight)
            sup = QHBoxLayout()
            sup.addWidget(QLabel("code")); sup.addWidget(self.f_code_tiers)
            sup.addWidget(QLabel("nom")); sup.addWidget(self.f_raison, 1)
            sw = QWidget(); sw.setLayout(sup)
            g.addWidget(sw, 3, 1, 1, 2)

            g.addWidget(QLabel("Depot"), 4, 0, Qt.AlignmentFlag.AlignRight)
            g.addWidget(self.f_code_depot, 4, 1, 1, 2)
            root.addWidget(ess)

            # --- bouton options avancees ---
            self.adv_btn = QPushButton("▸ Options avancees")
            self.adv_btn.setCheckable(True)
            self.adv_btn.setStyleSheet("QPushButton { text-align:left; border:none; color:#235; }")
            self.adv_btn.toggled.connect(self._toggle_advanced)
            root.addWidget(self.adv_btn)

            self.adv = QTabWidget(); self.adv.setVisible(False)
            self.adv.addTab(self._tab_connexion(), "Connexion")
            self.adv.addTab(self._tab_bon(), "Bon de reception")
            self.adv.addTab(self._tab_articles(), "Familles & articles")
            root.addWidget(self.adv)

            # --- actions ---
            actions = QHBoxLayout()
            self.btn_preview = QPushButton("Apercu (dry-run)")
            self.btn_preview.setToolTip("Simulation : n'ecrit rien (rollback).")
            self.btn_preview.clicked.connect(lambda: self.run_import(True))
            self.btn_import = QPushButton("Importer reellement")
            self.btn_import.clicked.connect(lambda: self.run_import(False))
            bf = QFont(); bf.setBold(True)
            for b in (self.btn_preview, self.btn_import):
                b.setFont(bf); b.setMinimumHeight(38)
            self.btn_import.setStyleSheet("QPushButton { background:#7a1f1f; color:white; }")
            self.btn_cancel = QPushButton("Arreter"); self.btn_cancel.setEnabled(False)
            self.btn_cancel.clicked.connect(self.cancel_run)
            actions.addWidget(self.btn_preview); actions.addWidget(self.btn_import)
            actions.addStretch(1); actions.addWidget(self.btn_cancel)
            root.addLayout(actions)

            self.busy = QProgressBar(); self.busy.setRange(0, 1); self.busy.setValue(0)
            self.busy.setTextVisible(False); self.busy.setFixedHeight(6)
            root.addWidget(self.busy)

            # --- sorties ---
            out = QTabWidget()
            self.preview_table = QTableWidget(0, 7)
            self.preview_table.setHorizontalHeaderLabels(
                ["Ref. Art.", "Designation", "Qte", "Prix HT", "TVA %", "Famille", "Total HT"])
            self.preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            self.preview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            self.preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            pv = QWidget(); pvl = QVBoxLayout(pv); pvl.setContentsMargins(0, 0, 0, 0)
            self.preview_count = QLabel("Aucun fichier charge.")
            pvl.addWidget(self.preview_count); pvl.addWidget(self.preview_table, 1)
            out.addTab(pv, "Apercu des lignes")

            out.addTab(self._build_summary_widget(), "Resume")

            self.log = QPlainTextEdit(readOnly=True)
            self.log.setFont(QFont("Menlo, Consolas, monospace"))
            out.addTab(self.log, "Journal")
            out.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            root.addWidget(out, 1)

            self.status = QLabel("Pret."); self.status.setStyleSheet("color:#555;")
            root.addWidget(self.status)

        def _toggle_advanced(self, on):
            self.adv.setVisible(on)
            self.adv_btn.setText(("▾ " if on else "▸ ") + "Options avancees")

        def _tab_connexion(self):
            w = QWidget(); f = QFormLayout(w)
            f.addRow("Hote", self.f_host)
            f.addRow("Port", self.f_port)
            f.addRow("Utilisateur", self.f_user)
            f.addRow("Mot de passe", self.f_password)
            f.addRow("Charset", self.f_charset)
            rc = QHBoxLayout(); rc.addWidget(self.f_fbclient, 1)
            bb = QPushButton("…"); bb.setFixedWidth(34); bb.clicked.connect(self.pick_fbclient)
            rc.addWidget(bb); rcw = QWidget(); rcw.setLayout(rc)
            f.addRow("Librairie cliente FB", rcw)
            return w

        def _tab_bon(self):
            w = QWidget(); f = QFormLayout(w)
            f.addRow("Type de piece", self.f_code_type_piece)
            f.addRow("Date du bon", self.f_date)
            f.addRow("", self.f_create_tiers)
            return w

        def _tab_articles(self):
            w = QWidget(); f = QFormLayout(w)
            f.addRow("Famille par defaut — code", self.f_default_famille)
            f.addRow("Famille par defaut — nom", self.f_default_famille_int)
            f.addRow("", self.f_match_famille)
            f.addRow("", self.f_create_familles)
            f.addRow("TVA par defaut", self.f_default_tva)
            f.addRow("", self.f_calc_ttc)
            f.addRow("Colonne prix", self.f_colonne_prix)
            f.addRow("Unite par defaut — code", self.f_default_unite)
            f.addRow("Unite par defaut — nom", self.f_default_unite_int)
            f.addRow("", self.f_barcode_ref)
            f.addRow("Seuil ID reserves", self.f_reserved)
            return w

        def _build_summary_widget(self):
            w = QWidget(); grid = QGridLayout(w); grid.setColumnStretch(1, 1)
            self.summary_state = QLabel("—")
            sf = QFont(); sf.setPointSize(13); sf.setBold(True); self.summary_state.setFont(sf)
            self.summary_state.setWordWrap(True)
            grid.addWidget(self.summary_state, 0, 0, 1, 2)
            self.sum_fields = {}
            rows = [("piece", "Bon de reception"), ("tiers", "Fournisseur / depot"),
                    ("lus", "Lignes lues"), ("crees", "Articles crees"),
                    ("existants", "Articles existants"), ("ht", "Total HT"),
                    ("tva", "Total TVA"), ("ttc", "Total TTC")]
            for i, (k, lab) in enumerate(rows, start=1):
                grid.addWidget(QLabel("<b>%s</b>" % lab), i, 0,
                               Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
                v = QLabel("—"); v.setWordWrap(True)
                v.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                self.sum_fields[k] = v; grid.addWidget(v, i, 1)
            grid.setRowStretch(len(rows) + 1, 1)
            return w

        # ---- defaults / settings --------------------------------------- #
        def _apply_defaults_from_tool(self):
            if not self.tool_mod:
                return
            dc = getattr(self.tool_mod, "DEFAULT_CONFIG", {})
            if dc.get("database"):
                self.f_database.setText(str(dc["database"]).replace("\\\\", "\\"))
            self.f_user.setText(dc.get("user", "SYSDBA"))
            self.f_password.setText(dc.get("password", "masterkey"))
            cs = dc.get("charset", "WIN1252")
            if cs and self.f_charset.findText(cs) < 0:
                self.f_charset.insertItem(0, cs)
            self.f_charset.setCurrentText(cs)

        def _settings(self):
            return QSettings(ORG, APP)

        def _load_settings(self):
            s = self._settings()
            if s.value("database"):
                cfg = {}
                for k in self.FIELD_KEYS:
                    val = s.value(k)
                    if val is not None and val != "":
                        cfg[k] = val
                for b in ("create_missing_tiers", "match_famille_par_intitule",
                          "create_missing_familles", "calc_prix_achat_ttc",
                          "barcode_depuis_ref"):
                    v = s.value(b)
                    if v is not None:
                        cfg[b] = str(v).lower() in ("true", "1")
                self._apply_config_dict(cfg)
            last_excel = s.value("last_excel")
            if last_excel and os.path.isfile(last_excel):
                self.excel_edit.setText(last_excel)
                self.load_preview(last_excel)

        def _save_settings(self):
            s = self._settings(); cfg = self._build_config_dict()
            for k in self.FIELD_KEYS:
                s.setValue(k, cfg.get(k, ""))
            for b in ("create_missing_tiers", "match_famille_par_intitule",
                      "create_missing_familles", "calc_prix_achat_ttc",
                      "barcode_depuis_ref"):
                s.setValue(b, bool(cfg.get(b)))
            if self.excel_edit.text().strip():
                s.setValue("last_excel", self.excel_edit.text().strip())

        def closeEvent(self, ev):
            try:
                self._save_settings()
            finally:
                super().closeEvent(ev)

        def _refresh_script_banner(self):
            if self.script_path:
                msg = "Outil detecte : <code>%s</code>" % self.script_path
                if self.tool_err:
                    msg += ("<br><span style='color:#a15c00;'>Apercu hors-ligne indisponible "
                            "(%s).</span>" % friendly_error(self.tool_err))
                self.script_banner.setStyleSheet(
                    "color:#235; background:#eef3fb; padding:6px; border-radius:4px;")
            else:
                msg = ("<b>Outil introuvable.</b> Placez <code>%s</code> dans le meme dossier, "
                       "ou <a href='#'>cliquez ici pour le localiser</a>." % SCRIPT_NAME)
                self.script_banner.setStyleSheet(
                    "color:#7a1f1f; background:#fbeeee; padding:6px; border-radius:4px;")
            self.script_banner.setText(msg)

        # ---- selecteurs ------------------------------------------------- #
        def locate_script(self, *_):
            path, _f = QFileDialog.getOpenFileName(self, "Localiser %s" % SCRIPT_NAME, "", "Python (*.py)")
            if path:
                self.script_path = path
                self.tool_mod, self.tool_err = load_tool_module(path)
                self._apply_defaults_from_tool(); self._refresh_script_banner()

        def pick_excel(self):
            path, _f = QFileDialog.getOpenFileName(self, "Choisir l'Excel", "", "Excel (*.xlsx *.xlsm)")
            if path:
                self.excel_edit.setText(path); self.load_preview(path)

        def pick_fdb(self):
            path, _f = QFileDialog.getOpenFileName(self, "Choisir la base", "", "Firebird (*.fdb *.FDB);;Tous (*)")
            if path:
                self.f_database.setText(path)

        def pick_fbclient(self):
            path, _f = QFileDialog.getOpenFileName(self, "Choisir fbclient", "", "Librairies (*.dll *.so *.dylib);;Tous (*)")
            if path:
                self.f_fbclient.setText(path)

        # ---- apercu hors-ligne ----------------------------------------- #
        def load_preview(self, path):
            self.preview_table.setRowCount(0)
            if not self.tool_mod:
                self.preview_count.setText(
                    "Apercu hors-ligne indisponible. Utilisez « Apercu (dry-run) ».")
                return
            try:
                base = json.loads(json.dumps(getattr(self.tool_mod, "DEFAULT_CONFIG", {})))
                base.update(self._build_config_dict())
                lines = self.tool_mod.read_excel(path, base)
            except Exception as exc:  # noqa: BLE001
                self.preview_count.setText("Lecture impossible : %s" % friendly_error(str(exc)))
                return
            total = 0.0
            self.preview_table.setRowCount(len(lines))
            for r, ln in enumerate(lines):
                ht = ln["qte"] * ln["prix"]; total += ht
                cells = [ln["ref_art"], ln["designation"], self._fmt(ln["qte"]),
                         self._fmt(ln["prix"]), self._fmt(ln["tva"]),
                         ln["famille"] or "(defaut)", self._fmt(ht)]
                for c, txt in enumerate(cells):
                    it = QTableWidgetItem(txt)
                    if c in (2, 3, 4, 6):
                        it.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    self.preview_table.setItem(r, c, it)
            self.preview_count.setText(
                "%d ligne(s)  ·  Total HT indicatif : %s   (comptes crees/existants "
                "apres un apercu dry-run)" % (len(lines), self._fmt(round(total, 2))))

        @staticmethod
        def _fmt(v):
            try:
                f = float(v)
            except (TypeError, ValueError):
                return str(v)
            return str(int(f)) if f == int(f) else ("%.2f" % f)

        # ---- config <-> formulaire ------------------------------------- #
        def _build_config_dict(self):
            return {
                "host": self.f_host.text().strip(),
                "port": int(self.f_port.value()),
                "database": self.f_database.text().strip(),
                "user": self.f_user.text().strip(),
                "password": self.f_password.text(),
                "charset": self.f_charset.currentText().strip(),
                "fb_client_library": self.f_fbclient.text().strip() or None,
                "code_type_piece": self.f_code_type_piece.text().strip() or "PC_AC_B",
                "etat": None,
                "code_tiers": self.f_code_tiers.text().strip(),
                "raison_sociale": self.f_raison.text().strip(),
                "code_depot": self.f_code_depot.text().strip(),
                "create_missing_tiers": self.f_create_tiers.isChecked(),
                "default_famille": self.f_default_famille.text().strip() or "TOUS",
                "default_famille_intitule": self.f_default_famille_int.text().strip() or "Tous",
                "default_unite": self.f_default_unite.text().strip(),
                "default_unite_intitule": self.f_default_unite_int.text().strip() or "Unite",
                "default_tva": float(self.f_default_tva.value()),
                "match_famille_par_intitule": self.f_match_famille.isChecked(),
                "create_missing_familles": self.f_create_familles.isChecked(),
                "calc_prix_achat_ttc": self.f_calc_ttc.isChecked(),
                "reserved_id_threshold": int(self.f_reserved.value()),
                "barcode_depuis_ref": self.f_barcode_ref.isChecked(),
                "colonne_prix": self.f_colonne_prix.currentText().strip() or "prix",
            }

        def _apply_config_dict(self, cfg):
            def s(key, widget):
                if key in cfg and cfg[key] is not None:
                    widget.setText(str(cfg[key]))
            s("database", self.f_database); s("host", self.f_host)
            if "port" in cfg:
                try: self.f_port.setValue(int(cfg["port"]))
                except (TypeError, ValueError): pass
            s("user", self.f_user); s("password", self.f_password)
            if "charset" in cfg: self.f_charset.setCurrentText(str(cfg["charset"]))
            if "fb_client_library" in cfg:
                self.f_fbclient.setText(str(cfg.get("fb_client_library") or ""))
            s("code_type_piece", self.f_code_type_piece)
            s("code_tiers", self.f_code_tiers); s("raison_sociale", self.f_raison)
            s("code_depot", self.f_code_depot)
            if "create_missing_tiers" in cfg: self.f_create_tiers.setChecked(bool(cfg["create_missing_tiers"]))
            s("default_famille", self.f_default_famille)
            s("default_famille_intitule", self.f_default_famille_int)
            if "default_unite" in cfg:
                self.f_default_unite.setText(str(cfg.get("default_unite") or ""))
            s("default_unite_intitule", self.f_default_unite_int)
            if "default_tva" in cfg:
                try: self.f_default_tva.setValue(float(cfg["default_tva"]))
                except (TypeError, ValueError): pass
            if "match_famille_par_intitule" in cfg: self.f_match_famille.setChecked(bool(cfg["match_famille_par_intitule"]))
            if "create_missing_familles" in cfg: self.f_create_familles.setChecked(bool(cfg["create_missing_familles"]))
            if "calc_prix_achat_ttc" in cfg: self.f_calc_ttc.setChecked(bool(cfg["calc_prix_achat_ttc"]))
            if "reserved_id_threshold" in cfg:
                try: self.f_reserved.setValue(int(cfg["reserved_id_threshold"]))
                except (TypeError, ValueError): pass
            if "barcode_depuis_ref" in cfg: self.f_barcode_ref.setChecked(bool(cfg["barcode_depuis_ref"]))
            if "colonne_prix" in cfg: self.f_colonne_prix.setCurrentText(str(cfg["colonne_prix"]))

        # ---- test connexion -------------------------------------------- #
        def test_connection(self):
            cfg = self._build_config_dict()
            if not cfg["database"]:
                self._set_conn(False, "Renseignez le chemin de la base."); return
            if not self.tool_mod:
                self._set_conn(False, friendly_error(self.tool_err)); return
            self._set_conn(None, "Connexion en cours…")
            QApplication.processEvents()
            try:
                con = self.tool_mod.connect(cfg)
                con.close()
                self._set_conn(True, "Connexion reussie.")
            except SystemExit as exc:
                self._set_conn(False, friendly_error(str(exc.code)))
            except Exception as exc:  # noqa: BLE001
                self._set_conn(False, friendly_error(str(exc)))

        def _set_conn(self, ok, msg):
            color = {True: "#1f7a33", False: "#7a1f1f", None: "#555"}[ok]
            mark = {True: "✓ ", False: "✗ ", None: ""}[ok]
            self.conn_result.setStyleSheet("color:%s;" % color)
            self.conn_result.setText(mark + msg)

        # ---- execution -------------------------------------------------- #
        def _validate(self):
            p = []
            if not self.script_path:
                p.append("L'outil import_bon_reception.py est introuvable.")
            ex = self.excel_edit.text().strip()
            if not ex:
                p.append("Choisissez un fichier Excel.")
            elif not os.path.isfile(ex):
                p.append("Le fichier Excel n'existe pas.")
            if not self.f_database.text().strip():
                p.append("Renseignez le chemin de la base Firebird (.FDB).")
            d = self.f_date.text().strip()
            if d and not re.match(r"^\d{4}-\d{2}-\d{2}$", d):
                p.append("La date doit etre au format AAAA-MM-JJ.")
            return p

        def run_import(self, dry_run):
            if self.proc is not None:
                return
            probs = self._validate()
            if probs:
                QMessageBox.warning(self, "A corriger", "\n".join("• " + x for x in probs)); return
            if not dry_run:
                ok = QMessageBox.question(
                    self, "Confirmer l'import reel",
                    "Cet import ECRIT en base (commit) et alimente le stock.\n\n"
                    "• Avez-vous fait un APERCU au prealable ?\n"
                    "• Avez-vous une SAUVEGARDE recente de la base ?\n\n"
                    "Relancer le meme fichier creerait un SECOND bon (stock double).\n\n"
                    "Lancer l'import reel ?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No)
                if ok != QMessageBox.StandardButton.Yes:
                    return
            try:
                fd, self.tmp_config_path = tempfile.mkstemp(suffix=".json", prefix="primenf_cfg_")
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    json.dump(self._build_config_dict(), fh, indent=2, ensure_ascii=False)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, "Erreur", "Config temporaire : %s" % exc); return

            cli = ["--run-cli", "--config", self.tmp_config_path, "--excel", self.excel_edit.text().strip()]
            d = self.f_date.text().strip()
            if d:
                cli += ["--date", d]
            if dry_run:
                cli.append("--dry-run")
            if getattr(sys, "frozen", False):
                program, args = sys.executable, cli
            else:
                program, args = sys.executable, [GUI_SCRIPT] + cli

            self.current_mode = "dry" if dry_run else "real"
            self.log.clear()
            self._append_log("$ %s\n" % " ".join(self._q(a) for a in [program] + args))
            self._reset_summary(dry_run)

            self.proc = QProcess(self)
            self.proc.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
            self.proc.readyReadStandardOutput.connect(self._on_output)
            self.proc.finished.connect(self._on_finished)
            self.proc.errorOccurred.connect(self._on_proc_error)
            self._set_running(True)
            self.status.setText("Apercu en cours…" if dry_run else "Import reel en cours…")
            self.proc.setProgram(program); self.proc.setArguments(args); self.proc.start()

        @staticmethod
        def _q(a):
            return '"%s"' % a if " " in a else a

        def cancel_run(self):
            if self.proc is not None:
                self.proc.kill(); self.status.setText("Arret demande…")

        def _on_proc_error(self, _e):
            self._append_log("\n[ERREUR] Lancement du processus impossible.\n")

        def _on_output(self):
            if self.proc is None:
                return
            self._append_log(bytes(self.proc.readAllStandardOutput()).decode("utf-8", "replace"))

        def _append_log(self, text):
            self.log.moveCursor(QTextCursor.MoveOperation.End)
            self.log.insertPlainText(text)
            self.log.moveCursor(QTextCursor.MoveOperation.End)

        def _on_finished(self, code, _st):
            self._parse_summary(self.log.toPlainText(), code)
            if self.tmp_config_path and os.path.isfile(self.tmp_config_path):
                try: os.remove(self.tmp_config_path)
                except OSError: pass
            self.tmp_config_path = None; self.proc = None
            self._set_running(False)
            self.status.setText("Termine (code 0)." if code == 0
                                else "Termine avec erreurs (code %s) — voir Journal." % code)

        def _set_running(self, running):
            self.busy.setRange(0, 0) if running else self.busy.setRange(0, 1)
            if not running:
                self.busy.setValue(0)
            for b in (self.btn_preview, self.btn_import, self.btn_test, self.adv):
                b.setEnabled(not running)
            self.btn_cancel.setEnabled(running)

        def _reset_summary(self, dry):
            self.summary_state.setText("Apercu en cours…" if dry else "Import en cours…")
            self.summary_state.setStyleSheet("color:#555;")
            for v in self.sum_fields.values():
                v.setText("—")

        def _parse_summary(self, text, code):
            def grab(pat):
                m = re.search(pat, text); return m.group(1).strip() if m else None
            nopiece = grab(r"NOPIECE=(\S+)")
            ref_piece = grab(r"REF_PIECE=(\S+)")
            type_piece = grab(r"\(type\s+([^)]+)\)")
            four = grab(r"Fournisseur\s*:\s*(.+?)\s{2,}Depot")
            depot = grab(r"Depot\s*:\s*(.+)")
            for key, pat in (("lus", r"Lignes lues dans l'Excel\s*:\s*(\d+)"),
                             ("crees", r"Articles crees\s*:\s*(\d+)"),
                             ("existants", r"Articles existants\s*:\s*(\d+)"),
                             ("ht", r"Total HT\s*:\s*([\d.,]+)"),
                             ("tva", r"Total TVA\s*:\s*([\d.,]+)"),
                             ("ttc", r"Total TTC\s*:\s*([\d.,]+)")):
                val = grab(pat)
                if val is not None:
                    self.sum_fields[key].setText(val)
            if nopiece:
                self.sum_fields["piece"].setText(
                    "NOPIECE %s · REF_PIECE %s · type %s" % (nopiece, ref_piece or "?", type_piece or "?"))
            self.sum_fields["tiers"].setText("%s  /  %s" % (four or "(aucun)", depot or "(aucun)"))

            committed = "Import termine" in text and "commit" in text.lower()
            rolled = "DRY-RUN" in text or "rollback" in text.lower()
            if code != 0:
                self.summary_state.setText("Echec — aucune ecriture. %s" % friendly_error(text))
                self.summary_state.setStyleSheet("color:#7a1f1f;")
            elif self.current_mode == "real" and committed:
                self.summary_state.setText("Import enregistre (commit)")
                self.summary_state.setStyleSheet("color:#1f7a33;")
            elif rolled:
                self.summary_state.setText("Apercu (dry-run) — rien n'a ete ecrit (rollback)")
                self.summary_state.setStyleSheet("color:#235;")
            else:
                self.summary_state.setText("Termine"); self.summary_state.setStyleSheet("color:#555;")

    return MainWindow


def build_gui():
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv)
    win = _window_class()()
    win.show()
    return app.exec()


def main():
    return build_gui()


if __name__ == "__main__":
    if "--run-cli" in sys.argv:
        rest = [a for a in sys.argv[1:] if a != "--run-cli"]
        sys.exit(run_cli(rest))
    sys.exit(main())
