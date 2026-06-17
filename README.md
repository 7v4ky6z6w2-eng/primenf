# PRIME — Éditeur en masse des articles (Firebird) avec interface graphique

Outil de bureau (Python + Tkinter) pour **modifier en masse** les articles de
votre base PRIME (Firebird `.FDB`). Il cible exactement ce que vous avez
demandé :

| Champ | Colonne(s) de la table `ARTICLE` |
|-------|----------------------------------|
| **Prix de vente** | `PRIXVENTEHT` / `PRIXVENTETTC` (recalcul HT⇄TTC automatique via `TAUX_TVA`) |
| **Référence article** | `REF_ART` (= le code scanné) |
| **Code-barres** | `CODE_BARRES` (60) et `CODE_BARRE` (35) |
| **Famille** | `CODEFAMILLE` — affectation **et création** d'une famille (code + nom) |

Le schéma a été repris de votre script `import_bon_reception.py` ; l'application
**détecte automatiquement** les colonnes réelles de la base (et leur taille) au
moment de la connexion, donc elle s'adapte même si vos noms diffèrent un peu.

---

## Installation

```bash
pip install -r requirements.txt        # fdb + openpyxl
# Linux uniquement : Tkinter -> sudo apt install python3-tk
```

Il faut aussi la **bibliothèque cliente Firebird** (`fbclient.dll` sous Windows,
livrée avec le serveur/Firebird Embedded). Si elle n'est pas trouvée
automatiquement, indiquez son chemin dans `fb_client_library` (voir config).

> Python **3.11** recommandé (testé sur 3.11 et 3.12).

## Configuration

Copiez `config.example.json` en `config.json` et adaptez-le :

```json
{
  "host": "localhost",
  "port": 3050,
  "database": "C:\\PRIME\\PR22.FDB",
  "user": "SYSDBA",
  "password": "masterkey",
  "charset": "WIN1252",
  "table": "ARTICLE",
  "famille_table": "FAMILLE",
  "fb_client_library": ""
}
```

* `host` vide (`""`) = accès **local direct** au fichier (Firebird Embedded).
* `charset` : `WIN1252` comme dans votre import.

## Lancement

```bash
python bulk_article_editor.py                 # utilise ./config.json (sinon, demande la connexion)
python bulk_article_editor.py --config mon.json
python bulk_article_editor.py --db "C:\PRIME\PR22.FDB"
python bulk_article_editor.py --ask           # force la fenêtre de connexion
python bulk_article_editor.py --demo          # données fictives, SANS Firebird (pour essayer l'interface)
```

Astuce : commencez par `--demo` pour découvrir l'interface sans risque.

## Utilisation

1. **Rechercher / filtrer** : tapez une référence, un code-barres ou un libellé,
   puis *Filtrer* (recherche sur `REF_ART`, `CODE_BARRES`, `CODE_BARRE`,
   `DESIGNATION`).
2. **Sélectionner** une ou plusieurs lignes (Ctrl+clic, Maj+clic).
3. **Modifier**, au choix :
   * **Double-clic** sur une cellule (réf., code-barres, prix, famille) pour
     éditer une valeur unique.
   * **Opérations en masse** (bas de fenêtre) sur la sélection :
     * **Prix de vente** : *Fixer*, *+/- %*, *+/- montant*, ou *Arrondir* —
       sur le **HT** ou le **TTC** ; l'autre prix est recalculé via la TVA.
       Arrondis commerciaux : `.99`, `.95`, `.90`, `0,50`, `0,10`, `0,05`, entier.
     * **Code-barres** : *Fixer* une valeur, *Recopier la référence*, *Vider*.
     * **Référence** : chercher-remplacer (préfixe/suffixe possible).
     * **Famille** : *Affecter* une famille existante, ou *Créer & affecter*
       une nouvelle famille (**code + nom** obligatoires — créée dans la même
       transaction, juste avant l'affectation, pour respecter la clé étrangère).
4. Les lignes modifiées sont **surlignées en jaune** ; rien n'est écrit tant que
   vous ne cliquez pas sur **Enregistrer**.
5. **Enregistrer** affiche un récapitulatif puis valide tout en **une seule
   transaction** (`commit`). **Annuler les modifs** fait un `rollback`.

### Import / Export

* **Importer Excel/CSV…** : met à jour `PRIXVENTEHT` / `PRIXVENTETTC` /
  `CODE_BARRES` des articles **par référence** (vous mappez les colonnes du
  fichier). Pratique en complément de votre import de bons de réception.
* **Exporter CSV…** : exporte la vue courante (séparateur `;`, UTF-8 BOM, prêt
  pour Excel).

## ⚠️ À propos de la modification d'une référence (`REF_ART`)

`REF_ART` est la clé de l'article et peut être référencée par les lignes de
pièces (stock, ventes). Selon les contraintes de votre base, la renommer peut
être refusée. L'application **prévient** avant, et en cas d'erreur la
transaction est **annulée intégralement** (aucune modification partielle).
Faites une **sauvegarde** (`gbak`) avant une campagne de renommage.

## Architecture

| Fichier | Rôle |
|---------|------|
| `bulk_article_editor.py` | Interface graphique Tkinter (point d'entrée). |
| `article_db.py` | Accès Firebird (`fdb`) : connexion, introspection, lecture, écriture transactionnelle, familles. Inclut un dépôt de démo en mémoire. |
| `editor_logic.py` | Logique pure (prix, arrondis, HT⇄TTC, chercher-remplacer, validation longueur, détection des colonnes). Sans base ni interface. |
| `test_editor_logic.py` | Tests unitaires (14) de la logique et du dépôt de démo. |
| `config.example.json` | Modèle de configuration de connexion. |
| `requirements.txt` | Dépendances (`fdb`, `openpyxl`). |

## Tests

```bash
python test_editor_logic.py        # ou: python -m pytest test_editor_logic.py
```

Les calculs (prix, arrondis, HT/TTC), les opérations texte, la détection des
colonnes et le cycle modifier→commit/rollback du dépôt de démo sont couverts.
Le chemin Firebird réel (`fdb`) a été validé contre une base Firebird de test
au schéma identique (introspection, mise à jour en masse, création de famille
sous contrainte de clé étrangère, commit/rollback).

## Sécurité des données

* Tout est écrit dans **une transaction** : succès complet ou aucun changement.
* Les textes sont **tronqués proprement** à la taille réelle des colonnes
  (en octets, selon le charset), comme votre script d'import.
* **Sauvegardez** toujours la base (`gbak`) avant une grosse campagne de
  modifications.
