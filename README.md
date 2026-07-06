# PRIME — Éditeur en masse des articles (Firebird) avec interface graphique

Outil de bureau (Python + Tkinter) pour **modifier en masse** les articles de
votre base PRIME (Firebird `.FDB`). Il cible exactement ce que vous avez
demandé :

| Champ | Colonne(s) de la table `ARTICLE` |
|-------|----------------------------------|
| **Prix de vente** | `PRIXVENTEHT` **et** `PRIXVENTETTC` reçoivent la même valeur (la TVA n'est pas modifiée) |
| **Prix promo** | `PRIXHTPROMO` / `PRIXTTCPROMO` (même valeur), avec `ACTIVEPROMO` (activée automatiquement) et dates `DATEDEBPROMO` / `DATEFINPROMO` |
| **TVA** | `TAUX_TVA` (édition du taux en masse, indépendante du prix) |
| **Référence article** | `REF_ART` (= le code scanné) |
| **Codes équivalents** | table `EQUIV_CBARRES` — **plusieurs codes-barres par article** (séparés par `;` dans la grille) |
| **Famille** | `CODEFAMILLE` — affectation **et création** d'une famille (code + nom) |

**Impression d'étiquettes** (directe, sur l'imprimante de votre choix, sans PDF) —
trois modèles : **M1** 40×20 mm (haut : désignation + **prix en grand** ; bas : code-barres), **M2** 80×20 mm
(désignation + prix), **M3** 40×20 mm (prix normal barré + prix promo, ticket de remise).

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
* Les chemins Windows à **simples antislashs** (`C:\oransoft\...`) sont tolérés
  automatiquement ; vous pouvez aussi utiliser des slashs (`C:/oransoft/...`).

## Connexion (écran graphique)

Vous **n'êtes pas obligé d'éditer le JSON à la main**. Au premier lancement (ou
avec `--ask`, ou via le bouton **« Changer de base… »** de la barre d'outils),
un **écran de connexion** s'ouvre :

* bouton **Parcourir…** pour choisir le fichier `.FDB` ;
* champs hôte / port / utilisateur / mot de passe / charset / table ;
* bouton **« Tester la connexion »** qui se connecte réellement et affiche, en
  vert ✓ le nombre d'articles et de familles trouvés, ou en rouge ✗ un message
  d'erreur clair (base introuvable, identifiants, fbclient, encodage…) ;
* bouton **« Se connecter »** (ouvre l'éditeur) ou **« Mode démo »**.

Les paramètres qui fonctionnent sont mémorisés dans `config.json` pour les
prochaines fois.

## Lancement

```bash
python bulk_article_editor.py                 # utilise ./config.json (sinon, demande la connexion)
python bulk_article_editor.py --config mon.json
python bulk_article_editor.py --db "C:\PRIME\PR22.FDB"
python bulk_article_editor.py --ask           # force la fenêtre de connexion
python bulk_article_editor.py --demo          # données fictives, SANS Firebird (pour essayer l'interface)
```

Astuce : commencez par `--demo` pour découvrir l'interface sans risque.

## Créer un .exe (Windows)

Pour utiliser l'outil sur un poste **sans Python**, fabriquez un exécutable avec
PyInstaller. Le plus simple :

```powershell
py -3.11 -m pip install -r requirements.txt -r requirements-build.txt
.\build_exe.ps1
```
(ou double-cliquez `build_exe.bat`). L'exécutable est créé dans
`dist\BulkArticleEditor.exe` — un **seul fichier**, copiable sur un autre poste.

Commande équivalente, à la main :
```powershell
py -3.11 -m PyInstaller --onefile --windowed --name BulkArticleEditor bulk_article_editor.py
```

Points importants :

* **`config.json`** n'est *pas* inclus dans l'exe (il contient vos identifiants).
  Au premier lancement, l'écran de connexion s'ouvre et **enregistre**
  `config.json` **à côté de l'exe**. Vous pouvez aussi y déposer votre propre
  `config.json` (copie de `config.example.json`).
* **`fbclient.dll`** (client Firebird) **n'est pas embarqué** : il doit être
  présent sur le poste cible — c'est le cas là où PRIME/NetFact est installé.
  Sinon, indiquez son chemin dans le champ *fbclient* de l'écran de connexion.
* **32 / 64 bits** : construisez l'exe avec un Python de la **même architecture**
  que votre `fbclient.dll` (souvent **32 bits** pour Firebird/PRIME). En cas
  d'erreur « n'est pas une application Win32 valide » ou fbclient introuvable,
  réinstallez un **Python 3.11 32 bits** et reconstruisez.
* `--windowed` masque la console. Pour diagnostiquer un souci, reconstruisez
  **sans** `--windowed` afin de voir les messages.

## Utilisation

1. **Rechercher / filtrer** : tapez une référence, un code équivalent ou un
   libellé, puis *Filtrer* (recherche sur `REF_ART`, `DESIGNATION` et les
   codes de la table `EQUIV_CBARRES`).
2. **Sélectionner** une ou plusieurs lignes (Ctrl+clic, Maj+clic).
3. **Modifier**, au choix :
   * **Double-clic** sur une cellule (réf., codes équiv., prix, famille) pour
     éditer une valeur unique.
   * **Opérations en masse** (bas de fenêtre) sur la sélection :
     * **Prix de vente** : *Fixer*, *+/- %*, *+/- montant*, ou *Arrondir*.
       Arrondis commerciaux : `.99`, `.95`, `.90`, `0,50`, `0,10`, `0,05`, entier.
       Dans la base PRIME, `PRIXVENTEHT` et `PRIXVENTETTC` contiennent **le même
       prix de vente** : l'outil écrit donc la **même valeur** dans les deux et
       **ne touche jamais au taux de TVA**.
     * **TVA %** : *Fixer la TVA* applique un taux (`TAUX_TVA`) à la sélection,
       indépendamment du prix.
     * **Prix promo** : *Fixer promo* écrit le prix promotionnel (HT et TTC =
       la même valeur, comme le prix normal) **et active** la promo
       (`ACTIVEPROMO`) ; dates de début/fin facultatives (`JJ/MM/AAAA`).
       *Désactiver promo* remet `ACTIVEPROMO` à 0. La colonne *Prix promo*
       s'édite aussi au double-clic (activation automatique).
     * **Codes équivalents** : un article peut avoir **plusieurs codes-barres**
       (table `EQUIV_CBARRES`). Double-clic sur la colonne *Codes equiv.* :
       la liste complète s'édite, codes **séparés par `;`** (vider = supprimer
       tous les codes). En masse : *Ajouter* un ou des codes, *Ajouter la
       référence* comme code, ou *Vider* tous les codes de la sélection.
     * **Référence** : chercher-remplacer (préfixe/suffixe possible).
     * **Famille** : *Affecter* une famille existante, ou *Créer & affecter*
       une nouvelle famille (**code + nom** obligatoires — créée dans la même
       transaction, juste avant l'affectation, pour respecter la clé étrangère).
4. Les lignes modifiées sont **surlignées en jaune** ; rien n'est écrit tant que
   vous ne cliquez pas sur **Enregistrer**.
5. **Enregistrer** affiche un récapitulatif puis valide tout en **une seule
   transaction** (`commit`). **Annuler les modifs** fait un `rollback`.

### Import / Export

* **Importer Excel/CSV…** : met à jour `PRIXVENTEHT` / `PRIXVENTETTC` et
  **ajoute des codes équivalents** (plusieurs codes possibles dans une même
  cellule, séparés par `;`) aux articles **par référence** (vous mappez les
  colonnes du fichier). Pratique en complément de votre import de bons de
  réception.
* **Exporter CSV…** : exporte la vue courante (séparateur `;`, UTF-8 BOM, prêt
  pour Excel).

### Impression d'étiquettes

Sélectionnez des articles, puis **Imprimer étiquettes…** :

1. Choisissez le **modèle** (M1 prix en grand + désignation en haut, code-barres en bas ; M2 désignation
   + prix ; M3 prix normal barré + prix promo).
2. Choisissez l'**imprimante** (liste des imprimantes installées) et le **nombre
   de copies** par article.
3. Un **aperçu** à l'écran montre l'étiquette du 1ᵉʳ article ; cliquez
   **Imprimer** pour envoyer directement à l'imprimante (aucun fichier PDF/HTML).

Le **code-barres** est un **Code 128 auto (jeux B/C)** généré en interne (Python
pur), scannable par n'importe quelle douchette ; il encode le 1ᵉʳ code équivalent
de l'article (ou, à défaut, sa référence). Les suites de chiffres (codes EAN)
passent en jeu C : les barres sont ~2× plus larges, donc lisibles même sur une
petite étiquette imprimée à 203 dpi. Le prix promo n'apparaît sur M3 que si la
promo est **active**. Les prix sont affichés en dinars (`DA`).

> L'impression directe nécessite **Windows** avec **pywin32** (installé
> automatiquement par les scripts de build sous Windows). L'aperçu, lui, marche
> partout.

#### Réglage de l'imprimante thermique (ex. Xprinter XP‑427D)

L'application **n'impose pas** la taille du papier : elle imprime sur le format
défini dans le **pilote**. Réglez-le **une fois** (sinon les étiquettes sortent
décalées ou le code-barres s'imprime en pavé noir) :

1. *Panneau de configuration → Périphériques et imprimantes* → clic droit sur
   **Xprinter XP‑427D** → **Options d'impression**.
2. **Taille du papier / stock** : créez/choisissez **40 × 20 mm** (largeur ×
   hauteur, comme le modèle imprimé). Pour le modèle M2, **80 × 20 mm**.
3. **Type de support** : *étiquettes avec espace (gap)* ; lancez la
   **calibration du capteur** (bouton FEED maintenu, ou outil de calibration du
   pilote) pour que l'imprimante détecte le début de chaque étiquette.
4. **Orientation** : si le texte/code-barres sort tourné, basculez
   *Portrait/Paysage* ici.
5. **Vitesse / contraste (darkness)** au besoin, puis *Appliquer*.

Ensuite, dans PRIME : sélectionnez des articles → **Imprimer étiquettes…** →
choisissez l'imprimante et le nombre de copies → **Imprimer**. Le code-barres
est tracé sur une **grille de pixels** (barres nettes à 203 dpi) et le prix +
la désignation occupent la moitié haute (modèle 1).

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
| `article_db.py` | Accès Firebird (`fdb`) : connexion, introspection, lecture, écriture transactionnelle, familles, tarifs, codes équivalents. Inclut un dépôt de démo en mémoire. |
| `editor_logic.py` | Logique pure (prix, arrondis, HT⇄TTC, chercher-remplacer, validation longueur, détection des colonnes, découpage des codes, dates/booléens). Sans base ni interface. |
| `label_print.py` | Étiquettes : encodage code-barres Code 128 (Python pur), mise en page partagée (aperçu Tk + impression), impression directe Windows (pywin32). |
| `test_editor_logic.py` | Tests unitaires (22) de la logique, du code-barres, de la mise en page des étiquettes et du dépôt de démo. |
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
