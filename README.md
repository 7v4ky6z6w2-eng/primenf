# primenf — Import Bon de Réception depuis Excel fournisseur

Outil en ligne de commande qui importe **directement dans la base Firebird**
de la gestion commerciale un fichier **Excel envoyé par le fournisseur**, sous
forme d'un **Bon de réception** (type de pièce `PC_AC_B`).

Au lieu de saisir les bons de réception à la main, vous demandez au
fournisseur d'exporter sa liste d'articles en Excel, et vous l'importez en
une commande.

## Ce que fait l'outil

Pour chaque ligne du fichier Excel :

1. **Crée l'article s'il n'existe pas** (clé = `Ref. Art.`).
   - **Prix d'achat HT** = le prix indiqué sur l'Excel du fournisseur
     (colonne `Prix HT`). C'est le prix de vente du fournisseur, donc ce que
     **vous** payez l'article.
   - **Prix de vente HT** calculé automatiquement = **prix d'achat + 50 %**
     (configurable via `marge_pct`), puis **arrondi vers le haut** par paliers :
     **< 200 → au 5 supérieur** (72 → 75, 163 → 165) ; **200–999 → au 10
     supérieur** (278 → 280, 432 → 440) ; **≥ 1000 → au 50 supérieur**
     (1334 → 1350). Paliers réglables via `arrondi_paliers`. Mettez
     `prix_vente_auto: false` pour laisser le prix de vente vide.
   - **Code-barres** : c'est la **référence article** (`Ref. Art.`) qui sert
     de code scanné — comme dans le logiciel, où le champ `CODE_BARRES` reste
     vide (et porte un index unique). Scanner le code-barres retrouve donc
     l'article par sa référence, **sans ressaisie**. Si votre Excel a une
     colonne code-barres distincte, elle est utilisée ; sinon mettez
     `barcode_depuis_ref: true` pour recopier la référence dans `CODE_BARRES`.
2. **Ajoute la ligne au Bon de réception**, ce qui **alimente le stock** avec
   la quantité reçue, au prix d'achat.

Les articles qui existent déjà ne sont **pas** modifiés (ni leur prix, ni leur
code-barres) ; seule la réception (entrée en stock) est créée.

## Conformité avec le logiciel

L'outil a été **comparé à une base de production réelle** et reproduit
fidèlement le fonctionnement interne de l'application :

| Élément              | Comportement                                                              |
|----------------------|---------------------------------------------------------------------------|
| `NOPIECE` / `NOITEM` | numérotés en `MAX + 1` (comme le logiciel), sans collision ; les plages réservées (ex. `8000000` inventaire) sont ignorées, et les générateurs `NEXTPIECE`/`NEXTITEM` sont synchronisés |
| `REF_PIECE`          | attribué automatiquement par le trigger `INSERT_PIECE`                    |
| `ETAT`               | laissé **`NULL`** (comme toutes les pièces de la base réelle)             |
| `COEFF` pièce/ligne  | lus depuis la définition du type de pièce (`PC_AC_B` → `1`)               |
| Ligne (`ITEM`)       | `COEFF = 1`, `ANNULEE = 1` → entrée en stock                              |
| `PRIXHT` de la ligne | = prix d'achat → le **PUMP** (coût moyen pondéré) est correct             |
| Article              | `CODE_BARRES` et unité laissés **vides** ; famille rattachée à une famille **existante** |
| Totaux pièce         | `MONTANTHT`, `TVA`, `MONTANTTTC` recalculés                               |

> **Vérifié via la procédure `SPSTOCK` du logiciel**, sur une base réelle de
> 8105 articles : après import, la quantité en stock et le PUMP correspondent
> exactement à une réception saisie à la main, et les articles déjà présents
> sont détectés et **non recréés**.

## Installation

```bash
pip install -r requirements.txt
```

Sur le poste qui possède déjà le logiciel, le client Firebird
(`fbclient.dll`) est en principe présent. Si besoin, indiquez son chemin via
`fb_client_library` dans la configuration.

## Configuration

Copiez `config.example.json` en `config.json` et adaptez :

```jsonc
{
  "database": "C:\\PRIME\\PR22.FDB",  // chemin de votre base
  "user": "SYSDBA",
  "password": "masterkey",
  "code_tiers": "F001",          // code fournisseur (optionnel)
  "raison_sociale": "Mon Fournisseur",
  "code_depot": "",              // code dépôt (optionnel)
  "charset": "WIN1256",          // arabe + français (voir ci-dessous)
  "default_tva": 19              // TVA par défaut si absente de l'Excel
}
```

Clés utiles :

- `code_tiers` / `code_depot` : fournisseur et dépôt du bon (optionnels,
  créés automatiquement s'ils manquent et `create_missing_tiers = true`).
- `match_famille_par_intitule` : si `true` (défaut), la colonne `Famille` de
  l'Excel est rapprochée d'une **famille existante** par son **nom** (intitulé,
  ex. `SCOLAIRE`, `Tous`) — jamais par code.
- `create_missing_familles` : si `true` (défaut), lorsqu'aucune famille
  existante ne correspond au nom, **une nouvelle famille est créée par son
  nom** (l'intitulé de l'Excel), avec un **code généré automatiquement**
  (numérique, suite des codes existants) et rattachée à `default_famille`.
  Si `false`, les familles non reconnues retombent sur `default_famille`.
- `default_famille` : famille racine / de repli. **Mettez-y le code d'une
  famille existante de votre base** (souvent la racine, ex. `TOUS`) ; les
  familles créées y sont rattachées. Elle est créée si absente.
- `default_unite` : unité de base des articles créés. **Vide par défaut**,
  comme dans le logiciel (les articles n'ont pas d'unité imposée).
- `etat` : état de la pièce. **`null` par défaut**, comme toutes les pièces
  de la base.
- `barcode_depuis_ref` : `false` par défaut — la **référence** sert de code
  scanné et `CODE_BARRES` reste vide (conforme au logiciel). Mettez `true`
  si vous voulez aussi recopier la référence dans `CODE_BARRES`.
- `reserved_id_threshold` : seuil au-dessus duquel les numéros de pièce sont
  considérés comme « plages réservées » (ex. `8000000` pour les inventaires)
  et ignorés dans le calcul du prochain `NOPIECE` (défaut `1000000`).
- `prix_vente_auto` : si `true` (défaut), un **prix de vente** est calculé
  pour les **nouveaux articles** (les articles existants ne sont jamais
  modifiés). Mettez `false` pour laisser le prix de vente vide.
- `marge_pct` : marge ajoutée au prix d'achat (`50` = +50 %).
- `arrondi_paliers` : arrondi **vers le haut** par paliers, sous forme de
  liste `[seuil_max, pas]` (`null` = au-delà). Défaut
  `[[200, 5], [1000, 10], [null, 50]]` : sous 200 → au 5 supérieur,
  200–999 → au 10 supérieur, à partir de 1000 → au 50 supérieur.
- `colonne_prix` : champ Excel utilisé comme prix d'achat (défaut `prix`,
  c.-à-d. la colonne `Prix HT`).
- `charset` : page de code utilisée pour écrire le texte. La base est en
  charset `NONE` (octets bruts) et le logiciel écrit selon la page ANSI de
  Windows :
  - **`WIN1256`** (défaut) : **arabe** + français accentué en **minuscules**
    (é, è, à, ç…). C'est l'encodage réellement utilisé par le logiciel sur un
    Windows arabe (vérifié sur la base : `لوحة الحجوم` est stocké à
    l'identique). Recommandé si vos désignations contiennent de l'arabe.
  - **`WIN1252`** : français seul (Europe de l'Ouest).
  Les majuscules accentuées (É, Ç…) n'existent pas en `WIN1256` (ni dans le
  logiciel) ; elles sont alors translittérées (É → E). Tout caractère absent
  de la page de code est translittéré/neutralisé pour ne pas bloquer
  l'écriture — plus de « ???? » à la place de l'arabe si le bon charset est
  choisi.

## Colonnes attendues

L'outil détecte automatiquement la ligne d'en-tête (celle qui contient
`Ref. Art.`) et reconnaît les colonnes par leur nom (insensible aux accents
et à la casse). Colonnes reconnues :

| Champ        | En-têtes acceptés (exemples)                              | Obligatoire |
|--------------|-----------------------------------------------------------|-------------|
| Référence    | `Ref. Art.`, `Référence`, `Code article`                  | **oui**     |
| Désignation  | `Désignation`, `Libellé`                                  | non         |
| Quantité     | `QTE`, `Quantité`                                         | **oui**     |
| Prix         | `Prix HT`, `Prix vente HT`, `PU HT`                       | **oui**     |
| TVA          | `TVA`, `Taux TVA`                                         | non         |
| Famille      | `Famille`, `Rayon`, `Catégorie`                           | non         |
| Code-barres  | `Code barres`, `EAN`, `Gencode`                           | non         |

Le format exporté par le logiciel (`Liste des articles de la pièce : …`) est
pris en charge tel quel.

Utilisez le bouton **« Créer un modèle Excel »** (GUI) ou la commande
`--make-template` pour obtenir un fichier `.xlsx` prêt à remplir.

### Référence ajoutée à la désignation

Certains fournisseurs donnent une **référence** (et parfois un code-barres)
mais une désignation **générique** (juste `stylo`, sans numéro). Pour que
l'article reste identifiable et rapprochable par la suite, l'outil **ajoute la
référence à la désignation** quand elle n'y figure pas déjà — par exemple
`stylo` (réf. `70010`) devient **`stylo 70010`**.

Exception : si la référence est en réalité un **code-barres** (suite de ≥ 8
chiffres, type EAN/UPC), elle n'est **pas** recopiée dans le nom (cela
n'apporterait rien de lisible). La référence reste alors la clé de scan, et le
nom est laissé tel quel.

Option `ref_dans_designation` (défaut **activé**) — décochable dans la GUI
(*Familles & articles → « Ajouter la référence à la désignation… »*).

### Codes-barres (où ils sont stockés)

Quand l'Excel fournit une colonne **code-barres**, l'outil l'enregistre à deux
endroits, pour que le code soit à la fois **scannable** et **visible** :

- **`EQUIV_CBARRES`** — la table que le logiciel **affiche dans la fiche
  article** (le code-barres que vous voyez à l'écran) ;
- **`ARTICLE.CODE_BARRES`** — clé de scan (index UNIQUE) ; renseignée seulement
  si le code est libre.

Cela vaut pour les articles créés **et** pour les articles déjà existants qui
n'ont pas encore de code-barres (option `maj_code_barres`, défaut **activé**).
Les doublons (même code sur deux articles, ou code déjà utilisé) sont **ignorés**
sans faire échouer l'import, et signalés dans le résumé.

#### Réparer les articles déjà importés

Si vous aviez importé des codes-barres avec une version précédente, ils sont
dans `ARTICLE.CODE_BARRES` mais **pas** dans `EQUIV_CBARRES` — donc invisibles
dans la fiche article. Le bouton **« Synchroniser codes-barres »** (ou la
commande `--sync-barcodes`) les recopie vers `EQUIV_CBARRES`. L'opération est
**sans risque** (aucune suppression, rien sur le stock) et **répétable** (elle
n'ajoute jamais de doublon).

```bash
python import_bon_reception.py --config config.json --sync-barcodes
```

### Prompt OCR (photo → Excel)

Si vous numérisez un bon fournisseur papier, donnez ce prompt à Claude ou
tout autre outil de lecture d'image, en attachant la photo du bon :

```
Lis ce bon de réception fournisseur et retourne un fichier Excel avec
exactement ces colonnes dans cet ordre :
  Ref. Art. | Désignation | QTE | Prix HT | TVA | Famille | Code barres
  
Règles :
- Ref. Art. : référence article exacte du fournisseur (ou vide si absente)
- QTE : quantité reçue (nombre)
- Prix HT : prix unitaire hors taxes (nombre, sans symbole)
- TVA : taux de TVA (0, 9, 19… — sans le %)
- Famille : catégorie ou rayon (vide si absent)
- Code barres : code EAN/gencode (vide si absent)
- Une ligne par article, sans ligne de total
```

## Utilisation

```bash
# Simulation (n'écrit rien, affiche le résumé) :
python import_bon_reception.py --config config.json --excel fournisseur.xlsx --dry-run

# Import réel :
python import_bon_reception.py --config config.json --excel fournisseur.xlsx

# Options de surcharge :
python import_bon_reception.py --config config.json --excel f.xlsx \
    --db "D:\\DATA\\PR22.FDB" --code-tiers F002 --date 2026-06-15
```

Faites toujours un `--dry-run` d'abord, et **sauvegardez votre base** avant le
premier import réel. Chaque exécution crée **un nouveau** bon de réception :
relancer le même fichier crée une seconde réception (le stock serait ajouté
deux fois).

## Réparer un import fait en UTF8 (noms arabes déformés)

Si un import a été fait avec le charset **UTF8**, l'arabe a été stocké en
octets UTF-8 alors que le logiciel lit en **WIN1256** : les noms s'affichent
en lettres accentuées illisibles (ex. `ظ„ظˆط­ط©`). `reparer_encodage.py`
retrouve ces textes et les ré-écrit dans le bon encodage **sur place** —
articles, prix, stock et bons **inchangés**, aucun re-import nécessaire.

```bash
python reparer_encodage.py --config config.json            # simulation
python reparer_encodage.py --config config.json --apply    # réparation
```

Dans la GUI : bouton **« Réparer l'arabe »** (simulation → confirmation →
correction). Répare `ARTICLE.DESIGNATION` et `FAMILLE.INTITULE` ; ne touche
que les textes réellement en UTF-8 (les textes corrects ne sont pas modifiés).

## Nettoyer un import fait avec le mauvais charset

Si un import a été fait avec un mauvais charset (arabe remplacé par des
« ? »), `nettoyer_articles.py` retrouve les articles corrompus et les
supprime avec toutes leurs lignes liées (ITEM, TARIF, etc.), pour pouvoir
réimporter proprement avec `WIN1256`.

```bash
# Simulation (liste sans rien supprimer) :
python nettoyer_articles.py --config config.json

# Suppression (articles + lignes liées) :
python nettoyer_articles.py --config config.json --apply

# + supprimer les bons devenus vides et les familles « ? » sans article :
python nettoyer_articles.py --config config.json --apply --purge-pieces --purge-familles
```

Fait toujours une **simulation** d'abord et **sauvegardez la base**. Le motif
recherché (`--motif`, défaut `?`) ne supprime que les désignations qui le
contiennent — vérifiez la liste affichée avant `--apply`.

## Interface graphique (optionnelle)

`import_bon_reception_gui.py` est une interface de bureau (PySide6) qui pilote
l'outil ci-dessus : choisir le fichier Excel, la base et le fournisseur,
**Aperçu (dry-run)** puis **Import réel** avec résumé. Elle ne réimplémente
aucune logique : l'aperçu utilise `read_excel`, et l'exécution relance le
`main()` de l'outil (comportement identique à la ligne de commande).

Elle offre en plus :

- **Table éditable** : toutes les colonnes (Ref., Désignation, Qté, Prix achat HT,
  Prix vente, TVA %, Famille, Code-barres) sont modifiables avant import ; le
  Prix vente et le Total HT se recalculent automatiquement.
- **Rechercher correspondances** : rapproche chaque ligne sans référence à un
  article existant via le numéro dans la désignation (colore en vert=exact,
  jaune=retrouvé, rose=nouveau) et pré-remplit la référence. Si un rapprochement
  est **erroné** (un article retrouvé qui n'est pas le bon), **clic droit** sur
  la ligne jaune → *« Ce n'est pas le même article — créer un nouvel article »* :
  la référence fournisseur d'origine est restaurée automatiquement (un nouvel
  article sera créé), sans avoir à la ressaisir. Le clic droit permet aussi de
  revenir sur un refus.
- **MAJ Prix achat** : cochez la case d'une ligne pour mettre à jour le prix
  d'achat d'un article existant (avec alerte si l'écart dépasse le seuil configuré).
- **Charger les listes** : après un test de connexion, remplit les listes
  déroulantes Famille, Fournisseur et Dépôt depuis la base réelle.
- **N° du bon fournisseur** : champ `refdoc` stocké dans `PIECE.REFDOC`
  (permet la détection de doublons).
- **Vérification doublons** : avant un import réel, détecte automatiquement
  un bon déjà importé (même N° de bon ou même fournisseur + même total).
- **Annuler le dernier import** : après un import réel réussi, annule le bon
  en base (le stock est repris — sans suppression, natif au logiciel).
- **Créer un modèle Excel** : génère un fichier `.xlsx` prêt à remplir.
- **Synchroniser codes-barres** : rend visibles dans la fiche article les
  codes-barres des articles **déjà importés** (voir ci-dessous).
- une colonne **Prix vente** dans l'aperçu (marge + arrondi appliqués en direct) ;
- les réglages **prix de vente** (marge, arrondi `seuil:pas`) directement à l'écran ;
- un bouton **« Réparer l'arabe »** qui corrige sur place les noms déformés par
  un import UTF8 (via `reparer_encodage.py`), et un bouton **« Nettoyer ? »**
  qui supprime les articles « ? » d'un import WIN1252 (via `nettoyer_articles.py`) ;
- un **garde-fou** : si le fichier contient de l'arabe et que le charset choisi
  n'est pas WIN1256, l'aperçu l'affiche en rouge **et l'import réel demande
  confirmation** (UTF8 → déformé dans le logiciel ; WIN1252 → « ? »).

### Base de données distante (réseau)

Pour accéder à une base sur un serveur (PC serveur avec Firebird installé) :

| Champ | Valeur |
|-------|--------|
| **Hôte** | adresse IP ou nom du serveur (ex. `192.168.1.10` ou `SERVEUR-PRIME`) |
| **Port** | `3050` (défaut Firebird) |
| **Base** | chemin du `.FDB` **tel que vu par le serveur** (ex. `C:\PRIME\PR22.FDB`) ou un alias Firebird |
| **Librairie cliente** | `fbclient.dll` du même bitness que Python, sur **le poste client** |

Chaque poste client doit avoir `fbclient.dll` ; le fichier `.FDB` reste sur
le serveur. La base est accédée via le protocole TCP/IP Firebird (port 3050).

> **Arabe : choisissez le charset `WIN1256`.** En `UTF8`, l'aperçu affiche bien
> l'arabe **mais l'import sera déformé** dans le logiciel (qui lit en WIN1256).
> En `WIN1252`, l'arabe devient « ? ». `WIN1256` stocke l'arabe exactement comme
> votre logiciel (vérifié octet par octet sur une base réelle).

```bat
py -3.11 -m pip install -r requirements-gui.txt
py -3.11 import_bon_reception_gui.py
```

Empaquetage en exécutable Windows (`.exe`) avec PyInstaller :

```bat
py -3.11 -m pip install pyinstaller -r requirements-gui.txt
py -3.11 -m PyInstaller --onefile --windowed --name ImportBonReception ^
    --collect-all openpyxl --collect-all et_xmlfile --collect-all fdb ^
    --add-data "import_bon_reception.py;." ^
    --add-data "nettoyer_articles.py;." ^
    --add-data "reparer_encodage.py;." import_bon_reception_gui.py
```

> Les `--collect-all openpyxl/et_xmlfile/fdb` sont **indispensables** : ces
> modules sont chargés dynamiquement par la GUI, donc PyInstaller ne les
> embarque pas tout seul. Sans eux, l'exe affiche « module openpyxl non
> installé ».

Le plus simple : double-cliquez **`build_exe.bat`** (il utilise `py -3.11`)
ou lancez **`build_exe.ps1`** sous PowerShell.
L'exécutable est créé dans `dist\ImportBonReception.exe`.

### Distribuer l'exe sur d'autres PC

L'`.exe` (`--onefile`) embarque Python, PySide6 et les scripts : **un seul
fichier à copier**. Sur chaque PC cible :

- **Firebird requis pour la connexion** : `fdb` charge `fbclient.dll`. Les PC
  qui font tourner le logiciel de gestion l'ont déjà. Sinon, copiez
  `fbclient.dll` à côté de l'exe et indiquez son chemin dans le champ
  *« Librairie cliente Firebird »* (ou `fb_client_library` du config).
- **⚠ Même architecture (32/64 bits)** que `fbclient.dll` : un exe 64 bits ne
  peut pas charger un `fbclient.dll` 32 bits. Firebird 2.5 est souvent
  **32 bits** → construisez avec un **Python 32 bits**.
- **Réglages par PC** : la GUI mémorise ses réglages (base, charset…) par
  utilisateur. Configurez une fois par poste ; gardez **Charset = WIN1256**.
- **Affichage de l'arabe** : cela dépend de Windows sur le PC (paramètres
  régionaux système = arabe, « UTF-8 bêta » décoché), pas de l'exe.

> Gardez `import_bon_reception.py` et `nettoyer_articles.py` dans le même dossier
> que la GUI (ou bundlés via `--add-data`).
