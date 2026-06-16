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
     **< 200 → au 5 supérieur** (72 → 75, 163 → 165) ; **≥ 200 → au 10
     supérieur** (278 → 280, 432 → 440). Paliers réglables via
     `arrondi_paliers`. Mettez `prix_vente_auto: false` pour laisser le prix
     de vente vide et le fixer vous-même.
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
  `[[200, 5], [null, 10]]` : sous 200 → au 5 supérieur, à partir de 200 → au
  10 supérieur. Exemple pour arrondir les milliers au 50 supérieur :
  `[[200, 5], [1000, 10], [null, 50]]`.
- `colonne_prix` : champ Excel utilisé comme prix d'achat (défaut `prix`,
  c.-à-d. la colonne `Prix HT`).

## Format du fichier Excel

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
