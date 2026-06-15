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
   - **Prix de vente** laissé **VIDE** (`NULL`) — vous fixez vous-même votre
     prix de vente ensuite dans le logiciel.
   - **Code-barres** = la référence article (`Ref. Art.`), ou la colonne
     code-barres de l'Excel si elle existe. Ainsi le scan fonctionne sans
     ressaisie.
2. **Ajoute la ligne au Bon de réception**, ce qui **alimente le stock** avec
   la quantité reçue, au prix d'achat.

Les articles qui existent déjà ne sont **pas** modifiés (ni leur prix, ni leur
code-barres) ; seule la réception (entrée en stock) est créée.

## Conformité avec le logiciel

L'outil reproduit fidèlement le fonctionnement interne de l'application :

| Élément              | Comportement                                                        |
|----------------------|---------------------------------------------------------------------|
| `NOPIECE` / `NOITEM` | tirés des générateurs `NEXTPIECE` / `NEXTITEM` (pas de collision)   |
| `REF_PIECE`          | attribué automatiquement par le trigger `INSERT_PIECE`              |
| Ligne (`ITEM`)       | `COEFF = 1`, `COEFF_TR = 0`, `ANNULEE = 1` → entrée en stock        |
| `PRIXHT` de la ligne | = prix d'achat → le **PUMP** (coût moyen pondéré) est correct       |
| Totaux pièce         | `MONTANTHT`, `TVA`, `MONTANTTTC` recalculés                         |

> Vérifié via la procédure `SPSTOCK` du logiciel : après import, la quantité
> en stock et le PUMP correspondent exactement à une réception saisie à la
> main.

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
- `default_famille` / `default_unite` : famille et unité affectées aux
  articles créés (créées si absentes).
- `match_famille_par_intitule` : si `true`, la colonne `Famille` de l'Excel
  est rapprochée d'une famille existante par son intitulé ; sinon la famille
  par défaut est utilisée.
- `barcode_depuis_ref` : si `true` (défaut), la référence article sert de
  code-barres quand l'Excel n'a pas de colonne code-barres.
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
