# ContactsDoctor — diagnostic et réparation des contacts

Outil en ligne de commande (Windows) pour comprendre et réparer un carnet de contacts abîmé
après une migration Exchange / un changement de téléphone / une « fusion des contacts » :

1. lit les contacts **du téléphone** (export `.vcf` de l'app Samsung Contacts),
2. lit les contacts **actuels d'Exchange** (via Outlook Classic),
3. lit les contacts **d'avant la migration** (fichier `.pst`),
4. regroupe les fiches qui sont la même personne (e-mail, numéro, nom même inversé),
5. dit **où** est le problème (téléphone, Exchange, ou les deux),
6. propose un **plan de correction** relisible dans Excel, l'applique avec **sauvegarde**
   et **journal**, et sait l'**annuler**.

Rien n'est jamais supprimé : les doublons sont déplacés dans un sous-dossier
« Doublons (ContactsDoctor) » et chaque contact modifié est d'abord sauvegardé en `.msg`.

> ⚠ Les fichiers produits (rapport, CSV, instantanés, sauvegardes) contiennent les données
> personnelles du client. Les garder sur le poste du client / le dossier d'intervention,
> ne pas les envoyer ni les committer (le `.gitignore` les exclut).

## Prérequis

- Windows avec **Outlook Classic** (le « nouvel Outlook » n'a pas d'accès COM).
- Le profil Outlook du client (compte Exchange) configuré sur ce poste.
- `ContactsDoctor.exe` (autonome, pas besoin d'installer .NET) — voir « Compiler » plus bas.

## 1. Sortir les contacts du Samsung

Dans l'app **Contacts** de Samsung :

1. Menu **☰** → **Gérer les contacts** → **Importer ou exporter des contacts** → **Exporter**.
2. Choisir **Stockage interne** (ou carte SD).
3. Choisir les comptes à exporter. **Conseil :** faire **un export par compte**
   (Téléphone, Samsung account, Exchange/Outlook, Google…) puis un export « tous » :
   cela montre où les contacts sont réellement stockés.
4. Le fichier `Contacts_….vcf` est créé dans `Documents` (ou à la racine). Le copier sur le PC
   (câble USB, ou l'envoyer à soi-même).

Vérifier aussi **Contacts → ☰ → Paramètres → Contacts à afficher** : si « Exchange » est la
source principale, le téléphone se synchronise avec Exchange et **corriger Exchange suffit**.

## 2. Préparer le PST

Travailler sur une **copie** du PST d'avant la migration (et retirer l'attribut « Lecture seule »,
sinon Outlook refuse de l'ouvrir). L'outil l'ouvre temporairement dans Outlook puis le referme.

## 3. Diagnostic

```bat
ContactsDoctor stores
ContactsDoctor analyze --ref pst=D:\Client\copie-ancien.pst --source exchange=outlook --source telephone=D:\Client\Contacts_tous.vcf --out D:\Client\diag
```

- `--ref` : la source qui fait foi (le PST d'avant la migration).
- `--source` : autant que nécessaire (`outlook` = boîte Exchange par défaut,
  `outlook:<nom>` = autre boîte, ou un fichier `.vcf`, `.csv`, `.pst`, `.json`).
  On peut ajouter un `--source` par export Samsung (`tel-exchange=…vcf`, `tel-local=…vcf`).

Résultats dans le dossier `--out` :

| Fichier | Contenu |
|---|---|
| `rapport.html` | Conclusions en clair + résumé par source + listes détaillées |
| `problemes.csv` | Un problème par ligne (inversion, doublon, manquant, fusion suspecte…) |
| `groupes.csv` | Toutes les fiches de chaque personne côte à côte (filtre par « Groupe ») |
| `plan-<source>.csv` | Plan de correction à relire (colonne **Appliquer** OUI/NON) |
| `<source>-corrige.vcf` | Pour une source `.vcf` : le carnet corrigé, prêt à réimporter |
| `instantanes\*.json` | Photo de chaque source au moment de l'analyse |

### Ce que l'outil détecte

| Problème | Signification |
|---|---|
| Prénom / nom inversés | Même personne, prénom et nom échangés par rapport à la référence |
| Nom dans un autre ordre / champ | Mêmes mots mais tout dans « prénom », ou ordre différent |
| Nom différent | Nom raccourci, orthographe différente… |
| Doublons | La même personne plusieurs fois dans la même source |
| Manquant | Présent dans la référence, absent de cette source |
| En plus | Absent de la référence (souvent : créé après la migration — normal) |
| E-mails / numéros perdus | Coordonnées de la référence absentes de cette source |
| Fusion suspecte | La fiche contient les coordonnées d'une **autre** personne de la référence (typique de « Fusionner les contacts ») |

Un numéro/e-mail partagé par 3 personnes ou plus (standard d'entreprise) n'est pas utilisé
pour rapprocher les fiches.

### Lire les conclusions

- Inversions **seulement dans Exchange** → le problème vient de la migration Exchange.
- **Les mêmes** inversions dans Exchange et le téléphone → le téléphone est synchronisé avec
  Exchange : corriger Exchange, le téléphone suivra.
- Inversions **différentes** → les carnets ne sont pas synchronisés : corriger chacun.
- Plus de doublons qu'avant → la migration ou la synchro a créé des copies.
- Fusions suspectes → l'outil de fusion Samsung a mélangé des personnes : à corriger à la main
  (le rapport indique qui est mélangé avec qui).

## 4. Corriger Exchange

1. Ouvrir `plan-exchange.csv` dans Excel, relire, mettre **OUI/NON** dans « Appliquer »,
   enregistrer au format CSV (séparateur `;`, UTF-8 ou ANSI : les deux sont acceptés). Par défaut :
   - `SwapName` (remettre prénom/nom dans le bon ordre) → OUI
   - `MoveDuplicate` (doublon → sous-dossier, ses e-mails/numéros en plus sont d'abord
     ajoutés au contact conservé) → OUI
   - `AddFromReference` (recréer un contact disparu depuis le PST) → OUI
   - `RenameFromReference` (nom différent : reprendre celui du PST) → **NON**, à décider au cas par cas
2. Simulation (ne modifie rien) :
   ```bat
   ContactsDoctor apply D:\Client\diag\plan-exchange.csv --pst D:\Client\copie-ancien.pst
   ```
3. Application réelle (demande de taper OUI) :
   ```bat
   ContactsDoctor apply D:\Client\diag\plan-exchange.csv --pst D:\Client\copie-ancien.pst --execute
   ```
   Crée `diag\sauvegarde_<date>\` avec un `.msg` par contact touché et `journal.csv`.
   Un contact dont le nom a changé depuis l'analyse est ignoré (jamais écrasé).
4. Relancer `analyze` pour vérifier.

### Annuler

```bat
ContactsDoctor undo D:\Client\diag\sauvegarde_<date>\journal.csv --execute
```

Remet les noms d'origine, ramène les doublons dans leur dossier (et retire les données
fusionnées), et met les contacts ajoutés dans « Éléments supprimés ».

## 5. Corriger le téléphone (seulement s'il n'est PAS synchronisé avec Exchange)

`telephone-corrige.vcf` contient le carnet corrigé (inversions remises, doublons fusionnés,
contacts manquants ajoutés). Pour l'utiliser : sauvegarder le téléphone, supprimer les contacts
du compte concerné, puis **Importer** ce fichier. Si le téléphone est synchronisé avec
Exchange, ne pas faire cette étape.

## Compiler

```bash
dotnet test ContactsDoctor/Tests                       # tests (Linux/Windows, sans Outlook)
dotnet publish ContactsDoctor/Cli -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o publish
```

Structure : `Core/` (lecture vCard/CSV, rapprochement, rapport, plan — sans dépendance
Windows, testé), `Cli/` (ligne de commande + accès Outlook par COM en liaison tardive,
compatible toutes versions d'Outlook Classic), `Tests/` (xUnit).
