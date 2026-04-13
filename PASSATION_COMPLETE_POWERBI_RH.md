# Passation Power BI RH - Etat au 2026-04-12

## 1. Objet

Ce document remplace l'ancienne passation centree sur `Data.xlsx`.

Etat de reference au 12 avril 2026:

- projet Power BI: `PRESENTATION.pbip`
- rapport: `PRESENTATION.Report`
- modele semantique: `PRESENTATION.SemanticModel`
- template de collecte cible: `RH_Collecte_Neemba_2026.xlsx`
- template PBIT de reference: `T_PRESENTATION.pbit`

Le projet est maintenant aligne sur le nouveau modele de collecte `tbl_*`.

## 2. Resume executif

- le modele semantique charge les nouvelles tables `tbl_*` du template RH 2026
- un parametre `FolderPathParameter` existe pour sortir les chemins de requete du code M
- la table `HR Date` a ete etendue pour couvrir les nouvelles tables RH
- les relations manquantes entre faits et dimensions ont ete ajoutees
- les visuels de la page `Overview` ont ete migres vers les nouvelles tables
- les tables legacy du modele ont ete retirees
- un segment `Filiale` a ete ajoute sur `Overview`, `HR Performance` et `TCDP`
- le probleme `Missing_References` sur `Effectif[Average Age]` a ete corrige
- les tailles de police ont ete augmentees dans les visuels du rapport et dans le theme partage
- tous les `visual.json` se parsèrent a nouveau apres reparation

## 3. Architecture actuelle

### 3.1 Pages du rapport

- `Overview`
- `HR Performance`
- `TCDP`

### 3.2 Tables metier actives du modele

- `Effectif`
- `Embauches`
- `Departs`
- `AbsenceMensuelle`
- `FormationMensuelle`
- `RecrutementMensuel`
- `RecrutementDetail`
- `MasseSalarialeMensuelle`
- `TCDP_Headcount`
- `TCDP_Entrees`
- `TCDP_Sorties`
- `TCDP_Genre`

### 3.3 Dimensions actives

- `DIM_Filiale`
- `DIM_TypeContrat`
- `DIM_StatutGeo`
- `DIM_CatConv`
- `DIM_Sexe`
- `DIM_MotifDepart`
- `DIM_MotifRecruit`
- `DIM_TCDPLevel`
- `DIM_Nationalite`
- `DIM_AbsenceType`
- `DIM_MotifSortieTCDP`
- `HR Date`

### 3.4 Tables legacy retirees du modele

Les objets suivants ont ete supprimes du modele semantique:

- `Tableau1`
- `Tableau24`
- `Depart`
- `Embauche`
- `Liste de presence au 01-12-2025`
- `LocalDateTable_*` associees

## 4. Modifications realisees

### 4.1 Parametrage du chemin source

Fichier cle:

- `PRESENTATION.SemanticModel/definition/expressions.tmdl`

Changement:

- ajout du parametre `FolderPathParameter`
- remplacement des chemins `File.Contents(...)` en dur dans les partitions M par `FolderPathParameter & "\\RH_Collecte_Neemba_2026.xlsx"`

Important:

- le parametre existe, mais sa valeur par defaut reste un chemin local du workspace
- la bascule vers `SharePoint.Files(...)` n'est pas encore faite

### 4.2 HR Date

Fichier cle:

- `PRESENTATION.SemanticModel/definition/tables/HR Date.tmdl`

Changement:

- `HR Date` inclut maintenant les dates provenant des nouvelles tables:
  - `Effectif[SnapshotDate]`
  - `Embauches[EventDate]`
  - `Embauches[MonthStartDate]`
  - `Departs[EventDate]`
  - `Departs[MonthStartDate]`
  - toutes les tables mensuelles
  - `TCDP_Entrees`
  - `TCDP_Sorties`
  - `TCDP_Genre`

### 4.3 Relations ajoutees

Fichier cle:

- `PRESENTATION.SemanticModel/definition/relationships.tmdl`

Relations ajoutees pendant cette phase:

- `Embauches[StatutGeoCode] -> DIM_StatutGeo[REF_StatutGeo]`
- `Embauches[CategorieConventionnelleCode] -> DIM_CatConv[REF_CatConv]`
- `Departs[SexeCode] -> DIM_Sexe[REF_Sexe]`
- `Departs[NationaliteCode] -> DIM_Nationalite[REF_Nationalite]`
- relations date vers `HR Date` pour les nouvelles tables

### 4.4 Measures corrigees

Fichier cle:

- `PRESENTATION.SemanticModel/definition/tables/Effectif.tmdl`

Probleme corrige:

- erreur `Underlying Error: Missing_References`
- `Effectif[Average Age]` dependait d'une mesure `Effectif[Age]` en erreur

Correction appliquee:

- `Average Age` calcule directement l'age a partir de `DateNaissance`
- `Average Seniority` calcule directement l'anciennete a partir de `DateEntree` et `DateEmbauche`

### 4.5 Migration report

Fichiers cles:

- `PRESENTATION.Report/definition/pages/.../visuals/*/visual.json`

Changements:

- migration des visuels `Overview` vers `Effectif`, `Departs` et les dimensions actives
- ajout d'un segment `Filiale` sur les 3 pages principales
- nettoyage des references report vers les tables legacy

### 4.6 Typographie

Changements:

- augmentation des tailles de police dans les `visual.json`
- augmentation des tailles de police dans `PRESENTATION.Report/StaticResources/SharedResources/BaseThemes/CY26SU02.json`

Valeurs du theme mises a jour:

- `callout`: `26 -> 28`
- `title`: `12 -> 14`
- `header`: `12 -> 14`
- `label`: `9 -> 11`

### 4.7 Reparation des fichiers report

Incident pendant la maintenance:

- l'augmentation en masse des `fontSize` a corrompu plusieurs blocs JSON de visuels

Correction appliquee:

- reparation des blocs `fontSize`
- validation de tous les `visual.json`

Resultat:

- `COUNT=0` fichier `visual.json` invalide apres correction

## 5. Etat de validation

### 5.1 Validation statique effectuee

- tous les `visual.json` du rapport se parsèrent
- le theme `CY26SU02.json` est coherent
- les mesures `Average Age` et `Average Seniority` ne dependent plus d'une reference manquante
- le modele ne reference plus les tables legacy supprimees

### 5.2 Validation manuelle encore requise

Cette etape n'a pas pu etre faite ici:

- ouvrir `PRESENTATION.pbip` dans Power BI Desktop
- renseigner la bonne valeur de `FolderPathParameter`
- lancer un refresh complet
- verifier les pages `Overview`, `HR Performance` et `TCDP`
- verifier le segment `Filiale`
- verifier les cartes `Average Age`, `Average Seniority`, `Turnover 2025`

## 6. Points encore ouverts

### 6.1 SharePoint non implemente

Le projet est parametre, mais pas encore consolide via SharePoint:

- pas de `SharePoint.Files(...)`
- pas de requete maitre dossier
- pas de `Table.Combine` multi-fichiers

### 6.2 Parametre encore local

Le parametre existe, mais sa valeur par defaut reste:

- `C:\\Users\\eklu\\Downloads\\111\\RH\\Lab`

Il faut le remplacer selon l'environnement cible.

### 6.3 Synchronisation du segment

Le segment `Filiale` est ajoute sur chaque page.

Reste a confirmer dans Power BI Desktop:

- synchronisation du slicer entre pages
- comportement de filtrage attendu

## 7. Prochaines actions recommandees

1. Ouvrir `PRESENTATION.pbip` dans Power BI Desktop.
2. Mettre `FolderPathParameter` sur le dossier ou l'URL cible.
3. Lancer un refresh complet et corriger les dernieres erreurs eventuelles au runtime.
4. Tester le segment `Filiale` sur les 3 pages.
5. Si le deploiement groupe est confirme, remplacer le pattern `File.Contents(...)` par `SharePoint.Files(...)`.
6. Construire ensuite la consolidation multi-fichiers via requete maitre + `Table.Combine`.

## 8. Fichiers a connaitre pour la reprise

Modele:

- `PRESENTATION.SemanticModel/definition/model.tmdl`
- `PRESENTATION.SemanticModel/definition/relationships.tmdl`
- `PRESENTATION.SemanticModel/definition/expressions.tmdl`
- `PRESENTATION.SemanticModel/definition/tables/Effectif.tmdl`
- `PRESENTATION.SemanticModel/definition/tables/Departs.tmdl`
- `PRESENTATION.SemanticModel/definition/tables/HR Date.tmdl`

Report:

- `PRESENTATION.Report/definition/pages/8417700847a190be312a`
- `PRESENTATION.Report/definition/pages/hrperf8a7b6c5d4e3f2a1b`
- `PRESENTATION.Report/definition/pages/tcdp9f4e5b2c1a6d7e8f`
- `PRESENTATION.Report/StaticResources/SharedResources/BaseThemes/CY26SU02.json`

Sources:

- `RH_Collecte_Neemba_2026.xlsx`
- `T_PRESENTATION.pbit`

## 9. Conclusion

Le projet n'est plus dans l'etat initial base sur `Data.xlsx`.

Au 12 avril 2026:

- le modele est aligne sur le template RH 2026
- les visuels principaux ont ete remigres
- les tables legacy ont ete retirees du modele
- le segment `Filiale` est en place
- la typographie a ete agrandie
- le correctif `Average Age` est applique

Le point de passage obligatoire avant diffusion reste maintenant un refresh complet dans Power BI Desktop.
