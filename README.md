# Team Planning

Outil Excel avec macros (VBA) pour gérer le planning d'une équipe : saisie des tâches, chaînage automatique, prise en compte des jours fériés, des congés et des formations, grille de charge par personne et par semaine avec code couleur, et résumé numérique.

## Contenu du dépôt

| Chemin | Description |
|---|---|
| `Sources/PlanningEquipe.bas` | Module VBA, **encodage ANSI (Windows-1252), fins de ligne CRLF** : à télécharger puis importer directement dans l'éditeur VBE (accents préservés) |
| `Sources/PlanningEquipe_UTF8.bas` | Même module en UTF-8 : pour lecture et copier-coller depuis github.com |
| `Documents/Guide_installation_et_utilisation.odt` | Guide complet : installation, utilisation, règles de calcul, personnalisation |
| `Documents/odt-src/` | Sources XML du guide ODT |
| `.github/workflows/build-assets.yml` | Action GitHub : régénère le `.bas` ANSI et l'ODT à partir des sources |

## Installation en 5 minutes

1. Créer un nouveau classeur Excel et l'enregistrer au format `.xlsm` (prenant en charge les macros).
2. Ouvrir l'éditeur VBA (`Alt + F11`), puis **Fichier > Importer un fichier** et sélectionner `Sources/PlanningEquipe.bas` (encodage ANSI : les accents sont préservés à l'import).
   - Alternative par copier-coller : ouvrir `Sources/PlanningEquipe_UTF8.bas`, copier le code **sans la première ligne `Attribute VB_Name`** et le coller dans un nouveau module (**Insertion > Module**).
3. De retour dans Excel, `Alt + F8` > exécuter `ConstruireClasseur` : les 5 feuilles sont créées, mises en forme et préremplies d'exemples.
4. Enregistrer, puis activer les macros à l'ouverture du fichier.

## Compatibilité

Syntaxe vérifiée pour **Office 2024** et **Microsoft 365**, compatible Excel 2016 / 2019 / 2021 (Excel desktop Windows, 32 et 64 bits). Le module n'utilise que des instructions VBA et des constantes Excel standard (la compatibilité est aussi notée en commentaire d'en-tête du module). Les macros ne fonctionnent ni dans Excel Online ni dans Excel mobile.

## Feuilles du classeur

- **Saisie** : une ligne = une tâche (prénom, projet, description, durée en jours ouvrés, début optionnel, avancement, état, retard). Début laissé vide = enchaînement automatique sur la tâche précédente de la même personne. Dates, retards et états calculés automatiquement ; les lignes saisies juste sous le tableau sont intégrées automatiquement au tableau. Mise en évidence : lignes terminées grisées, tâches en cours en texte bleu, état orange si une tâche non commencée a dépassé sa date de début, rouge si l'état contredit l'avancement.
- **Charge** : grille personnes x semaines (ISO), code couleur selon le taux de charge (vert clair / vert / orange / rouge, gris si semaine indisponible), capacité et total équipe, grille des absences (congés, formations), légende.
- **Résumé** : charge par personne, par projet (avancement pondéré), par semaine, et compteurs de tâches.
- **Équipe** : membres et capacités (temps partiel géré), absences datées (congés, formations), paramètres configurables (horizon en semaines, seuils de couleur, capacité par défaut).
- **Jours Fériés** : jours fériés français calculés automatiquement (Pâques, Ascension, Pentecôte inclus) + jours exceptionnels libres (ponts, fermetures).

## Principes de calcul

- Jour disponible = lundi à vendredi, hors jours fériés, hors absences de la personne.
- Fin théorique = durée en jours disponibles à partir du début effectif (un congé pendant la tâche la décale).
- Fin calculée = projection tenant compte d'aujourd'hui et de l'avancement : tâche en cours = reste à faire (durée x (100 - avancement) %, arrondi au jour supérieur) planifié depuis aujourd'hui ; tâche non commencée dont le début est dépassé = durée complète depuis aujourd'hui ; autres cas = fin théorique.
- Retard = écart en jours ouvrés entre la fin calculée et la fin théorique.
- Charge d'une semaine = jours couverts par les tâches, ramenés à la capacité de la personne.

## Macros disponibles (`Alt + F8`)

| Macro | Rôle |
|---|---|
| `ConstruireClasseur` | Construit ou reconstruit les 5 feuilles |
| `MajTout` | Recalcule tout (identique au bouton « Mettre à jour le planning ») |
| `AjouterFeriesAnnee` | Ajoute les jours fériés français d'une année donnée |

## Regénérer les fichiers générés

Toute modification de `Sources/PlanningEquipe_UTF8.bas` ou de `Documents/odt-src/` déclenche l'action GitHub `build-assets`, qui régénère le module ANSI (`Sources/PlanningEquie.bas`) et l'ODT, puis les commit automatiquement dans le dépôt.
