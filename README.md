# Team Planning

Outil Excel avec macros (VBA) pour gérer le planning d'une équipe : saisie des tâches, chaînage automatique, prise en compte des jours fériés, des congés et des formations, grille de charge par personne et par semaine avec code couleur, et résumé numérique.

## Contenu du dépôt

| Chemin | Description |
|---|---|
| `Sources/PlanningEquipe.bas` | Module VBA unique : construction du classeur, calculs et mise à jour |
| `Documents/Guide_installation_et_utilisation.odt` | Guide complet : installation, utilisation, règles de calcul, personnalisation |
| `Documents/odt-src/` | Sources XML du guide ODT |
| `.github/workflows/build-odt.yml` | Action GitHub qui assemble l'ODT depuis `Documents/odt-src/` |

## Installation en 5 minutes

1. Créer un nouveau classeur Excel et l'enregistrer au format `.xlsm` (prenant en charge les macros).
2. Ouvrir l'éditeur VBA (`Alt + F11`), puis **Insertion > Module** et coller le contenu de `Sources/PlanningEquipe.bas`. (Ou **Fichier > Importer un fichier** et sélectionner directement le `.bas`.)
3. De retour dans Excel, `Alt + F8` > exécuter `ConstruireClasseur` : les 5 feuilles sont créées, mises en forme et préremplies d'exemples.
4. Enregistrer, puis activer les macros à l'ouverture du fichier.
5. Compatibilité : Excel desktop Windows 2010 ou plus récent. Les macros ne fonctionnent ni dans Excel Online ni dans Excel mobile.

## Feuilles du classeur

- **Saisie** : une ligne = une tâche (prénom, projet, description, durée en jours ouvrés, début optionnel, avancement, état, retard). Début laissé vide = enchaînement automatique sur la tâche précédente de la même personne. Dates, retards et états calculés automatiquement.
- **Charge** : grille personnes x semaines (ISO), code couleur selon le taux de charge (vert clair / vert / orange / rouge, gris si semaine indisponible), capacité et total équipe, grille des absences (congés, formations), légende.
- **Résumé** : charge par personne, par projet (avancement pondéré), par semaine, et compteurs de tâches.
- **Équipe** : membres et capacités (temps partiel géré), absences datées (congés, formations), paramètres configurables (horizon en semaines, seuils de couleur, capacité par défaut).
- **Jours Fériés** : jours fériés français calculés automatiquement (Pâques, Ascension, Pentecôte inclus) + jours exceptionnels libres (ponts, fermetures).

## Principes de calcul

- Jour disponible = lundi à vendredi, hors jours fériés, hors absences de la personne.
- Fin calculée = durée en jours disponibles à partir du début effectif (un congé pendant la tâche la décale).
- Charge d'une semaine = jours couverts par les tâches, ramenés à la capacité de la personne.
- Retard = jours ouvrés entre la fin calculée et aujourd'hui, si la tâche n'est pas terminée.

## Macros disponibles (`Alt + F8`)

| Macro | Rôle |
|---|---|
| `ConstruireClasseur` | Construit ou reconstruit les 5 feuilles |
| `MajTout` | Recalcule tout (identique au bouton « Mettre à jour le planning ») |
| `AjouterFeriesAnnee` | Ajoute les jours fériés français d'une année donnée |

## Regénérer le guide ODT

Toute modification des fichiers de `Documents/odt-src/` déclenche automatiquement l'action GitHub `build-odt`, qui reconstruit `Documents/Guide_installation_et_utilisation.odt` et le commit dans le dépôt.
