# Agrégateur de sondages — Présidentielle 2027

## Objectif

Construire un agrégateur de sondages pour l'élection présidentielle française de 2027, inspiré du modèle de Nate Silver (FiveThirtyEight). L'outil produit une moyenne pondérée des sondages, corrigée des biais systématiques des instituts, et permet de suivre l'évolution des candidats dans le temps.

Le projet est porté par un journaliste politique qui maîtrise la collecte de données et les bases de Python/pandas, mais n'est pas data scientist. Le code doit rester lisible et maintenable.

## Stack technique

- Python 3, pandas, numpy
- Visualisation : Matplotlib, Plotly ou Flourish (à déterminer)
- Données stockées en fichiers Excel (.xlsx) pour la saisie manuelle, converties en CSV ou lues directement par pandas

## Architecture du modèle

### Pipeline de traitement d'un nouveau sondage

1. **Correction du house effect** : soustraire le biais systématique de l'institut pour chaque candidat
2. **Calcul du poids composite** : qualité institut × fraîcheur × √taille échantillon
3. **Intégration dans la moyenne pondérée**

La moyenne pondérée pour un candidat :

```
score_agrégé = Σ (score_corrigé_i × poids_i) / Σ poids_i
```

### Facteurs de pondération

#### 1. Qualité historique de l'institut (pollster rating)

Calculée à partir des **sondages terminaux** (dernier sondage publié par un institut avant chaque scrutin), comparés au résultat réel. Métrique : moyenne des erreurs absolues sur les candidats au-dessus du seuil (5% pour les présidentielles, 3% pour les européennes).

```
erreur_scrutin = (1/n) × Σ |estimation_candidat_i − résultat_candidat_i|
```

On peut aussi utiliser la RMSE pour pénaliser davantage les grosses erreurs.

L'erreur est agrégée par institut sur tous les scrutins, avec pondération dégressive dans le temps (les scrutins récents pèsent plus).

Conversion en poids : inverse de l'erreur (ou inverse du carré), normalisé par rapport au meilleur institut.

#### 2. Fraîcheur du sondage

Décroissance exponentielle. Chaque jour réduit le poids du sondage. Paramètre : demi-vie, à fixer entre 10 et 30 jours (plus courte à l'approche du scrutin). Point de départ suggéré : 14 jours.

```
poids_fraîcheur = 0.5 ^ (jours_écoulés / demi_vie)
```

#### 3. Taille d'échantillon

Poids proportionnel à √n. La plupart des sondages français sont entre 1000 et 2000, donc ce facteur est peu discriminant.

#### 4. Méthodologie (optionnel)

Téléphone, en ligne, mixte — chaque mode a ses biais propres. Peut être intégré comme facteur correctif.

### House effect (biais directionnel)

Concept distinct du rating de qualité. Calculé en comparant chaque institut à la moyenne de tous les instituts sur une période glissante (typiquement 6 mois). Le house effect est signé, par candidat.

```
house_effect_institut_candidat = moyenne_institut(candidat) − moyenne_tous_instituts(candidat)
```

Mis à jour régulièrement au fil de la campagne.

## Base historique pour le rating

### Scrutins à compiler

Fenêtre : 2012–2024. Pas de raison de remonter plus loin (instituts et paysage politique trop différents).

| Scrutin | Type | Poids dans le rating |
|---|---|---|
| Présidentielle 2012 T1 | National, multi-candidats | 1.0 (mais pondéré par ancienneté) |
| Présidentielle 2012 T2 | National, duel | 0.5 |
| Européennes 2014 | National, listes | 1.0 |
| Présidentielle 2017 T1 | National, multi-candidats | 1.0 |
| Présidentielle 2017 T2 | National, duel | 0.5 |
| Européennes 2019 | National, listes | 1.0 |
| Présidentielle 2022 T1 | National, multi-candidats | 1.0 |
| Présidentielle 2022 T2 | National, duel | 0.5 |
| Législatives 2022 (IV nationales T1) | National mais imparfait | 0.3 |
| Européennes 2024 | National, listes | 1.0 |
| Législatives 2024 (IV nationales T1) | National, anticipé, contexte volatile | 0.3 |

Les T2 présidentiels sont moins discriminants (deux candidats, instituts globalement bons). Les législatives sont exploitables via les intentions de vote nationales T1, mais avec un poids réduit car la comparaison avec le résultat réel agrégé est imparfaite.

Ne pas utiliser : régionales, départementales, municipales (trop fragmentaires).

### Seuils de prise en compte des candidats

- **Présidentielles** : candidats ayant obtenu ≥ 5% au résultat réel
- **Européennes** : listes ayant obtenu ≥ 3% au résultat réel
- **Législatives** : partis ayant obtenu ≥ 5% au résultat réel agrégé

Les petits candidats introduisent du bruit (erreur relative disproportionnée) pour peu de signal.

### Données à collecter par sondage terminal

- Institut
- Commanditaire (média)
- Date de terrain (début et fin)
- Date de publication
- Taille d'échantillon
- Méthodologie (en ligne, téléphone, mixte)
- Scores par candidat/liste au-dessus du seuil

Sources : pages Wikipédia des sondages par élection, site de la commission des sondages, NSPPolls.

## Règles de recomposition des périmètres

Tous les sondages doivent mesurer le même périmètre par bloc politique. Quand les instituts découpent différemment l'offre :

- **Fusionner systématiquement** les composantes d'un même bloc (ex : RN + UDR = une seule ligne "RN et alliés" pour les européennes/législatives 2024)
- Suivre la logique des nuances du ministère de l'Intérieur comme référence
- **Documenter chaque recomposition** dans une colonne "notes" du tableur

Cas connus :
- Législatives/européennes 2024 : certains instituts séparent RN et UDR (ex-LR Ciotti), d'autres non → fusionner
- Possibles cas similaires en 2027 avec des alliances mouvantes

## Structure des fichiers de données

Un fichier Excel par scrutin (ou un fichier unique avec une feuille par scrutin). Structure de chaque feuille :

**Section métadonnées** : nom du scrutin, date, seuil, poids dans le rating

**Section données** :
- Ligne 1 : résultat réel
- Lignes suivantes : un institut par ligne, avec date, échantillon, scores par candidat

**Section écarts** (calculée automatiquement) :
- Écart absolu par candidat par institut
- Erreur moyenne par institut

Un template Excel a déjà été produit (template_sondages.xlsx).

## Roadmap de développement

### Phase 1 — Moyenne pondérée simple
Moyenne des N derniers sondages pondérée par la fraîcheur (décroissance exponentielle). C'est le MVP.

### Phase 2 — Pondération par échantillon
Ajouter √n comme facteur de poids.

### Phase 3 — House effects
Calculer le biais directionnel de chaque institut et le corriger avant agrégation.

### Phase 4 — Pollster rating
Intégrer la base historique et pondérer par la qualité de l'institut.

### Phase 5 — Visualisation
Graphique en courbes (temps × score agrégé par candidat), éventuellement avec rubans d'incertitude.

Chaque phase améliore le modèle de façon mesurable. On peut s'arrêter à n'importe quelle phase et avoir un outil fonctionnel.
