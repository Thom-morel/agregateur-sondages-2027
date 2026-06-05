"""
probabilites.py  —  PALIER 1 : de la moyenne à la probabilité
──────────────────────────────────────────────────────────────
Transforme la moyenne pondérée produite par agregateur.py en
PROBABILITÉS de qualification au second tour, par simulation.

Idée centrale
─────────────
Un score de sondage n'est PAS une vérité, c'est le centre d'une
cloche (loi normale) dont la largeur σ = l'erreur typique des
sondages. On tire au hasard N « élections » dans ces cloches, on
classe les blocs à chaque tirage, et on compte :
  • la part de tirages où le bloc finit 1er
  • la part de tirages où le bloc est dans le top 2 (= qualifié T2)

C'est tout. « 70 % de chances » = « dans 70 % des élections
simulées, le bloc se qualifie ».

Limite assumée du palier 1
──────────────────────────
Les tirages sont INDÉPENDANTS d'un bloc à l'autre : on ignore le
fait que les scores sont corrélés (si LFI monte, un autre baisse)
et la somme ne fait pas exactement 100 % à chaque tirage. Corriger
cela = palier 2 (matrice de covariance). Pour un CLASSEMENT, c'est
sans effet notable.

Usage :
  python3 probabilites.py
"""

import csv
import numpy as np

# ══════════════════════════════════════════════════════════════
# PARAMÈTRES — modifier ici
# ══════════════════════════════════════════════════════════════

FICHIER_HISTORIQUE = 'historique_agregateur.csv'
N_SIMULATIONS      = 100_000

# σ = erreur-type des sondages, EN POINTS DE POURCENTAGE.
# C'est LE paramètre qui pilote tout.
#
# Valeur MESURÉE sur template_sondages.xlsx (agrégat final vs
# résultat réel, présidentielles T1 2012/2017/2022) : σ ≈ 2,0.
# C'est l'erreur TERMINALE : celle des sondages quand le champ
# des candidats est fixé et le vote imminent.
#
# POSTURE « INSTANTANÉ » (assumée tant qu'on est loin du vote) :
# on garde σ ≈ 2-3 et on lit le résultat comme « SI le vote avait
# lieu maintenant », PAS comme une prévision pour avril 2027. Loin
# de l'échéance, l'incertitude réelle n'est pas du bruit gaussien
# (candidatures qui surgissent, primaires, affaires) : non
# modélisable par un σ. On réévaluera la dérive à ~2-3 mois du vote.
SIGMA = 2.0

# Pour comparer la sensibilité du résultat à ce choix :
SIGMA_SENSIBILITE = [2.0, 3.0, 4.0, 5.0, 6.0]


# ══════════════════════════════════════════════════════════════
# ÉTAPE 1 — Lire la dernière ligne de l'agrégateur
# ══════════════════════════════════════════════════════════════

def lire_dernier_calcul(fichier):
    with open(fichier, encoding='utf-8') as f:
        lignes = list(csv.DictReader(f))
    if not lignes:
        raise SystemExit(f"Aucun calcul dans {fichier}. Lance d'abord agregateur.py.")
    derniere = lignes[-1]
    date = derniere.pop('date_calcul')
    derniere.pop('total', None)
    blocs   = list(derniere.keys())
    scores  = np.array([float(derniere[b]) for b in blocs])
    return date, blocs, scores


# ══════════════════════════════════════════════════════════════
# ÉTAPE 2 — Simulation Monte-Carlo
# ══════════════════════════════════════════════════════════════

def simuler(scores, sigma, n):
    """Renvoie, pour chaque bloc : P(1er) et P(top 2)."""
    k = len(scores)
    # n tirages × k blocs, autour de chaque score, largeur sigma
    tirages = np.random.normal(loc=scores, scale=sigma, size=(n, k))
    tirages = np.clip(tirages, 0, None)          # pas de score négatif

    # classement décroissant : ordre[:, 0] = vainqueur de chaque tirage
    ordre  = np.argsort(-tirages, axis=1)
    premier = ordre[:, 0]
    top2    = ordre[:, :2]

    p_premier = np.bincount(premier, minlength=k) / n
    p_top2    = np.bincount(top2.ravel(), minlength=k) / n

    # duels du second tour les plus fréquents (paire non ordonnée)
    a = np.minimum(top2[:, 0], top2[:, 1])
    b = np.maximum(top2[:, 0], top2[:, 1])
    code = a * k + b
    comptes = np.bincount(code, minlength=k * k)
    duels = [((c // k, c % k), n_occ / n)
             for c, n_occ in enumerate(comptes) if n_occ > 0]
    duels.sort(key=lambda x: -x[1])
    return p_premier, p_top2, duels


# ══════════════════════════════════════════════════════════════
# ÉTAPE 3 — Affichage
# ══════════════════════════════════════════════════════════════

def barre(p, largeur=30):
    return '█' * int(round(p * largeur))


def main():
    date, blocs, scores = lire_dernier_calcul(FICHIER_HISTORIQUE)
    p_premier, p_top2, duels = simuler(scores, SIGMA, N_SIMULATIONS)

    ordre_aff = np.argsort(-scores)

    print('═' * 72)
    print(f"PROBABILITÉS — T1 PRÉSIDENTIELLE 2027  (données du {date})")
    print(f"σ = {SIGMA} pts · {N_SIMULATIONS:,} simulations".replace(',', ' '))
    print("⚠ INSTANTANÉ : probabilités SI le vote avait lieu maintenant.")
    print("  Ce n'est PAS une prévision pour avril 2027.")
    print('═' * 72)
    print()
    print(f"  {'Bloc':<12} {'Score':>6}   {'P(1er)':>7}   {'P(qualifié T2)':>15}")
    print(f"  {'─' * 60}")
    for i in ordre_aff:
        print(f"  {blocs[i]:<12} {scores[i]:>5.1f}%   "
              f"{p_premier[i]*100:>6.1f}%   {p_top2[i]*100:>6.1f}%  {barre(p_top2[i])}")
    print()

    print('─' * 72)
    print("  DUELS DE SECOND TOUR LES PLUS PROBABLES")
    print('─' * 72)
    for (i, j), p in duels[:6]:
        print(f"    {blocs[i]} vs {blocs[j]:<14} {p*100:>5.1f}%")
    print()

    # ── Sensibilité au choix de σ ──────────────────────────────
    print('─' * 72)
    print("  SENSIBILITÉ À σ — P(qualifié T2), pour voir l'effet du paramètre")
    print('─' * 72)
    entete = "  {:<12}".format('Bloc') + ''.join(f"σ={s:<5}" for s in SIGMA_SENSIBILITE)
    print(entete)
    lignes_sens = {}
    for s in SIGMA_SENSIBILITE:
        _, pt2, _ = simuler(scores, s, N_SIMULATIONS)
        lignes_sens[s] = pt2
    for i in ordre_aff:
        cellules = ''.join(f"{lignes_sens[s][i]*100:>5.0f} " for s in SIGMA_SENSIBILITE)
        print(f"  {blocs[i]:<12}" + cellules)
    print()
    print("  → Plus σ est grand (= on est loin du vote, plus d'incertitude),")
    print("    plus les probabilités se rapprochent les unes des autres.")
    print()


if __name__ == '__main__':
    main()
