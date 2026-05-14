# Intégration du Machine Learning — `silver_faces.vbs`

---

## 1. Problème fondamental du système actuel

La macro est entièrement **basée sur des règles fixes** et n'utilise qu'**une seule caractéristique** pour prendre chaque décision :

$$\text{Aire (mm}^2\text{)} < 0{,}01 \Rightarrow \text{silver face}$$

Cette approche est fragile. Les deux constantes magiques — `SEUIL_MM2 = 0.01` et `SEUIL_DELETE = 0.0001` — ont été choisies manuellement par un ingénieur. Elles ne peuvent pas s'adapter à :

- Des domaines industriels différents (aéronautique vs. automobile : les tolérances diffèrent d'un ordre de grandeur)
- Des styles de modélisation variés (import IGES/STEP vs. modélisation native CATIA)
- Le contexte géométrique (une surface de 0,008 mm² adjacente à une aube de turbine est un **défaut critique** ; la même aire sur un panneau décoratif est du **bruit**)

Le Machine Learning remplace ces seuils figés par une **fonction de décision apprise** à partir de l'historique réel des corrections.

---

## 2. Espace de caractéristiques — Ce que l'on peut extraire dès maintenant

L'API SPA de CATIA expose bien plus que la simple aire. Voici le **vecteur de caractéristiques** que l'on peut construire pour chaque `HybridShape`, entièrement extractible via l'objet `Measurable` :

| #   | Caractéristique              | API CATIA                         | Intérêt                                                                                         |
| --- | ---------------------------- | --------------------------------- | ----------------------------------------------------------------------------------------------- |
| 1   | **Aire** (mm²)               | `oM.Area × 1e6`                   | Déjà utilisée — signal principal                                                                |
| 2   | **Périmètre** (mm)           | `oM.Perimeter`                    | Les faces fines ont un périmètre énorme par rapport à leur aire                                 |
| 3   | **Compacité**                | $4\pi \cdot A / P^2$              | Quotient isopérimétrique — cercle = 1,0, sliver ≈ 0,001                                         |
| 4   | **Rapport d'aspect**         | BBox max / BBox min               | Les slivers sont extrêmement allongées — ratio pouvant dépasser 1000:1                          |
| 5   | **Dimensions du BBox**       | `MinimumMeasure`                  | 3 valeurs : longueur, largeur, hauteur                                                          |
| 6   | **Type de surface**          | `TypeName(oHS)`                   | `HybridShapePlane` vs. `HybridShapeCircle` vs. NURBS — les imports ont une signature différente |
| 7   | **Profondeur dans l'arbre**  | parcours récursif                 | Imbrication profonde → souvent résultat d'une opération Booléenne/Trim                          |
| 8   | **Nombre de surfaces sœurs** | `oHB.HybridShapes.Count`          | Un corps contenant une seule surface après un Trim est suspect                                  |
| 9   | **Position dans l'arbre**    | index `Item(j)`                   | La dernière surface d'un corps après un Trim est souvent un résidu                              |
| 10  | **Courbure minimale**        | `GetCurvatureAnalysis`            | Forte courbure → artefact d'Offset                                                              |
| 11  | **Déviation de la normale**  | comparaison à la moyenne du corps | Une face perpendiculaire à toutes ses voisines est anormale                                     |

La métrique de **compacité** est particulièrement discriminante :

$$C = \frac{4\pi A}{P^2}, \quad C \in (0,\ 1]$$

Un cercle parfait donne $C = 1$. Une silver face — par exemple une bande de 10 mm × 0,001 mm — donne :

$$C = \frac{4\pi \times 0{,}01}{(20{,}002)^2} \approx 3{,}1 \times 10^{-4}$$

Cela discrimine les slivers bien mieux que l'aire seule, car une **petite surface valide** (congé, chanfrein) est compacte, tandis qu'une **silver face** est dégénérée en forme.

---

## 3. Algorithmes candidats — Du plus simple au plus avancé

### 3.1 Régression logistique — La référence de base

Remplacer les deux seuils par une frontière de décision linéaire apprise sur le vecteur de caractéristiques :

$$P(\text{sliver}) = \sigma(\mathbf{w}^T \mathbf{x} + b)$$

- **Entrées** : aire, compacité, rapport d'aspect, encodage du type de surface
- **Sortie** : probabilité que cette surface soit une silver face
- **Pourquoi commencer ici** : interprétable, entraînable sur moins de 500 exemples étiquetés, rapide à déployer

Le seuil de décision devient réglable : on ne choisit plus 0,01 mm² arbitrairement — le modèle le détermine en fonction du contexte.

---

### 3.2 Forêt aléatoire (Random Forest) — Le choix pragmatique

Pour un système de production réel, la **forêt aléatoire** est le **premier modèle sérieux à adopter** :

- Gère les types de caractéristiques mixtes (continu : aire, compacité ; catégoriel : type de surface)
- Fournit naturellement les **importances des caractéristiques** — on découvre lesquelles comptent le plus
- Robuste aux valeurs aberrantes et aux mesures manquantes (CATIA échoue parfois à calculer l'aire pour une géométrie dégénérée)
- Ne nécessite pas de normalisation des caractéristiques
- Se généralise bien à partir de 1 000 à 10 000 exemples étiquetés

Chaque arbre de la forêt vote :

$$\hat{y} = \frac{1}{T} \sum_{t=1}^{T} h_t(\mathbf{x})$$

La forêt fournit également des **probabilités calibrées**, permettant d'ajuster l'équilibre entre :

- **Faux négatifs** (slivers manquées → coûteux en fabrication) — priorité au rappel
- **Faux positifs** (surfaces valides signalées à tort → agaçant mais non dangereux)

En contrôle qualité, le **rappel est primordial**. On fixerait un seuil de probabilité bas (ex. $P > 0{,}3$ = signaler) pour ne rien manquer.

---

### 3.3 Forêt d'isolation (Isolation Forest) — Détection d'anomalies non supervisée

Le principal obstacle pratique à l'apprentissage supervisé est le **coût de l'étiquetage**. Chaque exemple étiqueté nécessite qu'un expert ouvre un fichier CAO et décide : « oui, c'est une sliver / non, ce n'est pas une sliver ».

**L'Isolation Forest ne nécessite aucune étiquette**. Elle apprend à quoi ressemble une géométrie « normale » et signale tout ce qui s'en écarte :

- Elle partitionne aléatoirement l'espace de caractéristiques avec des arbres
- Les points anormaux (comme les slivers) sont isolés en **moins d'étapes** — ils sont faciles à séparer de la masse
- Le **score d'anomalie** est inversement proportionnel à la longueur moyenne du chemin

$$s(\mathbf{x}, n) = 2^{-\frac{E[h(\mathbf{x})]}{c(n)}}$$

où $h(\mathbf{x})$ est la longueur du chemin et $c(n)$ la constante de normalisation.

On l'entraîne sur un ensemble de surfaces issues de pièces connues comme propres. Ensuite, on l'applique aux nouvelles pièces — tout ce dont le score dépasse 0,7 est signalé. **Aucune étiquette n'est requise pour l'entraînement.**

C'est l'algorithme à implémenter **en premier**, précisément parce qu'il démarre à partir de zéro.

---

### 3.4 Classifieur multi-classe pour la stratégie de correction

C'est ici que le ML apporte une valeur ajoutée considérable **au-delà de la simple détection**. La logique de correction actuelle est :

```
aire < 0,0001  →  Stratégie A (suppression)
aire < 0,01    →  Stratégie B (healing)
sinon          →  traitement manuel
```

Un classifieur entraîné pourrait prédire **laquelle des 5 stratégies** (A, B, résidu de Trim, artefact d'Offset, artefact d'import) est la plus appropriée — ce que les trois cas manuels actuels ne permettent pas de déterminer automatiquement.

**Algorithme** : Arbres à gradient boosté (XGBoost / LightGBM)

Caractéristiques qui permettent de distinguer les cas :

- **Résidu de Trim** → rapport d'aspect élevé, surface en fin de HybridBody, le corps parent contient un `HybridShapeSplit` ou `HybridShapePartialComplement`
- **Artefact d'Offset** → forte courbure minimale sur les surfaces voisines, type généralement spline/NURBS
- **Artefact d'import** → pas d'historique de construction, `TypeName` retourne `HybridShapeDatumSurface`
- **Quasi-dégénérée** → aire infime, compacité proche de zéro, périmètre proche de zéro

Cela transforme les trois cas « traitement manuel » en corrections automatisées.

---

### 3.5 Réseaux de neurones sur graphes (GNN) — La frontière avancée

L'approche la plus puissante. Au lieu de traiter chaque surface de manière indépendante, on modélise l'**ensemble du Part comme un graphe** :

- **Nœuds** = HybridShapes (surfaces), chacun avec son vecteur de caractéristiques
- **Arêtes** = adjacence physique entre surfaces (arêtes partagées, écarts < tolérance)
- **Passage de messages** = chaque surface apprend de ses voisins topologiques

$$\mathbf{h}_v^{(k)} = \text{UPDATE}\!\left(\mathbf{h}_v^{(k-1)},\ \text{AGGREGATE}\!\left(\{\mathbf{h}_u^{(k-1)} : u \in \mathcal{N}(v)\}\right)\right)$$

Une silver face présente par définition des **relations anormales avec ses voisines** (extrêmement fine, quasi parallèle aux surfaces adjacentes, partage presque tout son périmètre avec une seule voisine). Les GNN capturent cela topologiquement, et non seulement géométriquement. Cette approche requiert davantage de données (~10 000 pièces étiquetées) mais atteindrait une détection quasi parfaite, y compris pour les cas difficiles que les méthodes basées sur l'aire seule ne repèrent pas.

---

## 4. Architecture ML complète

```
┌─────────────────────────────────────────────────────────────┐
│  CATIA  (couche VBS — existante)                             │
│                                                              │
│  silver_faces.vbs → extraction des caractéristiques         │
│    (aire, compacité, rapport d'aspect, type, position)      │
│    → écriture dans features.json (un enregistrement/surf.)  │
└────────────────────────┬────────────────────────────────────┘
                         │ appel HTTP / sous-processus
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  Service ML Python  (Flask / FastAPI)                        │
│                                                              │
│  1. Chargement de features.json                              │
│  2. Prétraitement (normalisation, encodage du type)          │
│  3. Inférence :                                              │
│     Phase 1 : Isolation Forest  → score d'anomalie          │
│     Phase 2 : Random Forest     → P(sliver), classe strat.  │
│  4. Retour JSON :                                            │
│     { "surface_name": "...",                                 │
│       "is_sliver": true,                                     │
│       "confidence": 0.94,                                    │
│       "recommended_strategy": "B_Healing",                   │
│       "reason": "rapport d'aspect élevé + corps import"  }  │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│  silver_faces.vbs  (consomme la prédiction)                  │
│                                                              │
│  - Affichage du rapport avec scores de confiance            │
│  - Application de la stratégie A / B selon la recommandation│
│  - Journalisation de l'acceptation/rejet → boucle retour    │
└─────────────────────────────────────────────────────────────┘
```

---

## 5. La boucle de retour — Comment le modèle s'améliore

Chaque fois qu'un utilisateur **accepte ou rejette une correction**, cela constitue une étiquette d'entraînement.

La macro VBS enregistre ces retours dans un fichier CSV :

```csv
surface_name, aire_mm2, compacite, rapport_aspect, type_surface, prof_corps, pred_modele, decision_expert
Face.217, 0.0043, 0.0021, 847.3, HybridShapeDatum, 4, sliver, accepte
Face.043, 0.0071, 0.412, 3.2, HybridShapeFillet, 2, sliver, rejete
```

Après 300 à 500 corrections enregistrées, le modèle est réentraîné sur ce jeu de données grandissant. Au fil du temps, il apprend les **normes géométriques propres au domaine industriel**.

---

## 6. Apprentissage actif — Rendre l'étiquetage efficace

L'apprentissage actif demande aux experts d'étiqueter uniquement les **prédictions les plus incertaines** :

$$\text{Incertitude} = 1 - \max_k P(\text{classe}_k \mid \mathbf{x})$$

Quand $P = 0{,}51$ pour une sliver, le modèle est maximalement incertain — ces cas sont mis en file d'attente pour une revue humaine. Quand $P = 0{,}99$, le modèle est confiant — aucune vérification n'est nécessaire. Cela réduit l'effort d'étiquetage de 70 à 80 % tout en atteignant la même précision.

---

## 7. Feuille de route d'implémentation

| Étape | Action                                                                                           | Horizon                |
| ----- | ------------------------------------------------------------------------------------------------ | ---------------------- |
| **1** | Modifier le VBS pour extraire 10 caractéristiques par surface et journaliser en CSV              | Maintenant (1–2 jours) |
| **2** | Collecter 100+ corrections étiquetées en usage réel                                              | ~2–4 semaines          |
| **3** | Entraîner un Isolation Forest en Python, valider les scores d'anomalie                           | Après l'étape 2        |
| **4** | Entraîner un Random Forest avec validation croisée à 5 plis                                      | Après l'étape 2        |
| **5** | Construire un endpoint Python minimal (Flask) lisant features.json et retournant les prédictions | ~1 semaine             |
| **6** | Modifier le VBS pour écrire features.json et lire predictions.json (Shell + FileSystemObject)    | ~1 jour                |
| **7** | Implémenter la journalisation de la boucle de retour                                             | ~1 jour                |
| **8** | Après 500+ exemples : explorer XGBoost, ajouter les caractéristiques de courbure                 | ~2 mois                |
| **9** | Si données > 5 000 pièces : expérimenter avec un GNN sur le graphe d'adjacence                   | Long terme             |

---

## 8. Métriques d'évaluation clés

En contrôle qualité, la fonction de coût est **asymétrique** :

| Métrique                      | Formule                             | Cible                                                             |
| ----------------------------- | ----------------------------------- | ----------------------------------------------------------------- |
| **Rappel** (priorité absolue) | $\dfrac{VP}{VP+FN}$                 | > 0,97 (manquer une sliver est coûteux)                           |
| **Précision**                 | $\dfrac{VP}{VP+FP}$                 | > 0,85 (les fausses alertes mobilisent inutilement l'expert)      |
| **Score F2**                  | $\dfrac{5 \cdot P \cdot R}{4P + R}$ | > 0,94 (pénalise les faux négatifs 2× plus que les faux positifs) |
| **Calibration** (ECE)         | Erreur de calibration attendue      | < 0,05 (les scores de confiance doivent être fiables)             |

Le score F2 pénalise explicitement les faux négatifs davantage que les faux positifs — ce qui est exactement l'exigence du contrôle qualité en fabrication.

---

## Synthèse

La transition du VBS actuel vers un système augmenté par le ML repose sur quatre évolutions parallèles :

1. **Extraction de caractéristiques enrichies** — de 1 caractéristique (l'aire) à 10+ (compacité, rapport d'aspect, courbure, type, contexte arborescent)
2. **Décisions probabilistes plutôt que binaires** — des seuils codés en dur aux scores de confiance calibrés
3. **Recommandation de stratégie** — de 2 corrections automatiques + 3 manuelles à 5 stratégies automatisées guidées par un classifieur multi-classe
4. **Apprentissage continu** — chaque correction d'expert améliore la prédiction suivante

L'algorithme le plus impactant à court terme est l'**Isolation Forest** (non supervisé, sans données étiquetées, déployable en quelques jours), et l'algorithme le plus puissant à long terme est un **Réseau de Neurones sur Graphes** qui capture les relations topologiques entre surfaces — ce qui distingue fondamentalement une silver face d'une petite surface géométriquement valide.
