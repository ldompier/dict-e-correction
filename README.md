# Dictée diagnostique 4e année

Outils TypeScript pour traiter les résultats d'une dictée diagnostique :

1. **Colorier un tableau Word existant** (cases rouges = erreurs)
2. **Remplir un template de notes** (scores par élève avec pourcentages)

---

## Fichiers CSV requis

### `mapping_v2.csv` — référentiel des mots
3 lignes, séparées par `;` :
```
ligne 1 : labels affichés (les mots : imprudent, impression, grand, …)
ligne 2 : codes          (1, 2, 3a, 3b, 10a, 10b, 11a, 11b, …)
ligne 3 : niveaux        (3 ou 4 pour chaque mot)
```

### `mapping.csv` — version simplifiée (sans niveaux)
2 lignes seulement (labels + codes). Utilisé par `colorise_dictee.ts` uniquement.

### `compilation_mot.csv` — résultats des élèves
```
ligne 1 : eleve;1;2;3a;3b;…   (en-tête avec codes)
ligne 2+ : <n_eleve>;x;;x;…   (x = erreur, vide = correct)
```

---

## Scripts

### `shade_and_fill_notes.ts` ⭐ (script principal)
Prend un tableau Word existant, le colorie **et** remplit le template de notes en une seule commande.

```bash
npx ts-node shade_and_fill_notes.ts \
  --src tableau_original.docx \
  --mapping mapping_v2.csv \
  --compil compilation_mot.csv \
  --out Tableau_colore.docx \
  --notes-template Notes_eleves.docx \
  --notes-out Notes_eleves_renseigne.docx
```

Produit :
- `Tableau_colore.docx` — tableau avec cases rouges là où l'élève a fait une erreur
- `Notes_eleves_renseigne.docx` — scores par élève (total 3e, 4e, global + %) avec ligne Moyenne

Options supplémentaires :
| Option | Défaut | Description |
|--------|--------|-------------|
| `--header-row-index` | `1` | Index (0-basé) de la ligne des mots dans le tableau |
| `--student-row-offset` | `1` | Élève 1 → ligne `1 + 1 = 2` |
| `--first-data-col-index` | `1` | Colonne 0 = numéro élève, colonnes 1+ = mots |
| `--notes-start-row-index` | `1` | Première ligne de données dans le template notes |
| `--notes-student-col-index` | `0` | Colonne du numéro d'élève dans le template notes |
| `--scores` | `correct` | `correct` = compte les réussites, `errors` = compte les erreurs |

---

### `shade_docx_in_place.ts`
Colorie uniquement le tableau (sans remplir les notes). Utilise `mapping.csv` (2 lignes).

```bash
npx ts-node shade_docx_in_place.ts \
  --src tableau_original.docx \
  --mapping mapping.csv \
  --compil compilation_mot.csv \
  --out tableau_colore.docx
```

---

### `colorise_dictee.ts`
Génère un tableau Word **from scratch** (sans document source). Utilise `mapping.csv`.

```bash
npx ts-node colorise_dictee.ts \
  --mapping mapping.csv \
  --compil compilation_mot.csv \
  --out tableau_colore.docx
```

---

## Structure du tableau source (`tableau_original.docx`)

Le tableau doit respecter cette structure :

| (ligne 0) | *(optionnel : titre ou vide)* |
|-----------|-------------------------------|
| (ligne 1) | en-tête avec les mots, dans l'ordre des codes de `mapping.csv` ligne 2 |
| (ligne 2) | élève 1 : numéro en col 0, cases vides en col 1+ |
| (ligne 3) | élève 2 |
| … | … |

---

## Installation

```bash
npm install
```

Dépendances principales : `docx`, `jszip`, `xmldom`  
Dev : `ts-node`, `typescript`, `@types/node`, `@types/xmldom`
