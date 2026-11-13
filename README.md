# Macro Excel – Visualisation du foisonnement

Macro VBA pour relire plus facilement une matrice de traçabilité **exigences amont → exigences avales → preuves attendues** où le foisonnement est représenté sur plusieurs lignes (valeurs amont / aval répétées sur chaque ligne).

La macro crée une feuille **« Revue foisonnement »** à partir de la feuille active, sans modifier la source :

- les cellules amont et aval répétées sont fusionnées en un seul bloc ;
- une bande de couleur par exigence amont, trait épais entre exigences amont, trait fin entre exigences avales ;
- **rouge** : exigence amont sans exigence aval ;
- **orange** : exigence aval sans preuve.

Relancer la macro régénère la feuille.

## Contenu

| Fichier | Description |
|---|---|
| `src/Foisonnement.bas` | Module VBA commenté |
| `exemple/Analyse_Fonctionnelle_Circuit_test_foisonnement.xlsx` | Fichier d'exemple : onglet source `Matrice tracabilite` et onglet `Revue foisonnement` généré par la macro |

## Format attendu

| Amont (ID, libellé…) | Aval (ID, libellé…) | Preuve |
|---|---|---|
| F-001 | EXT-01 | Rapport d'essai |
| F-001 | EXT-01 | Courbe constructeur |
| F-001 | EXT-02 | Note de calcul |
| F-002 | | |

- Une ligne d'en-tête, puis une ligne par lien amont / aval / preuve.
- Le regroupement se fait sur la première colonne de chaque bloc (l'identifiant).

## Installation

1. Dans Excel : `Alt+F11` → **Fichier › Importer un fichier** → `src/Foisonnement.bas`.
2. Adapter les constantes en haut du module (A=1, B=2…) :

   | Constante | Rôle | Défaut |
   |---|---|---|
   | `LIGNE_ENTETE` | ligne des titres | 1 |
   | `COL_AMONT_DEB` / `COL_AMONT_FIN` | colonnes amont | 1 / 2 (A–B) |
   | `COL_AVAL_DEB` / `COL_AVAL_FIN` | colonnes aval | 3 / 4 (C–D) |
   | `COL_PREUVE_DEB` / `COL_PREUVE_FIN` | colonnes preuve | 5 / 5 (E) |
   | `MODE_FUSION` | `True` fusionne les répétitions, `False` les grise (garde le filtrage) | `True` |
   | `TRIER` | `True` trie par amont puis aval avant traitement (si les lignes d'une même exigence ne se suivent pas) | `False` |

3. Se placer sur la feuille source, `Alt+F8` → `VisualiserFoisonnement`.
4. Enregistrer le classeur au format `.xlsm`.

## Licence

MIT – voir [LICENSE](LICENSE).
