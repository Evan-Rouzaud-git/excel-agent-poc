# Demo Guide

This document contains the demo dataset, the exact prompts used in the Excel walkthrough, and the recommended order for reproducing the showcase.

## Demo Dataset

The dataset is intentionally imperfect so the demo can show audit, formatting, formula, and charting behavior on the same workbook.

```csv
Mois,Client,ChiffreAffaires,Cout,Region,Statut,Segment,Commentaire
Jan,Atlas,12000,7000,IDF,OK,Enterprise,Initial sale
Feb,Beacon,13500,7900,Ouest,OK,SMB,Renewal
Mar,Coda,14800,8600,IDF,OK,Enterprise,Upsell
Apr,Delta,0,5000,Sud,Review,SMB,Zero revenue anomaly
May,Echo,15600,9200,IDF,OK,Enterprise,Strong month
Jun,Flux,14200,N/A,Ouest,Review,SMB,Malformed cost value
Feb,Beacon,13500,7900,Ouest,OK,SMB,Renewal
Aug,Gamma,,8300,IDF,Missing,Enterprise,Missing revenue
Sep,Halo,16400,9400,IDF,OK,Enterprise,Peak season
Oct,Iris,15100,8700,Ouest,OK,SMB,Healthy margin
Nov,Jade,14600,9100,IDF,OK,Enterprise,Lower margin
Dec,Kite,17000,9600,IDF,OK,Enterprise,Year-end close
```

The table contains:

- one exact duplicate row
- one missing revenue value
- one malformed cost value
- one zero-revenue anomaly

That is enough to show that the agent can inspect workbook content before it changes it.

## Why the Dataset Is Imperfect

The data is designed to create visible results in Excel:

- the audit step has real issues to find
- the formatting step has a table worth styling
- the formula step has a clear calculated column
- the chart step has a sensible time series to visualize

## Demo Prompts

The prompts are intentionally in French because the original target environment used French Excel and French-speaking business users.

### Prompt 1

```text
Vérifie les données manquantes, anomalies et doublons dans le tableau.
```

Technical goal:

- Exercise the `validate_data` macro.
- Show source inspection before any modification.
- Produce an Issues sheet and highlight the problematic cells.
- Trigger the confirmation flow that follows the audit.

What the reviewer should see in Excel:

- a new Issues table
- highlighted anomalies in the source table
- visible logs in the add-in

### Prompt 2

```text
Mets le tableau au format corporate_blue avec un en-tête visible, des colonnes ajustées et une présentation propre.
```

Technical goal:

- Exercise the controlled formatting preset.
- Show that formatting goes through a macro, not manual UI styling.
- Demonstrate header styling and automatic column sizing.

What the reviewer should see in Excel:

- a blue corporate header
- readable table formatting
- adjusted columns

### Prompt 3

```text
Ajoute une colonne Marge égale à ChiffreAffaires moins Cout.
```

Technical goal:

- Exercise structured formula writing.
- Show that the add-in can create a derived column without manual formula entry.
- Demonstrate deterministic execution with workbook-aware column placement.

What the reviewer should see in Excel:

- a new `Marge` column on the right of the table
- formulas filled down the table
- values updating consistently row by row

### Prompt 4

```text
Crée un graphique du ChiffreAffaires par Mois et place-le à droite du tableau.
```

Technical goal:

- Exercise chart creation from workbook context.
- Show that the add-in maps columns into a chart through a controlled macro.
- Demonstrate output placement relative to the source table.

What the reviewer should see in Excel:

- a chart placed to the right of the table
- `Mois` on the horizontal axis
- `ChiffreAffaires` as the plotted series

## Recommended Demo Order

1. Run the audit prompt first.
2. Apply the `corporate_blue` preset second.
3. Add the `Marge` column third.
4. Create the chart last.

This order shows the intended control flow: inspect first, then format, then compute, then visualize.

## Demo Video

Demo video: <link to be added>
