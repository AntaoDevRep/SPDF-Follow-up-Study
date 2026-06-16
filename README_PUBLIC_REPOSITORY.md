# SPDF follow-up study analysis code

This repository contains the public-facing analysis script for generating Tables 1–4 and Figures 1–4 source-data/validation outputs for the SPDF follow-up manuscript.

## Important data note

Raw patient-level clinical data are not included in the public repository because they contain sensitive health information and are subject to institutional and ethical restrictions. To rerun the full pipeline locally, place the restricted input files under `data_private/follow_up/`. The `data_private/` folder is intentionally excluded from version control.

## Expected repository structure

```text
R/                         cleaned helper functions
scripts/                   main analysis script
data_private/follow_up/    restricted input data; do not commit
outputs/main_analysis/     generated outputs; do not commit by default
templates/ppt/             optional PowerPoint templates
```

## Required helper files

The public script expects cleaned helper files in `R/`:

- `functions_intergroup_tests.R`
- `functions_global.R`
- `themes.R`
- `functions_survival.R`

These should define at least `build.df.summary()`, `save.fig.to.pptx()`, `draw.survial.plot()`, `m.normal.theme`, and `m.top.legend.theme`.

## Running the analysis

Run the script from the repository root:

```r
source("scripts/SPDF_followup_main_analysis_public_repository.R")
```

By default, patient-level internal/model dataset exports are disabled:

```r
export.internal.datasets <- FALSE
export.model.datasets <- FALSE
```
