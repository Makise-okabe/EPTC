# EPTC 2026 — TIM1 thermomechanical co-design

**Overleaf-ready LaTeX source for the EPTC 2026 technical report.**

## Compile / synchronize

1. In Overleaf, import the GitHub repository, or pull the latest `main` commit using your existing Git integration.
2. Set **Main document** to `main.tex`.
3. Select **pdfLaTeX** and compile. Bibliography is generated from `references.bib` using **BibTeX** (`unsrtnat`).
4. When this repository changes, synchronize/pull again. A GitHub push does not necessarily update an existing Overleaf project automatically.

The report figures (package cross-section, prescribed 700 W map, thermal DOE, matched nine-case response) are drawn with TikZ/PGFPlots. No local Windows file paths or proprietary solver binaries are required for LaTeX compilation.

## Files

- `main.tex` — complete report, figures and tables
- `references.bib` — bibliographic entries cited in the report
- `data/v1_material_screen.csv` — material-family thermal screening data
- `data/v3b_topology.csv` — modeled topology sensitivity
- `data/v4c_recovered_4case.csv` — recovered retained-contact shortlist (do not substitute the failed/incomplete export)
- `data/v4d_effective_brackets.csv` — mechanical-property sensitivity comparisons
- `data/v6a_9point_matched.csv` — matched nine-case thermal–mechanical DOE
- `.github/workflows/latex.yml` — automated PDF compilation / downloadable build artifact

### Evidence boundaries

These are **curated report tables**, not original AEDT/MAPDL solver files. Numerical outputs are user/project-generated simulations; literature papers support material context and explicitly labeled dimensional anchors. The 700 W tile map and chosen die geometry are **study assumptions**. The topology extreme contact-loss cases are hypothetical, not validated aging temperatures. The matched nine-case results use fixed assumed MID mechanics and fully bonded one-way thermal handoff. Local peak stress is **not converged** and must not be used for failure prediction. The preferred engineering result is a **two-candidate design window**, not a globally optimal or qualified product.

**Team:** Fang Tianchi, Ke Tangxin, Du Yanzhang, Duan Yihe

**Evidence freeze:** 23 October 2026

## Update policy

The revised report follows the organiser's required sequence: **Title Page & Team Info → Abstract → Problem Statement → Methodology → Results & Analysis → Discussion & Optimization → Conclusion & Recommendations → References**. The compiled report is **9 A4 pages**, including references, as of 26 October 2026.

### Pre-submission checks

- **Replace `[INSERT REGISTERED TEAM NAME]` and `[INSERT TEAM ID]` on the cover page with the team's actual registration details. Do not submit with these placeholders.**
- Check whether the submission portal expects an actual solver-generated Icepak temperature/hotspot contour. The report currently contains a clearly labeled prescribed 4×4 *heat-source map*, not a claimed solved temperature-field contour. Do not mislabel it.
- Compile `main.tex` and verify no PDF issues after replacing the placeholders.
- Keep the raw simulation exports available upon request; the `data/` folder contains curated comparison tables.

The current report source lives on `main`. Future changes should edit these same files and preserve the frozen source data, with corrections clearly recorded in Git history. Re-run the LaTeX build after changes and verify the compiled PDF page count is below ten pages before competition submission.
