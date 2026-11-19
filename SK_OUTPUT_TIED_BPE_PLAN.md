> **SUPERSEDED (2026-11-19, Robbie): do not execute.** Replaced by the KEEI 2025 alignment staged on `develop_4.0.7`. BPCiObIC now follows KEEI 2025 output and value added (not the BIFUbC energy index). BPE process CO2 for cement, steel and chemicals follows KEEI physical production (clinker, BOF steel, basic petrochemicals). The PERAC cement multiplier is relinked to BPE. BGDP follows KEEI GDP plus the semiconductor mega-cluster. See `handoffs/keei2025-alignment/flow-logic-2026-11-19.md`.

# PLAN (DRAFT FOR STAFF REVIEW): tie SK BPE, PERAC and BPCiObIC to the corrected BIFUbC trajectory

> **STATUS: DRAFT. Not signed off.** Step 1 (BIFUbC) is done and uncommitted. Steps 2-5 wait for answers to the open decisions below.
> Model: eps-southkorea `develop_4.0.7` @ f712d29. This is an input for staff review, not an EI position. Verify all figures against the files and sources before use.

## Step 1: BIFUbC mapping fix (DONE, uncommitted)
Workbook `InputData/indst/BIFUbC/BAU Industrial Fuel Use before CCS.xlsx`, applied per the colleague's note (KEEI row 62 is 비금속, non-metallic minerals):
- On the 8 fuel sheets (electricity, coal, natural-gas, LPG, petroleum-diesel, heavy or residual oil, biomass, heat):
  - Row 83 (glass 231 + cement 239) growth now uses row 62 (was Others, row 66). 27 cells per sheet: `$O$62`, `$K$62`, `$L$62`, `$M$62`.
  - Row 85 (other metals 242 + metal products 25) now uses row 64, Machinery (was row 62). 27 cells per sheet.
  - Labels renamed: A62, H45, A83 to `Non-metallic Minerals (비금속)`; H47, A85 to `Machinery`.
- `KEEI BAU - Industry` A15, A87:A92 and `Others Split` B8 renamed (the `0` prefix on A87 was kept).
- 8 CSVs re-exported as LF (`BIFUbC-*` for LPG, biomass, coal, electricity, heat, heavy-or-residual-oil, natural-gas, petroleum-diesel). Only the 4 rows 231, 239, 242, 25 changed in each; every other row is equal by value. `crude-oil` and `hydrogen` are unchanged (still CRLF).
- Result for cement 239, all ten fuels incl. electricity (TBtu): 2021 169.4 (unchanged); 2024 184.0 to **154.8**; 2030 195.6 to **175.9**; 2040 160.0 to **175.5**; 2050 134.9 to **170.9**. Index vs 2024: 1.136 / 1.133 / 1.104 (2030/40/50); vs 2021: 1.038 / 1.036 / 1.009.
- Source of the 2024 drop: the 2023-to-2024 bridge factor `O62` for coal is 0.79 (Others was 1.08).
- Original workbook is backed up in the session scratchpad and equals the committed version.

## Step 2: BPCiObIC (output), for approval
Replace the US AEO x IRA growth with the BIFUbC index, as in the earlier handoff, on a shared `Energy Index` tab (one row per ISIC, `=SUM` of the ten fuels).
- 239: `E(t)/E(2021) - 1`, which gives 0.0 / +3.8% / +3.6% / +0.9% in 2021/2030/2040/2050.
- Open: do 231, 242 and 25 also take their own BIFUbC index? Their fuel growth changed in Step 1, so consistency says yes.
- Do not touch `Rescale to Fit GDP Proj`. The D54 GDP-fit residual will move; no other ISIC row moves while D52 stays typed at 0.954 (re-check if 242/25 are in the non-AEO group).
- Paste E(t) as values with a source note and recompute from the new CSVs.

## Step 3: BPE, everything on output except chemicals and F-gases
Current drivers (BPE workbook):
| Series | Driver now |
|---|---|
| CO2, `Process CO2 Growth` rows 59-67 | ag and water/waste: UN population; coal mining, pipelines, refining: CH4 trend; chemicals, steel, other mfg: ETC production TREND; cement: E(t) (my earlier edit) |
| CH4, N2O (`BPE-CH4`, `BPE-N2O`) | EPA non-CO2 projections by country x multiplier (chemicals CH4 uses the CO2 chemicals row) |
| VOC, CO, NOx, PM10, PM25, SOx, BC, OC (`Calcs`) | US AEO23 sector growth ('Growth rates') plus diesel |
| F-gases | EPA/ClimateDeck/AIM path |
Proposed: add an `Output Index` tab, `OI(t) = (1 + BPCiObIC[ISIC,t]) / (1 + BPCiObIC[ISIC,2024])` for 2025-2050 (1 for 2020-24, as today), matched to each BPE industry by its ISIC code, pasted from the new `BPCiObIC.csv` with a source note. Replace the growth driver with OI for every industry x gas except chemicals 20 (all gases) and F-gases. Then re-export the BPE CSVs.
- Open: does "everything" include ag, water/waste, coal mining, O&G/pipelines (supply- or population-driven today) and the non-GHG pollutants?
- BAU only: EGGRA = 0, so the model's GDP adjustment adds no second scaling.
- Gate: chemicals and F-gas rows equal by value to f712d29.

## Step 4: PERAC (`indst/PERAC/Proc Emis Reductions and Costs.xlsx`)
- BPE's About says PERAC must use the same multipliers and adjustments as BPE. Per the 2026-11-16 SK F-gas log, the PERAC `Multipliers and Adjustments` block E23:AI50 is pasted from the legacy BPE (not the template peg). That is the existing bug.
- Plan: read PERAC first (read-only) to see how the cement potential (8.3 / 9.1 / 10.4 / 11.75 Mt) is built, then relink PERAC's multiplier block to the updated BPE, and rescale cement so the potential keeps its share of the new BPE.
- PERAC changes already on the branch (commit 2fd5f59, scripts in `fgas-perac-handoff/`; I read the script headers, not the chat itself):
  - F-gas per-source time-series multipliers are computed in a new BPE tab `F-gas Source Multipliers` and pasted into the PERAC `Multipliers and Adjustments` tab; row 37 is set to 1 so it cannot stack.
  - Korea duplicate-copy correction: EPA lists Korea twice, so all EPA-derived PERAC quantities were 2x; a KOR = 0.5 factor is applied in `EPA_MACCs` column I.
  - The new multiplier pattern (BPE computes, PERAC reads) is the one to reuse for non-F-gas rows. F-gases stay out of the output scaling, as you said.
- **Please confirm** that these two changes are the whole of the PERAC changes from the colleague review chat, or paste anything else.

## Step 5: other IO files (checked; read-only)
- **No change needed (base-year levels, scaled in-model by BPCiObIC):** `BObIC`, `BVAbIC`, `BEbIC`, `BECbIC` (EPS.mdl 11188, 14920, 5834). Value added, jobs and compensation follow output automatically.
- **Check only:** `BPCiDCSbIC` (domestic content share % change, same workbook), `LPGRbIC` (labor productivity growth, used at EPS.mdl 25425/34435/34571). Confirm their 231/239/242/25 rows are not derived from the AEO output growth being replaced.
- **Static, no change:** `BDCSoCbIC`, `FoGPbEaIC`, `GaHEbIC`, `URPbIC`, `BEaCP`, `BGDP` and the other IO files.

## Superseded work to clean up
- The earlier cement-only edits (E(t) tab, row 64, row 16, both CSVs) used the OLD anchor. They are no longer uncommitted: commit 2fd5f59 (the f-gas correction) already contains them, so the committed BPE-CO2.csv and BPCiObIC.csv carry the old-anchor cement values (29.20 / 30.72 / 21.19 Mt). `git checkout` would not remove them. Proposal: overwrite them forward in Steps 2-3 (replace the `Cement Energy Index` tab and rows 64 / 16). Do not push the branch as-is.
- The draft note to the PERAC owner used the old BPE numbers; do not send it.

## Gates (per step, max three)
1. Target values reproduced in a headless BAU run. 2. No-change control (chemicals and F-gases, other ISICs, all fuel use outside the four ISIC rows). 3. LF only, BAU reaches 2050.

## Open decisions
1. Step 2 scope: 239 only, or 231/239/242/25?
2. Step 3 scope: all industries and all gases except chemicals and F-gases, including ag/waste/mining and the non-GHG pollutants?
3. Index peg year for BPE: 2024 (as now).
4. Confirm the PERAC changes listed in Step 4 are complete.
5. Approve overwriting the committed cement-only edits forward (no `git checkout`; they are in 2fd5f59).
6. Leftovers from Step 1: `crude-oil` sheet (separate layout; cement crude = 0), `KEEI BAU - Industry` B87, and the `Non-ferrous Metals Split` sheet name and B8 (now covering Machinery-driven rows).
