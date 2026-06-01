# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

This repository is a documentation system for routine and maintenance testing of protection relays. It produces two paired deliverables per relay:

- **Check sheet** (`.docx`) — printed/electronic procedure followed during the test visit; filled in by the tester.
- **Results template** (`.xlsx`) — live workbook capturing measured values, with auto pass/fail calculation and a rolled-up overall outcome.

The **Specification** (`Protection_Relay_Test_Documentation_Specification.docx`) is the authoritative structural reference. The **Onboarding Form** (`New_Relay_Onboarding_Form.xlsx`) captures per-relay parameters and drives generation of new pairs.

## Reference Implementations

| Relay | Type | Files |
|-------|------|-------|
| SEL-300G | Generator protection | `SEL-300G_Test_Check_Sheet.docx` + `SEL-300G_Test_Results_Template.xlsx` |
| SEL-311L | Line current differential & distance | `SEL-311L_Test_Check_Sheet.docx` + `SEL-311L_Test_Results_Template.xlsx` |
| ABB REF615 | Feeder protection | `REF615_Test_Results_Template.xlsx` |
| ABB RET615 | Transformer protection | `RET615_Test_Results_Template.xlsx` |
| ABB REM615 | Motor protection | `REM615_Test_Results_Template.xlsx` |
| ABB REU615 | Voltage/frequency protection | `REU615_Test_Results_Template.xlsx` |

## Workflow for Creating a New Relay Pair

See `workflow for creating a checksheet.docx` for the full workflow. The recommended approach:

1. Fill in `New_Relay_Onboarding_Form.xlsx` for the new relay (ratios, element list, binary I/O labels, relay-specific blocking checks).
2. Provide the Specification, the filled Onboarding Form, and the closest existing pair (SEL-311L for line relays; SEL-300G for generator relays) to an LLM.
3. Prompt: *"Generate the `<Relay>_Test_Check_Sheet.docx` and `<Relay>_Test_Results_Template.xlsx` for `<relay model>` following the attached Specification, using the Onboarding Form parameters and the provided pair as a worked example."*
4. Review and iterate.

## Document Structure

### Check Sheet (.docx) — fixed section order

| Section | Content | Varies per relay? |
|---------|---------|-------------------|
| Cover | Logo, title, relay subtitle | Subtitle only |
| 1 — Test Information | Site, personnel, relay under test (CT/VT ratios), equipment, drawings, SharePoint links | Ratio rows in 1.3 |
| 2 — Safety & Permits | Standard LOTO checklist + isolation record | Append relay-specific blocking checks |
| 3 — Pre-Test Checks | Comms, status, clock, event records, SOE, as-found settings, settings vs PFD | No |
| 4 — Analogue Input Verification | Primary-value injection tables (one sub-section per CT/VT group) | Yes — sub-sections vary |
| 5 — Binary I/O | Input pickup + output continuity tables | Prefilled function labels vary |
| 6 — Protection Element Testing | Omicron test structure, element tracker, per-element result template | Element list varies |
| 7 — Post-Test Checks | Re-compare settings, clear events, time sync, restore plant | Restoration items vary |
| 8–10 — Findings, Files Index, Sign-Off | Defect log, artefact index, tester/witness/engineer sign-off | No |

**Key design principle (Section 4):** Always verify the relay reports the correct **PRIMARY** value (injected secondary × ratio), not just the secondary value. This catches ratio-setting errors that secondary-only checks miss.

### Results Template (.xlsx) — 9 tabs

| Tab | Mirrors section | Live calculation |
|-----|----------------|-----------------|
| 1. Cover | Front matter + relay info | Ratios drive analogue formulas; tolerances drive P/F; auto roll-up |
| 2. Equip & Refs | Equipment + drawings + locations | Clickable hyperlinks |
| 3. Pre-Test | Section 3 | OK/FAIL conditional formatting |
| 4. Analogue Inputs | Section 4 | Expected Primary, Mag Err %, Angle Err°, P/F — all formulas |
| 5. Binary IO | Section 5 | PASS/FAIL conditional formatting |
| 6. Element Tracker | Section 6.3 | PASS/FAIL |
| 7. Element Results | Section 6.4 | Pickup err %, Time err %, Result — formulas |
| 8. Post-Test | Section 7 | OK/FAIL |
| 9. Files & Sign-Off | Sections 8–10 | Overall result pulled from Cover |

### Formula Architecture (xlsx)

**Layer 1 — Defined names on Cover sheet** (workbook-scoped, readable names):
- `Ratio_VTmain`, `Ratio_CTstator`, etc. → ratio cells on Cover
- `TolV_pct`, `TolI_pct`, `TolAngle_deg`, `TolPickup_pct`, `TolTime_pct` → tolerance cells (yellow fill)

**Layer 2 — Per-row in Analogue Inputs tab:**
- Expected Primary = `injected_secondary × Ratio_<applicable>`
- Mag Err % = `(Relay Primary − Expected Primary) / Expected Primary × 100`
- Angle Err° = `Relay Angle − Injected Angle`
- P/F = `IF(OR(ABS(mag_err) > TolV_pct, ABS(angle_err) > TolAngle_deg), "FAIL", "PASS")`
- All formulas wrapped in `IFERROR` + blank-input guards (no `#DIV/0!` on empty rows)

**Layer 3 — Cover roll-up:**
- Analogue: `COUNTIF` for FAIL/PASS across Analogue Inputs tab
- Elements: `COUNTIF` across Element Results tab
- Auto overall: FAIL if either is FAIL; PASS if both PASS; else INCOMPLETE
- The **manual dropdown** (PASS / PASS WITH NOTES / FAIL) on Cover is the authoritative signed result; auto overall is a cross-check only

## Branding & Style Conventions

| Element | Value |
|---------|-------|
| Brand colour | `#1F4E79` (deep blue) — section headers, h1 fills, table headers |
| Accent colour | `#D9E2F3` (light blue) — label cells |
| Note background | `#FFF2CC` (yellow) — caution callouts, tolerance input cells |
| Summary background | `#FFF8E7` (cream) — free-text summary boxes |
| PASS | Fill `#C6EFCE` / text `#006100` |
| FAIL | Fill `#FFC7CE` / text `#9C0006` |
| Page size | A4 portrait, 1 cm margins |
| Font | Calibri; body 11 pt; h1 14 pt bold white-on-blue; h2 12 pt bold blue |
| User-input cells (xlsx) | Blue text (RGB 0,0,255) |
| Auto-calculated cells (xlsx) | Black bold on grey/alt-row — tester does not edit |
| Tolerance cells (xlsx) | Yellow fill — editable setting, not data |
