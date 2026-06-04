# Protection Relay Test Documentation

Test results templates and biennial check sheets for routine maintenance testing of protection relays at Contact Energy sites.

## Relay Coverage

### Test Results Templates (9-tab xlsx)

| File | Relay | Manufacturer | Type |
|------|-------|-------------|------|
| `SEL-300G_Test_Results_Template.xlsx` | SEL-300G | SEL | Generator |
| `SEL-311L_Test_Results_Template.xlsx` | SEL-311L | SEL | Line diff + distance |
| `SEL-787_Test_Results_Template.xlsx` | SEL-787 | SEL | Transformer (2-winding) |
| `SEL-700G_Test_Results_Template.xlsx` | SEL-700G | SEL | Generator |
| `REG670_Test_Results_Template.xlsx` | REG670 | Hitachi Energy (Relion 670) | Generator |
| `RET670_Test_Results_Template.xlsx` | RET670 | Hitachi Energy (Relion 670) | Transformer |
| `REF615_Test_Results_Template.xlsx` | REF615 | ABB (Relion 615) | Feeder |
| `RET615_Test_Results_Template.xlsx` | RET615 | ABB (Relion 615) | Transformer |
| `REM615_Test_Results_Template.xlsx` | REM615 | ABB (Relion 615) | Motor |
| `REU615_Test_Results_Template.xlsx` | REU615 | ABB (Relion 615) | Voltage/frequency |
| `GE-889_Test_Results_Template.xlsx` | GE Multilin 889 | GE Grid Solutions | Generator |

### Biennial (2-Yearly) Visual Check Sheets

| File | Covers |
|------|--------|
| `SEL_Biennial_Visual_Check.xlsx` | SEL-300G, SEL-311L, SEL-787, SEL-700G |
| `HIT_670_Biennial_Visual_Check.xlsx` | REG670, RET670 |
| `ABB_615_Biennial_Visual_Check.xlsx` | REF615, RET615, REM615, REU615 |
| `GE-889_Biennial_Visual_Check.xlsx` | GE Multilin 889 |

### Check Sheets (.docx)

| File | Relay |
|------|-------|
| `Doc files/SEL-300G_Test_Check_Sheet.docx` | SEL-300G |
| `Doc files/SEL-311L_Test_Check_Sheet.docx` | SEL-311L |
| `REG670_Test_Check_Sheet.docx` | REG670 |

## How It Works

Each results template has a **Cover sheet** where you enter CT/VT ratios and tolerances. These drive named formulas across the workbook — the analogue inputs tab automatically calculates expected primary values and pass/fail results without manual calculation.

All files carry the Contact Energy logo and print footer, applied by `add_branding.py`.

See `CLAUDE.md` for full technical documentation and the workflow for adding new relay types.
