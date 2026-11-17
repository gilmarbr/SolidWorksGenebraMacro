<p align="center">
  <img src="docs/geneva.png" width="160" alt="Geneva wheel">
</p>

# SOLIDWORKS Geneva Mechanism Macro

A VBA macro for **SOLIDWORKS 2024** that designs and builds a complete external Geneva drive from a few parameters: the Geneva wheel, the crank (drive wheel with locking disk and pin), the base plate, the shafts, and a fully mated assembly with built-in animation.

**Authors:** Gilmar Barbosa and Max Barbosa · **Autom**

> The user interface and messages are in Portuguese. A Portuguese guide is in [`GenebraMacro/LEIA-ME.md`](GenebraMacro/LEIA-ME.md).

## Features

- **Parametric design:** size the mechanism by center distance (C), crank radius (r) or wheel outer diameter, with 3 to 18 slots and 1 or 2 pins.
- **Live results:** geometry, indexing and dwell times, angular velocity and acceleration, plus errors and warnings. They update on every change, before anything is generated.
- **Validation:** checks for slot, locking-arc, hub, keyway, relief and shaft interferences.
- **Fork tip thickness:** sets the thinnest wall between each slot and its locking arc, so the wheel horns don't end in a fragile sharp tip.
- **Generated files:** wheel, crank, base and shaft parts, the assembly, a kinematics CSV (`Cinematica.csv`) and step-by-step Motion Study instructions (`LEIA-ME_Motion.txt`).
- **Options:** shaft bores, keyways and hubs on the wheel and crank; disk- or arm-shaped crank plate; base plate with rounded corners and mounting holes.
- **Materials and colors:** SOLIDWORKS library material per part, with optional colors. The default is no color, which keeps the material's own appearance.
- **Animation:** a slider to set the crank angle, and an **Animate** button that drives the wheel and crank with exact Geneva kinematics. It works in any SOLIDWORKS edition, no Motion add-in needed.
- **Presets:** save and load all parameters to `GenebraPreset.ini`.

## Repository contents

| Path | Description |
|---|---|
| `Genebra.swp` | Ready-to-run macro |
| `GenebraMacro/modGenebra.bas` | Main module: calculation, validation, part and assembly creation, mates, animation, CSV export |
| `GenebraMacro/clsEvento.cls` | Event class for the controls created at run time |
| `GenebraMacro/frmGenebra_codigo.txt` | UserForm code, to paste into an empty UserForm |
| `GenebraMacro/frmGenebra.frm` / `.frx` | Exported UserForm (use `frmGenebra_codigo.txt` as the reference source) |
| `GenebraMacro/Autom_Logo.bmp` | Logo shown in the form footer |
| `GenebraMacro/LEIA-ME.md` | Full guide in Portuguese |

The source files contain **ASCII characters only**. Accented strings are written as `U("Sa\u00EDda")` and decoded at run time, so file encoding doesn't matter when importing or copying.

## Quick start

To run the macro as it is, use **Tools > Macro > Run**, select `Genebra.swp` and choose `main`.

## Building from source

1. In SOLIDWORKS, go to **Tools > Macro > New…** and save as `Genebra.swp`. The VBA editor opens.
2. **Insert > UserForm**, then in the Properties window (F4) set **(Name)** to `frmGenebra`. This also adds the *Microsoft Forms 2.0* reference.
3. **File > Import File…**: import `clsEvento.cls`, then `modGenebra.bas`. Import them instead of pasting, because the `Attribute`/`VERSION` header lines only work on import.
4. Open the `frmGenebra` code (F7), delete everything, and paste the full contents of `frmGenebra_codigo.txt`.
5. Delete the auto-created module (`Module1` or similar), because it has its own `main`.
6. Under **Tools > References**, make sure these are checked:
   - *SldWorks 2024 Type Library*
   - *SOLIDWORKS 2024 Constant type library*
   - *Microsoft Forms 2.0 Object Library*
7. Run **Debug > Compile VBAProject**, then save.

Default part and assembly templates must be set in *Tools > Options > System Options > Default Templates*. You can also enter the template paths in the **Saída** (Output) tab.

## Usage

1. Fill in the **Geral** (General), **Roda** (Wheel), **Manivela** (Crank) and **Base e eixos** (Base and shafts) tabs, and check the **Resultados** (Results) tab.
2. In **Saída** (Output), choose the output folder and click **GERAR** (Generate).
3. With the assembly open, use the **Movimento** (Motion) tab to position or animate the mechanism. **Parar** (Stop) or **Esc** stops the animation.

## How the assembly works

- The **base is fixed**. It holds the reference axes `EIXO_R` / `EIXO_M` and planes `PL_RODA` / `PL_MAN`.
- The wheel and the crank each get a coincident axis mate and a coincident plane mate, which leaves only rotation free.
- Standard mates can't describe the intermittent Geneva motion, since the pin enters and leaves the slots. Motion comes from two sources:
  1. the macro's own animation (exact kinematics);
  2. a SOLIDWORKS Motion Study with a rotary motor on the crank and solid-body contact, following the generated `LEIA-ME_Motion.txt`.

## Limitations

- External Geneva only, with the pin integral to the crank (no roller pin).
- No fillets or chamfers on slot edges. Add them manually if needed.
- Slot depth is always the maximum safe depth, plus the bottom clearance.
- Material names must match the SOLIDWORKS material library exactly. If a name isn't found, the macro only logs a warning.

## Troubleshooting

If something fails, the last line of `%TEMP%\Genebra_log.txt` shows the step where it stopped. See the troubleshooting table in [`LEIA-ME.md`](GenebraMacro/LEIA-ME.md) for common issues: missing templates, flipped parts, animation not updating, rotation direction.
