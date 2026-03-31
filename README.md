# Field Layout Plot Randomizer

Simple browser-based layout generator for randomized field/treatment plans.

## How to use

1. Open `http://localhost/field-layout-randomizer/` in your browser (via XAMPP).
2. Paste entries in the format:
   - `CODE,check` or `CODE,test`
   - or `CODE check` / `CODE test`
3. Set field `rows`, `columns`, number of `replications`, and the layout order:
   - `Row-col order` or `Column serpentine`
4. Choose the starting corner (`Top left`, `Top right`, `Bottom left`, `Bottom right`).
5. Click **Generate layout**.
6. Click **Download Excel** to export the randomized assignments with `X` and `Y` coordinates.

## Output

The Excel includes one row per plot cell per replication (unused plots are blank).

