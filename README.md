# Sensor-Data-Processing

Data Cleaning/Processing Files For Rack Greenhouse

## process_month.py

Combines daily temperature sensor TXT log files for one month into a single formatted Excel workbook. Each file in the input folder represents one day of readings; the script cleans out repeated header rows, extracts the relevant sensor columns (T01–T06), inserts per-row average formulas, and writes a final month-wide average summary row at the bottom.

### Requirements

Python 3 and openpyxl. Install the library with:

```
pip install openpyxl
```

### Setup

Open `process_month.py` in any text editor. At the very top you will see two lines — change these to match your folder and desired output file:

```python
INPUT_FOLDER = r"USB0\2025-11"   # path to the folder containing the TXT files for the month
OUTPUT_FILE  = r"November.xlsx"  # name (and optional path) for the output Excel file
```

- `INPUT_FOLDER` should point to the folder that contains the `.txt` files for **one month only**. Use a raw string (`r"..."`) so that backslashes in Windows paths are handled correctly.
- `OUTPUT_FILE` is where the finished workbook will be saved. If you give just a filename (e.g. `November.xlsx`) it will be created in the same folder as the script. You can also give a full path (e.g. `r"C:\Users\you\Documents\November.xlsx"`).

### Running

Open a terminal in the folder that contains `process_month.py` and run:

```
py -3 process_month.py
```

You will see output like:

```
Reading TXT files from: C:\...\USB0\2025-11
Loaded 2881 data rows  (2025-11-01 -> 2025-11-30)
Writing: C:\...\November.xlsx
Done.
```

The script will exit with a clear error message if the input folder cannot be found or contains no valid data.

### Processing a different month

To process October, December, or any other month, just update the two config lines:

```python
INPUT_FOLDER = r"USB0\2025-10"
OUTPUT_FILE  = r"October.xlsx"
```

Then re-run the script.

### Input format

- The input folder must contain files named exactly `RDL_YYYY-MM-DD_USB0.txt`, one per day.
- Files are processed in alphabetical order, which is the same as chronological order given the naming convention.
- Each file may contain repeated column-header rows (the logger writes one at the start of every 30-minute block). These are automatically skipped.
- Sensor columns T07 and T08 are excluded from the output because they frequently contain invalid readings.

### Output format

The output is a single `.xlsx` workbook with one sheet named after the month (e.g. `November`).

| Column | Content |
|--------|---------|
| A | Date (`mm-dd-yy`) |
| B | Time of reading (`h:mm:ss`) |
| C | T01 |
| D | T02 |
| E | `=AVERAGE(C:D)` — *Under Modules inside Pots* |
| F | T03 |
| G | T04 |
| H | `=AVERAGE(F:G)` — *Under Modules* |
| I | T05 |
| J | T06 |
| K | `=AVERAGE(I:J)` — *Control* |

- **Row 1** — header row. Columns A and B show the timestamp of the first log entry; columns C–K show the sensor and group labels.
- **Rows 2 to N** — one row per sensor reading.
- **Last row** — month-wide column averages (`=AVERAGE(C2:C{N})` etc.), with no date or time values.
