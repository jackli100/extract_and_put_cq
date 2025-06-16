# Area Sum Utility

This repository contains sample data (`0615.xlsx`) and small scripts for computing the total of column F. The new `calc_area.py` script is minimal:

```bash
python calc_area.py
```

First install the requirements:

```bash
pip install -r requirements.txt
```

`calc_area.py` reads `0615.xlsx`, cleans up the sixth column (column F) and converts any numeric strings to floats. Cells containing non-numeric text are ignored. The script then prints the total sum of the numeric values.
