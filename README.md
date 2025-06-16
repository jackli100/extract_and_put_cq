# Area Sum Utility

This repository contains sample data (`0615.xlsx`) and small scripts for computing the total of column F. The new `calc_area.py` script is minimal:

```bash
python calc_area.py
```

First install the requirements:

```bash
pip install -r requirements.txt
```

`calc_area.py` automatically converts the sixth column (column F) of `0615.xlsx`
to integers before computing the sum. Any cells that cannot be parsed as numbers
are ignored. Run the script after installing the requirements to see the total
area value printed to the console.
