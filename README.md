## Report: Scraping metro.de shopping website

Install the requirements (`splinter` - web automation, `openpyxl`, `pandas` - eases work with excel files,  and `flet` - GUI) using:

```bash
pip install -r requirements.txt
```

### GUI Capture
TBA

### How to try?
TBA

[//]: # (Download the files for your OS from the releases section.)

### Notes

- Running `utils.py` will run the automation, but without showing a GUI. If you want the GUI, then run `gui.py` ([flet](https://flet.dev) is used for it, so it should be installed).

- Required files if not in GUI mode: `source.xlsx`, `gui.py`, `utils.py`, `logs.txt`

- The project must contain a source file with named `source.xlsx` (hardcoded), which contains at least two columns with headers `"Metro Artikelnummer"` and `"Link"` (equally hardcoded). The code could be modified though to work without the first column, but the second is absolutely necessary(contains the direct link to each item).

- A new Excel file named `Results.xlsx` will be created at the end of the automation/scraping and contains the scraped results.


Made with ❤ by TheEthicalBoy!