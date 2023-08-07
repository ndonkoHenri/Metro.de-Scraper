## Report: Scraping metro.de shopping website

The code files in this directory are not cleaned and will not be further maintained nor updated. 
Simply because I moved FROM using normal selenium and later on selenium-base TO [Splinter](https://github.com/cobrateam/splinter). I find it faster. Also, I wanted to build a GUI app for it, and hence needed something which could be easily integrated without calling `os.system(...)` or similar.

**Note**: this is to be used only if running main_with_links.py or main3_withsbase.py because they make use of selenium-base

To run, use the below:

```bash
pytest main.py --demo
```

`--demo` here slows down the automation, making sure everything goes smoothly.

The project must contain a source file with named `source.xlsx` (hardcoded), which contains at least one column with header "`source`" (equally hardcoded). This column should contain the item numbers to be searched on the website.

A new Excel file named `Ergebnis.xlsx` will be created at the end of the automation and contains the results.


Made with ❤ by TheEthicalBoy!