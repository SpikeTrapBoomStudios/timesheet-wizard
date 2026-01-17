An app that builds timesheets based on raw time data.

**NOTE:** If you are on Linux, then `pip install -r requirements.txt` will fail. This is because this app utilized a windows-only python package to build a PDF from the exported xlsx file. However, the app still works on Linux, just without this feature. **To build on linux, use the command:** `cat requirements.txt | xargs -n 1 pip install`
