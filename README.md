# Purpose
    Combine FCT files from branch to one file

# Install packages

Only need to run once
- run `pip install -r requirements.txt` in terminal


# Run the app to combine files

1. Copy all files need to combine to `input` folder
2. Copy master data file to `template` folder. Make sure all workbook links are removed.
    `Open Excel -> Data -> Queries & Connections -> Workbook Links -> Remove all`
3. Edit list of colums to copy value to master file (`app.py` from line 18)
4. Run `python app.py` in terminal or run `app.py` in VSCode
5. The output file will be in `output` folder with name `%Y%m%d %H%M%S Topline FCT (combined).xlsx`
6. Check logs folder for more information


# Logs

Each run will create a log file in `logs` folder, which contains 2 files:
- `%Y%m%d %H%M%S.log`: contains runtime log
- `%Y%m%d %H%M%S ERROR.log`: contains errors log or warnings during runtime