# Purpose
    Combine FCT files from branch to one file

# Install packages

Only need to run once
- run `pip3 install -r requirements.txt` in terminal


# Run the app to combine files

1. Copy all files need to combine to `input` folder
2. Copy master data file to `template` folder
3. Run `python3 app.py` in terminal or run `app.py` in VSCode
4. The output file will be in `output` folder with name `%Y%m%d %H%M%S Topline FCT (combined).xlsx`
5. Check logs folder for more information


# Logs

Each run will create a log file in `logs` folder, which contains 2 files:
- `%Y%m%d %H%M%S.log`: contains runtime log
- `%Y%m%d %H%M%S ERROR.log`: contains errors log or warnings during runtime