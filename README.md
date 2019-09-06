# Excel

Excel is a small library of Excel utility functions compiled for personal needs. There's 
nothing too fancy nor anything you can't find from another library, but Excel consists of
smaller functions to be used rather than relying on larger packages.

These functions include things like to CSV, to JSON, to rows, checking if Excel file paths are 
valid and getting all Excel files from a directory.

## Personal Note

Excel is only on Github because I reference it in other projects. I don't have any plans 
to maintain this project, but I will update it from time to time. 

# Install

You can install this project directly from Github via:

```bash
$ pip3.7 install git+https://github.com/kelmore5/python-excel-utilities.git
```

# Usage

Once installed, you can import the main class like so:

    >>> from kelmore_excel import ExcelTools as Excel
    >>>
    >>> path_to_excel_file: str = '/home/username/Downloads/some_excel_file.xlsx'
    >>>
    >>> Excel.path.files('/home/username/Downloads')    # ['some_excel_file.xlsx']
    >>> Excel.path.is_valid(path_to_excel_file)         # True
    >>> Excel.transform.to_rows(path_to_csv_file)       # [['first name', 'last name'], ['kyle', 'elmore']]
    >>> Excel.transform.to_json(path_to_csv_file)       # [{'first name': 'kyle', 'last name': 'elmore'}]
    .
    .
    .

# Documentation

To be updated
