# Mutual Funds Performance Query

Python library for extracting certain performance parameters of Mutual Funds in India from https://moneycontrol.com

![Python](https://img.shields.io/badge/python-v3.6+-blue.svg)
# [![Pypi](https://img.shields.io/badge/pypi-v1.8-green)](https://pypi.python.org/pypi/mftool)
![Dependencies](https://img.shields.io/badge/dependencies-up%20to%20date-brightgreen.svg)
![License](https://img.shields.io/pypi/l/selenium-wire.svg)

Introduction
============
Were you given a set of Mutual Funds as suggestion to invest? Are you already invested in few mutual funds?

Run this script to keep track of their performances and key metrics to help you decide on buying or holding or redeeming your investments.

Features
=============

* Get performance and key metrics for any Mutual Fund in India on-demand.
* Input can be a single mutual fund or multiple through an Excel file.

Instructions
=============
1. Download the package.
2. Launch terminal session from the root folder of the package.
3. Install requirements `pip3 install -r requirements.txt`
4. For single mutual fund, obtain the URL of the requisite mutual fund from https://www.moneycontrol.com/mutualfundindia/ .
5. For multiple mutual funds, build an Excel (.xlsx) workbook with the following. See Sample.xlsx for sample excel workbook.
    1. Sheetname should be Mutual Funds
    2. Column 1 with the URL(s) of requisite mutual funds from https://www.moneycontrol.com/mutualfundindia/ . 
    3. Row 1 with headers from column 2 onwards for writing the data in respective cells. Order of headers is not significant.
        1. Category
        2. CRISIL Star Rank
        3. Risk Rating
        4. TER % vs Avg
        5. 3 YR Return
        6. 3 YR Rank / Cat Funds
        7. 5 YR Return
        8. 5 YR Rank / Cat Funds
        9. Since Inception Return
        10. Since Inception Rank / Cat Funds
        11. Updated On
6. Run the script `python3 MFQuery.py`
7. Choose 1 if providing a direct mutual fund URL. Output will be presented on the terminal session itself.
8. Choose 2 if providing a Excel workbook with mutual fund URL(s). Output will be written to the same Excel workbook.