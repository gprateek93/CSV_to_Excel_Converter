# CSV_to_Excel_Converter
This script can be used to convert CSV files to Excel files 

## Requirements
It is assumed that you have the latest version of python with openpyxl module installed.
<br>You can install the latest python version from the following link:
<br>
https://docs.anaconda.com/anaconda/install/

## Working setting
The script is designed for a specific setting:
1. This code runs if there are csv files inside subfolders of the root folder. No csv file should be directly present under the root folder. It must be present in some subfolder.
2. Top four rows in the csv file is text and after that each subsequent row contains float values.

If your working setting is different, you can fork the repo and reuse the code as per your requirement.

## To run the code:
```
python convert.py <root folder path>
```
