# ReadMe extractv2.py

## How to Use the script

1. place all the pdf’s you need to scan into the input folder - *make sure that they are the correct **token** pdf format.* 
2. place the  helper sheets you need into the input folder as well. **Make sure they are EXACTLY formatted the following way and named the following way too.** 
    1.  **table** with the instrument id for each instrument: name = **instrument_ids**__
    
    | instrument id | best name |
    | --- | --- |
    | InstrumentID | TokenName |
    
    **b.  lookup table** between best id and client id in tdx : name = ***id vs best***
    
    | best_id | key |
    | --- | --- |
    | Integer | Client_ID |
    
     **c**. The export file out of tdx **:** name = ***import tdx***
    
    Create the **Sum** column which is the sum of the columns : Available, Credit, Withdrawalsreserved, OrdersReserved 
    
    Move the ClientAccountID column to the become the first column
    
    | ID | ClientAccountID | SubAccountID | InstrumentID |
    | --- | --- | --- | --- |
    
    | Available | Credit | WithdrawalsReserved | OrdersReserved |  Sum |
    | --- | --- | --- | --- | --- |

## Python

Python version : `Python 3.11.9`

### Dependencies

This script requires the following Python libraries:

- os
- fitz (PyMuPDF)
- pandas
- openpyxl

Required installs (first time only)

`pip3 install os`

`pip3 install PyMUPDF`

`pip3 install pandas`

`pip3 install openpyxl`


1. create the “input“ folder *
2. create the “output” folder *
3. Place both `output_folder`  and `input_folder` in the same directory as the  `extractv2.py`
4. in powershell navigate to the `extractv2.py` folder `cd enter/file/path/`
    1. for example `documents/thomas/sizeautomate` ← folder where `extractv2.py`is stored
5. run `python3 extractv2.py` in PowerShell

*If you want to name the folders something else, go and change the `input_folder` and `output_folder` variables in the source code.
## Common mistakes

You will most likely get an error from either misnaming headers inside your sheets or misnaming the sheets. Refer to “How to use the script” for correct labelling. 
- The column names will be output to the console when one of them is mislabelled
- Make sure the input file format is -.xlsx for the input data (other than pdf) <- If they are incorrectly named you will get a "One of the input files did not load" error.

# Excel

1. Navigate to the very last sheet called summary
2. Follow the instructions on the page
    1. Copy paste the formulas into the cells above (dont copy the quotation marks)
    2. Delete the instruction text below
    3. drag down the cells all the way until the end
    4. filter out the NA’s from ClientAccountID and filter out the 1 from best id
