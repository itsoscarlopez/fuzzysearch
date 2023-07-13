# fuzzysearch
This file is a python based script that uses the OpenPyXL, sys, fuzzyWuzzy, re and pickle libraries to create a name duplicate algorithm for a .xlsm document.


The "all_names" file is a Python pickle output that contains a list of 3 items:
    (1) all account first names
    (2) all account middle names
    (3) all accout last names
All these names are as written on the Excel document they were outsourced from.


The "all_accounts" file is another Python pickle output that contains a list of
the account information of every name found in the Excel file. Each is a class
called "Account()" (created in the duplicate_check.py file) that contains the 
following 7 attributes:
    (1) Excel index
    (2) list of name matches
    (3) First Name
    (4) Middle Name
    (5) Last Name
    (6) Full Name
    (7) Excel Name*

*The excel name differs from First, Middle, Last, and Full names in that it
contains ALL whitespaces and punctuations.
    EXAMPLE: Your Account Name is Ritchie* A. Field
        (1) Excel index             - 1882
        (2) list of name matches    - [None]
        (3) First Name              - Ritchie
        (4) Middle Name             - A
        (5) Last Name               - Field
        (6) Full Name               - RitchieAField
        (7) Excel Name              - Ritchie* A. Field

