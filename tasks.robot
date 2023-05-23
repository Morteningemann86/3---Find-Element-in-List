*** Settings ***
Documentation       Check if name is in main customer database. If it is a new customer, the bot will process the new customer data

Library             RPA.Excel.Files
Library             RPA.Tables


*** Tasks ***
Main
    ${list_main}    Excel to list    customers_main.xlsx    Email
    ${list_new}    Excel to table    customers_new.xlsx
    FOR    ${element}    IN    @{list_new}
        Log To Console    ${element}[First Name]
        IF    "${element}[E-mail Address]" not in ${list_main}
            Process New Customer    ${element}
        END
    END


*** Keywords ***
Excel to list
    [Arguments]    ${filepath}    ${column_name}

    Open Workbook    ${filepath}
    ${table}    Read Worksheet As Table    header=True
    Close Workbook
    ${list}    Get Table Column    ${table}    ${column_name}

    RETURN    ${list}

Excel to table
    [Arguments]    ${filepath}

    Open Workbook    ${filepath}
    ${table}    Read Worksheet As Table    header=True
    Close Workbook

    RETURN    ${table}

Process New Customer
    [Arguments]    ${element}
    Log To Console    Customer with the name of ${element}[First Name] ${element}[Last Name] name with address ${element}[Address] has been added to the customer database
