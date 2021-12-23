*** Settings ***
Documentation     Insert the sales data for the week and export it as a PDF.
Library           RPA.Browser.Selenium    auto_close=${FALSE}
Library           RPA.HTTP
Library           RPA.Excel.Files
Library           RPA.PDF

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the Excel file
    Fill the form using the data from the Excel file
    Collect the results
    Export the table as a PDF
    [Teardown]    Log out and close browser

*** Keywords ***
Open the intranet website
    Open Available Browser    https://robotsparebinindustries.com/    maximize=true

Log in
    ${logout_visible}    Is Element Visible    logout
    IF    ${logout_visible} == ${TRUE}
        Click Button    logout
        Wait Until Page Contains Element    id:username
    END
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=true    target_file=${OUTPUT_DIR}${/}sales_data.xlsx

Fill and submit the form for one person
    [Arguments]    ${sales_reps}
    Input Text    firstname    ${sales_reps}[First Name]
    Input Text    lastname    ${sales_reps}[Last Name]
    Input Text    salesresult    ${sales_reps}[Sales]
    Select From List By Value    salestarget    ${sales_reps}[Sales Target]
    Click Button    Submit

Fill the form using the data from the Excel file
    Open Workbook    ${OUTPUT_DIR}${/}sales_data.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=true
    Close Workbook
    ${delete_all_sales_entries_exists}    Does Page Contain Button    xpath://*[@id="root"]/div/div/div/div[2]/div[3]/button[2]
    IF    ${delete_all_sales_entries_exists} == ${TRUE}
        Click Button    xpath://*[@id="root"]/div/div/div/div[2]/div[3]/button[2]
        Wait Until Page Does Not Contain Element    xpath://*[@id="root"]/div/div/div/div[2]/div[3]/button[2]
    END
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales_summary.png

Export the table as a PDF
    Wait Until Element Is Visible    id:sales-results
    Log    Storing HTML markup of sales results table
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}sales_results.pdf    overwrite=true

Log out and close browser
    Click Button    logout
    Close Browser
