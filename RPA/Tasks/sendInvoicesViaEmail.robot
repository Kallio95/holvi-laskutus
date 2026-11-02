*** Settings ***
Library   SeleniumLibrary
Library   ExcelLibrary
Suite Setup    Initiate Holvi
Suite Teardown    Close Holvi
Variables    ../../env.json

*** Tasks ***
TestExcel
    Open Excel Document    ${FILE}    1
    ${columnData}=    Read Excel Column    col_num=1 
    ${rowCount}=    Evaluate    len(${columnData})
    Log To Console    \nRivien lkm: ${rowCount}
    Sleep    10s
    FOR    ${i}    IN RANGE    2    ${rowCount}
        ${lname}=  Read Excel Cell  row_num=${i}  col_num=1    sheet_name=Sheet1
        ${fname}=  Read Excel Cell  row_num=${i}  col_num=2    sheet_name=Sheet1
        ${email}=  Read Excel Cell  row_num=${i}  col_num=3    sheet_name=Sheet1
        ${sum}=  Read Excel Cell  row_num=${i}  col_num=4    sheet_name=Sheet1
        ${msg}=  Read Excel Cell  row_num=${i}  col_num=5    sheet_name=Sheet1
        ${subject}=  Read Excel Cell  row_num=${i}  col_num=6    sheet_name=Sheet1
        ${category}=  Read Excel Cell  row_num=${i}  col_num=7    sheet_name=Sheet1
        ${name}=    Catenate    ${lname}    ${fname}
        Fill Invoice Details    ${name}    ${email}    ${sum}    ${msg}    ${subject}    ${category}
        Log To Console    \n${name} lähetetty!\nYht: ${i}/${rowCount}
    END

*** Keywords ***
Initiate Holvi
    Open Browser    https://login.app.holvi.com   chrome
    Wait Until Element Is Visible    id=onetrust-accept-btn-handler    20s
    Sleep    1s
    Click Button    id=onetrust-accept-btn-handler
    Input Text    id=email    ${USERNAME}
    Input Password    id=password    ${PASSWORD}
    Click Button    id=submit

Close Holvi
    Close All Browsers
    Close All Excel Documents

Fill Invoice Details
    [Arguments]    ${name}    ${email}    ${sum}    ${msg}    ${subject}    ${category}
    # Wait for login to go through
    Wait Until Element Is Visible    id=account.invoicing    60s
    # Go to new invoice page
    Click Element    id=account.invoicing
    Sleep    1s
    Click Element    id=sidebar-invoicing-new
    Wait Until Element Is Visible    id=receiver.name-label    10s
    # Fill in receiver name
    Click Element    xpath=//*[@id="invoice-edit-form"]/section[1]/div/div/div[1]/div[1]/div/div/div/div
    Press Keys    None    ${name}
    Sleep   1s
    # Select email as delivery method
    Press Keys    None    TAB    TAB    TAB    ARROW_DOWN    ENTER
    Capture Page Screenshot    screenshot-${name}.png
    # Fill in receiver email
    Input Text    id=receiver.email    ${email}
    Capture Page Screenshot    screenshot-${name}-email.png
    Input Text    id=subject    ${subject}
    Input Text    id=items.0.description    ${msg}
    Select From List By Label    id=new-payment-category-select-id-0    ${category}
    Input Text    id=items.0.quantity    1
    Input Text    name=items.0.detailed_price.gross    ${sum}
    Scroll Element Into View    xpath=//button[normalize-space()='Lähetä nyt']
    Click Element    xpath=//button[normalize-space()='Lähetä nyt']
    Sleep    3s
