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
    Wait Until Element Is Visible    id=sidebar-invoicing    60s
    Click Element    id=sidebar-invoicing
    Sleep    1s
    Click Element    id=sidebar-invoicing-new    
    Wait Until Element Is Visible    id=receiver.name-label    10s
    Press Keys    None    TAB    TAB
    Press Keys    None    ${name}
    Sleep   1s
    Input Text    id=receiver.email    ${email}
    Input Text    id=subject    ${subject}
    Input Text    id=items.0.description    ${msg}
    Select From List By Label    id=new-payment-category-select-id-0    ${category}
    Input Text    id=items.0.quantity    1
    Input Text    name=items.0.detailed_price.gross    ${sum}
    Scroll Element Into View    xpath=//button[normalize-space()='Lähetä']
    Click Element    xpath=//button[normalize-space()='Lähetä']
    Sleep    2s
    Click Link    xpath=//a[normalize-space()='Lähetä sähköpostitse']
    Sleep    3s
