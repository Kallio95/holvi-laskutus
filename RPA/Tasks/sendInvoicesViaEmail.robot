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
    FOR    ${i}    IN RANGE    2    ${rowCount}+1
        ${lname}=  Read Excel Cell  row_num=${i}  col_num=1    sheet_name=Sheet1
        ${fname}=  Read Excel Cell  row_num=${i}  col_num=2    sheet_name=Sheet1
        ${email}=  Read Excel Cell  row_num=${i}  col_num=3    sheet_name=Sheet1
        ${sum}=  Read Excel Cell  row_num=${i}  col_num=4    sheet_name=Sheet1
        ${msg}=  Read Excel Cell  row_num=${i}  col_num=5    sheet_name=Sheet1
        ${subject}=  Read Excel Cell  row_num=${i}  col_num=6    sheet_name=Sheet1
        ${category}=  Read Excel Cell  row_num=${i}  col_num=7    sheet_name=Sheet1
        IF    $lname is None or $email is None
            Log To Console    \nRivi ${i} ohitettu: nimi tai sähköposti puuttuu
            CONTINUE
        END
        ${name}=    Evaluate    ' '.join(str(p) for p in ($lname, $fname) if p)
        Fill Invoice Details    ${name}    ${email}    ${sum}    ${msg}    ${subject}    ${category}
        Log To Console    \n${name} lähetetty!\nYht: ${i}/${rowCount}
    END

*** Keywords ***
Initiate Holvi
    # Holvi UI language follows the browser; locators expect Finnish
    Open Browser    https://login.app.holvi.com   chrome    options=add_experimental_option("prefs", {"intl.accept_languages": "fi"})
    Wait Until Element Is Visible    id=email    20s
    ${cookieBanner}=    Run Keyword And Return Status    Element Should Be Visible    id=onetrust-accept-btn-handler
    IF    ${cookieBanner}    Click Button    id=onetrust-accept-btn-handler
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
    Wait Until Element Is Visible    id=invoice.receiver.name-label    10s
    # Fill in receiver name
    Click Element    xpath=//*[@id="invoice-edit-form"]/section[1]/div/div/div[1]/div[1]/div/div/div/div
    Press Keys    None    ${name}
    Sleep   1s
    # Select email as delivery method
    Press Keys    None    TAB    TAB    TAB    ARROW_DOWN    ENTER
    Capture Page Screenshot    screenshot-${name}.png
    # Fill in receiver email
    Input Text    id=invoice.receiver.email    ${email}
    Capture Page Screenshot    screenshot-${name}-email.png
    Input Text    id=subject    ${subject}
    Input Text    id=items.0.description    ${msg}
    IF    $category
        Click Element    css=[aria-owns="listbox-new-payment-category-select-id-0"]
        ${option}=    Set Variable    xpath=//ul[@id='listbox-new-payment-category-select-id-0']//span[normalize-space()='${category}']
        Wait Until Element Is Visible    ${option}    5s
        Click Element    ${option}
    END
    Input Text    id=items.0.quantity    1
    Input Text    name=items.0.detailed_price.gross    ${sum}
    Scroll Element Into View    xpath=//button[normalize-space()='Lähetä nyt']
    Click Element    xpath=//button[normalize-space()='Lähetä nyt']
    Sleep    3s
