*** Settings ***
Documentation     Template robot main suite.

Library    RPA.HTTP
Library    RPA.Tables
Library    RPA.Browser.Selenium
Library    RPA.PDF
Library    RPA.FileSystem
Library    RPA.Archive
Library    RPA.Dialogs
Library    RPA.Robocorp.Vault


*** Variables ***
${OrderexcelDownloadpath} =    BotData/Temp/orders.csv
${Orderexcelpath}=    BotData/Input/orders.csv
${OrderFileLink}=   https://robotsparebinindustries.com/orders.csv
#${WebsiteURL}=    https://robotsparebinindustries.com/#/robot-order

*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    
    ${Website}   Get URL from Vault
    Process Prechek
    Download the Excel file
    Move Order File
    ${Order_data}=    Read data from Excel file
    Open the robot order website   ${Website}
    
    FOR    ${row_data}    IN    @{Order_data}
        Order Form Polpulation  ${row_data}
        Order Submission
        Create Reciept   ${row_data}
    END
    
    [Teardown]   Log out and close the browser
    Create Zip file

    Display Output
    ${Feedback}=   Feedback form dialog
    Display Thanks Note   ${Feedback}
  

*** Keywords ***
Download the Excel file
    Download    ${OrderFileLink}    overwrite=True    target_file=${OUTPUT_DIR}${/}${OrderexcelDownloadpath}

Move Order File
    Move File    ${OrderexcelDownloadpath}    ${Orderexcelpath}   overwrite=true

Read data from Excel file
    ${Order_data}=    Read table from CSV     ${Orderexcelpath}    header=true
    [Return]   ${Order_data}

Open the robot order website
    [Arguments]    ${Website}
    Open Available Browser    ${Website}    maximized=true


Order Form Polpulation
    [Arguments]    ${row_data}
    Wait Until Page Contains Element  //button[@class="btn btn-dark"]
    Click Button    OK
    Wait Until Page Contains    Build and order your robot!    timeout=10 

    Select From List By Value    head    ${row_data}[Head]
    Select Radio Button    body    ${row_data}[Body]
    Set Focus To Element    css:input[type="number"]
    Input Text    css:input[type="number"]    ${row_data}[Legs]    clear=True
    Set Focus To Element    id=address
    Input Text    id=address    ${row_data}[Address]

    Click Button    Preview
    Wait Until Page Contains Element    id=robot-preview-image
    Sleep    2
    Screenshot    id=robot-preview-image     ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].png  

Order Submission   
    
    FOR    ${i}    IN RANGE    100
        #${ErrorMSG}=   Is Element Visible    //div[@class="alert alert-danger"]
        ${ErrorMSG}=  Is Element Visible  //div[@class="alert alert-danger"]  
        #${ErrorMSG}=      Element Should Be Visible    alert alert-danger
        IF    ${ErrorMSG}== True
            Click Button    Order  
        ELSE
            ${res}    Is Element Visible    //button[@id="order"]
            IF    ${res} == True
                Click Button    Order 
            ELSE
                Exit For Loop
            END
        END           
    END


Create Reciept
    [Arguments]    ${row_data}
    Wait Until Page Contains Element    id=receipt
    ${reciept_data}=     Get Element Attribute    id=receipt   outerHTML
    Html To Pdf    ${reciept_data}    ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].pdf
    Add Watermark Image To Pdf    ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].png    ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].pdf    ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].pdf
    Move File    ${OUTPUT_DIR}${/}BotData/Temp/${row_data}[Order number].pdf    ${OUTPUT_DIR}${/}BotData/Output/OrderReciepts/${row_data}[Order number].pdf   overwrite=true

    Click Button  //button[@id="order-another"]


Create Zip file
    Archive Folder With Zip    ${OUTPUT_DIR}${/}BotData/Output/OrderReciepts    ${OUTPUT_DIR}${/}BotData/Output/OrderReciepts.zip

Log out and close the browser
    Close Browser

Process Prechek  
    ${File}   Is Directory Empty    ${OUTPUT_DIR}${/}BotData/Input/
    IF    ${File} == False
        Remove Directory    ${OUTPUT_DIR}${/}BotData/Input   recursive=true     
    END
    Create Directory     ${OUTPUT_DIR}${/}BotData/Input

    ${File}   Is Directory Empty    ${OUTPUT_DIR}${/}BotData/Output
    IF    ${File} == False
        Remove Directory    ${OUTPUT_DIR}${/}BotData/Output  recursive=true
    END
    Create Directory     ${OUTPUT_DIR}${/}BotData/Output
    Create Directory     ${OUTPUT_DIR}${/}BotData/Output/OrderReciepts

    ${File}   Is Directory Empty    ${OUTPUT_DIR}${/}BotData/Temp
    IF    ${File} == False
        Remove Directory    ${OUTPUT_DIR}${/}BotData/Temp   recursive=true
    END
    Create Directory     ${OUTPUT_DIR}${/}BotData/Temp


Feedback form dialog
    Add heading       Feedback Form
    Add text input    name     label=Name
    Add text input    email    label=E-mail address
    Add text input    message
    ...    label=Feedback
    ...    placeholder=Enter feedback here
    ...    rows=5
    ${Feedback}=    Run dialog
    [Return]   ${Feedback.name}


Display Output
    Add icon      Success
    Add heading   Your process completed sucessfully
    Add files     ${OUTPUT_DIR}${/}BotData/Output/OrderReciepts.zip
    Run dialog    title=Success
    
Display Thanks Note
    [Arguments]   ${Feedback}
    Add icon      Success
    Add heading   Thanks ${Feedback} 
    Run dialog    title=Success

Get URL from Vault
    ${Secret}=   Get Secret      websitedata
    [Return]    ${Secret}[url]
    