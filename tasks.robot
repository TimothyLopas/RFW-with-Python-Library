*** Settings ***
Documentation       Template robot main suite.

Library             RPA.Browser.Playwright    timeout=00:00:30    auto_closing_level=SUITE    run_on_failure=Take Screenshot \ EMBED
Library             RPA.Robocorp.Vault
Library             RPA.Cloud.AWS
Library             RPA.Salesforce
Library             TeamsMessages.py
Library             RPA.Excel.Files
Library             RPA.Tables
Library             String


*** Variables ***
${CASE_FILE_NAME}=          Review Cases.xlsx
${AWS_DOWNLOAD_BUCKET}=     robocorp-test


*** Test Cases ***
Collect cases and append property tax details
    Authenticate to S3
    Authenticate to Salesforce
    Download file from S3 bucket    ${AWS_DOWNLOAD_BUCKET}    ${CASE_FILE_NAME}
    ${cases}=    Open Excel file and extract cases
    Open Google Maps webpage
    Collect property tax details, append to Salesforce case and send Teams message    ${cases}


*** Keywords ***
Authenticate to S3
    ${secret}=    Get Secret    aws
    Init S3 Client
    ...    ${secret}[AWS_KEY_ID]
    ...    ${secret}[AWS_KEY]
    ...    ${secret}[AWS_REGION]

Authenticate to Salesforce
    ${sf_secret}=    Get Secret    salesforce
    Auth With Token    ${sf_secret}[api_username]    ${sf_secret}[api_password]    ${sf_secret}[api_token]

Download file from S3 bucket
    [Arguments]    ${bucket_name}    ${file_name}
    @{file_list}=    Create List    ${file_name}
    Download Files    ${bucket_name}    ${file_list}    ${OUTPUT_DIR}
    Log    AWS S3 file download complete

Open excel file and extract cases
    Open Workbook    ${OUTPUT_DIR}${/}${CASE_FILE_NAME}
    ${cases_table}=    Read Worksheet As Table    Sheet1    header=${TRUE}
    ${cases_column}=    Get Table Column    ${cases_table}    Cases
    RETURN    ${cases_column}

Open Google Maps webpage
    New Browser    headless=${FALSE}
    New Page    https://www.google.com/maps

Collect property tax details, append to Salesforce case and send Teams message
    [Arguments]    ${cases}
    ${secret}=    Get Secret    salesforce
    ${base_url}=    Set Variable    ${secret}[base_url]
    FOR    ${case}    IN    @{cases}
        ${address}    ${case_id}=    Find mailing address from case number    ${case}
        # ${latest_year_taxes}    ${assessed_value}=
        Search For propery    ${address}
        Append property image to case notes    ${address}    ${case_id}
        ${teams_message}=    Set Variable
        ...    Case ${case} has been updated on SalesForce: https://${base_url}.lightning.force.com/lightning/r/Case/${case_id}/view
        Send Message To Sfdc Messages Channel    ${teams_message}
    END

Find mailing address from case number
    [Arguments]    ${case_number}
    ${case_query}=
    ...    Salesforce Query Result As Table
    ...    SELECT Id, ContactId FROM Case WHERE CaseNumber = '${case_number}'
    ${case_id}=    Set Variable    ${case_query}[0][0]
    ${contact_id}=    Set Variable    ${case_query}[0][1]
    ${contact_query}=
    ...    Salesforce Query Result As Table
    ...    SELECT Id, MailingAddress FROM Contact WHERE Id = '${contact_id}'
    ${mailing_address_dict}=    Set Variable    ${contact_query}[0][1]
    ${mailing_address}=    Set Variable
    ...    ${mailing_address_dict}[street]${SPACE}${mailing_address_dict}[city],${SPACE}${mailing_address_dict}[state]${SPACE}${mailing_address_dict}[postalCode]
    RETURN    ${mailing_address}    ${case_id}

Search for propery
    [Arguments]    ${address}
    ${address_upper}=    Convert To Upper Case    ${address}
    ${address_filename}=    Replace String    ${address_upper}    ${SPACE}    _
    Fill Text    id=searchboxinput    ${address_upper}
    Keyboard Key    press    Enter
    Sleep    5
    Take Screenshot    ${OUTPUT_DIR}${/}${address_filename}    fullPage=False

Append property image to case notes
    [Arguments]    ${address}    ${case_id}
    ${address_upper}=    Convert To Upper Case    ${address}
    ${address_filename}=    Replace String    ${address_upper}    ${SPACE}    _
    ${caseFeed_data}=
    ...    Create Dictionary
    ...    ParentId=${case_id}
    ...    Body=${address}
    ...    Title=${address}
    ${caseFeed}=    Create Salesforce Object    FeedItem    ${caseFeed_data}
    ${binary_file_data}=    Evaluate    open("output/${address_filename}.png", 'rb').read()
    ${base64_encoded_data}=    Evaluate    base64.encodebytes($binary_file_data).decode('utf-8')
    ${contentVersion_data}=
    ...    Create Dictionary
    ...    VersionData=${base64_encoded_data}
    ...    PathOnClient=${OUTPUT_DIR}${/}${address_filename}.png
    ...    FirstPublishLocationId=0058c000009XrFN
    ...    Origin=H
    ...    ContentLocation=S
    ${contentVersion}=    Create Salesforce Object    ContentVersion    ${contentVersion_data}
    ${caseAttachment_data}=
    ...    Create Dictionary
    ...    Type=Content
    ...    FeedEntityID=${caseFeed}[id]
    ...    RecordId=${contentVersion}[id]
    Create Salesforce Object    FeedAttachment    ${caseAttachment_data}
    Log    Case Notes Updated

Delete all case comments
    ${casecomment_query}=
    ...    Salesforce Query Result As Table
    ...    SELECT Id, CommentBody, ParentId FROM CaseComment
    FOR    ${row}    IN    @{casecomment_query}
        ${casecomment_id}=    Set Variable    ${row}[Id]
        Delete Salesforce Object    CaseComment    ${casecomment_id}
    END
