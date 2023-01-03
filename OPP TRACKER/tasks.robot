*** Settings ***
Documentation       Opportunity Tracker

Library             RPA.Excel.Files
Library             RPA.Robocloud.Items
Library             RPA.HTTP
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Robocorp.Vault
Library             RPA.Desktop.Windows
Library             XML
Library             RPA.Word.Application
Library             String
Library             RPA.Outlook.Application
Library             DateLogic.py
Library             DateTime
Library             Collections

Suite Teardown      RPA.Outlook.Application.Quit Application
Task Setup          RPA.Outlook.Application.Open Application


*** Variables ***
${CalculatedDate}


*** Tasks ***
Opportunity Tracker
    TRY
        ${configlist}=    Read Config file
        Open Workbook    ${configlist}[0]    overwrite=True
        ${InputExcel}=    Read Worksheet As Table    header=True
        ${rowCount}=    Set Variable    2
        Opportunity Tracker main task    ${InputExcel}    ${configlist}    ${rowCount}
    EXCEPT
        Log    Unable to read Excel
    END


*** Keywords ***
Opportunity Tracker main task
    [Arguments]    ${InputExcel}    ${configlist}    ${rowCount}
    TRY
        FOR    ${row}    IN    @{InputExcel}
            Log    ${row}
            ${client data list}=    Getting Single Client Data    ${row}
            IF    "${client data list}[4]" == "Sales Dependency"
                Log    Proceed further
                ${CalculatedDate}=    Is_Weekend    ${client data list}[6]    ${client data list}[7]
                ${List Date}=    Converting Current,FollowUp and Calculated Dates
                ...    ${client data list}
                ...    ${configlist}
                ...    ${CalculatedDate}

                Writing Calculated date In Excel
                ...    ${configlist}
                ...    ${rowCount}
                ...    ${CalculatedDate}
                ...    ${List Date}
                IF    ${List Date}[2] == ${List Date}[1]
                    Log    Dates are equal
                    Checking Communication Type and sending mail
                    ...    ${configlist}
                    ...    ${client data list}
                    ...    ${rowCount}
                    ${rowCount}=    Evaluate    ${rowCount}+1
                    #Writing Status Completed In excel sheet    ${configlist}    ${rowCount}
                ELSE IF    ${List Date}[2] < ${List Date}[1]
                    Log    FollowUpdate date is greater than current date
                    ${rowCount}=    Evaluate    ${rowCount}+1
                ELSE
                    Log    FollowUpDate date is smaller than current date
                    IF    ${List Date}[0] == ${List Date}[2]
                        Checking Communication Type and sending mail
                        ...    ${client data list}
                        ...    ${configlist}
                        ...    ${rowCount}
                        ${rowCount}=    Evaluate    ${rowCount}+1
                        #Writing Status Completed In excel sheet    ${configlist}[0]    ${rowCount}
                    ELSE IF    ${List Date}[0] > ${List Date}[2]
                        Writing changed dates if followup<Calculated In excel
                        ...    ${configlist}
                        ...    ${rowCount}
                        ...    ${List Date}
                        ${rowCount}=    Evaluate    ${rowCount}+1
                    ELSE
                        Inconvinence Checking Communication Type and sending mail
                        ...    ${rowCount}
                        ...    ${configlist}
                        ...    ${client data list}
                        ${rowCount}=    Evaluate    ${rowCount}+1
                    END
                END
            ELSE
                Log    Stop
                Emptying Cell from excel    ${configlist}    ${rowCount}    ${CalculatedDate}
                ${rowCount}=    Evaluate    ${rowCount}+1
            END
        END
    EXCEPT
        Log    Unable to process the bot
    END

Read Config file
    ${TodayDate}=    Get Current Date
    Open Workbook    oppTrackConfig.xlsx
    ${Config}=    Read Worksheet As Table    header=True
    FOR    ${row}    IN    @{Config}
        ${InputFilePath}=    Set Variable    ${row}[Input Excel File Path]
        ${EnquiryEmailBody}=    Set Variable    ${row}[EnquiryEmailBody]
        ${FollowUpEmailBody}=    Set Variable    ${row}[FollowUpEmailBody]
        ${InconvinienceEnquiry}=    Set Variable    ${row}[InconvinienceEnquiry]
        ${InconvinienceFollowUp}=    Set Variable    ${row}[InconvinienceFollowUp]
        ${configlist}=    Create List
        Append To List
        ...    ${configlist}
        ...    ${InputFilePath}
        ...    ${EnquiryEmailBody}
        ...    ${FollowUpEmailBody}
        ...    ${TodayDate}
        ...    ${InconvinienceEnquiry}
        ...    ${InconvinienceFollowUp}
        RETURN    ${configlist}
    END

Getting Single Client Data
    [Arguments]    ${row}

    ${customer}=    Set Variable    ${row}[Customer]
    ${proposal}=    Set Variable    ${row}[Proposal]
    ${SalesRep}=    Set Variable    ${row}[Sales SPOC]
    ${SalesRepEmail}=    Set Variable    ${row}[Sales SPOC Email ID]
    ${Status}=    Set Variable    ${row}[Status]
    ${CommunicationType}=    Set Variable    ${row}[Communication Type]
    ${FollowUpDate}=    Set Variable    ${row}[Follow Up Date]
    ${ReminderFrequency}=    Set Variable    ${row}[Reminder Frequency]
    ${client data list}=    Create List
    Append To List
    ...    ${client data list}
    ...    ${customer}
    ...    ${proposal}
    ...    ${SalesRep}
    ...    ${SalesRepEmail}
    ...    ${Status}
    ...    ${CommunicationType}
    ...    ${FollowUpDate}
    ...    ${ReminderFrequency}
    RETURN    ${client data list}

Converting Current,FollowUp and Calculated Dates
    [Arguments]    ${client data list}    ${configlist}    ${CalculatedDate}
    ${CalculatedDateFiltered}=    Convert Date    ${CalculatedDate}    result_format=%d-%m-%Y
    ${FollowUpDateFiltered}=    Convert Date    ${client data list}[6]    result_format=%d-%m-%Y
    ${TodayFilteredDate}=    Convert Date    ${configlist}[3]    result_format=%d-%m-%Y
    ${List Date}=    Create List
    Append To List    ${List Date}    ${CalculatedDateFiltered}    ${FollowUpDateFiltered}    ${TodayFilteredDate}
    RETURN    ${List Date}

Writing Calculated date In Excel
    [Arguments]    ${configlist}    ${rowCount}    ${CalculatedDate}    ${List Date}
    TRY
        Open Workbook    ${configlist}[0]    overwrite=True
    EXCEPT
        Log    Writing Excel Option
    ELSE
        Set Cell Value    ${rowCount}    K    ${List Date}[0]
        Save Workbook
    END

checking Communication Type and sending mail
    [Arguments]    ${client data list}    ${configlist}    ${rowCount}
    IF    "${client data list}[5]" == "Enquiry"
        Log    ${client data list}[2]
        Log    ${client data list}[1]
        Enquiry Process    ${configlist}    ${client data list}
        Log    enquiry email should be sent
    ELSE IF    "${client data list}[5]" == "Follow Up"
        Log    ${client data list}[2]
        Log    ${client data list}[1]
        Follow up Process    ${configlist}    ${client data list}
        Log    follow up email should be sent
    ELSE
        Log    No communication needed
    END

Writing Status Completed In excel sheet
    [Arguments]    ${configlist}    ${rowCount}
    TRY
        Open Workbook    ${configlist}[0]    overwrite=True
    EXCEPT
        Log    Writing Excel Option
    ELSE
        Set Cell Value    ${rowCount}    F    Completed
        Save Workbook
    END

Writing changed dates if followup<Calculated In excel
    [Arguments]    ${configlist}    ${rowCount}    ${List Date}
    TRY
        Open Workbook    ${configlist}[0]    overwrite=True
    EXCEPT
        Log    Writing Excel Option
    ELSE
        Set Cell Value    ${rowCount}    L    ${List Date}[0]
        Save Workbook
    END

Inconvinence Checking Communication Type and sending mail
    [Arguments]    ${rowCount}    ${configlist}    ${client data list}
    Log    ${client data list}[5]
    IF    "${client data list}[5]" == "Enquiry"
        InconvinienceEnquiry Process    ${configlist}    ${client data list}
        Log    enquiry email should be sent
    ELSE IF    "${client data list}[5]" == "Follow Up"
        InconvinienceFollowUp Process    ${configlist}    ${client data list}
        Log    follow up email should be sent
    ELSE
        Log    No communication needed
    END

Emptying Cell from excel
    [Arguments]    ${configlist}    ${rowCount}    ${CalculatedDate}
    TRY
        Open Workbook    ${configlist}[0]    overwrite=True
    EXCEPT
        Log    Writing Excel Option
    ELSE
        Set Cell Value    ${rowCount}    K    ${EMPTY}
        Save Workbook
    END

Enquiry Process
    [Arguments]    ${configlist}    ${client data list}
    ${EnquiryEmailBody}=    Replace String    ${configlist}[2]    <SalesExecName>    ${client data list}[2]
    ${EnquiryEmailBody}=    Replace String    ${EnquiryEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[2]
    Log    ${client data list}[3]

    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${EnquiryEmailBody}

Follow up Process
    [Arguments]    ${configlist}    ${client data list}
    ${FollowUpEmailBody}=    Replace String    ${configlist}[2]    <SalesExecName>    ${client data list}[2]
    ${FollowUpEmailBody}=    Replace String    ${FollowUpEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[2]
    Log    ${client data list}[3]
    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${FollowUpEmailBody}

InconvinienceEnquiry Process
    [Arguments]    ${configlist}    ${client data list}
    ${EnquiryEmailBody}=    Replace String    ${configlist}[5]    <SalesExecName>    ${client data list}[2]
    ${EnquiryEmailBody}=    Replace String    ${EnquiryEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[5]
    Log    ${client data list}[3]

    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${EnquiryEmailBody}

InconvinienceFollowUp Process
    [Arguments]    ${configlist}    ${client data list}
    ${FollowUpEmailBody}=    Replace String    ${configlist}[5]    <SalesExecName>    ${client data list}[2]
    ${FollowUpEmailBody}=    Replace String    ${FollowUpEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[5]
    Log    ${client data list}[3]
    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${FollowUpEmailBody}
