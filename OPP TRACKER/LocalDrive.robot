*** Settings ***
Documentation       Opportunity Tracker

Library             RPA.Excel.Files
Library             RPA.Browser.Selenium    auto_close=${FALSE}
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
                IF    "${client data list}[8]" == "Yes"
                    Log    Proceed Further
                    ${FollowUpDateFiltered}
                    ...    ${TodayFilteredDate}=
                    ...    Converting Current,FollowUp and Calculated Dates
                    ...    ${client data list}
                    ...    ${configlist}
                    ...    ${CalculatedDate}
                    ${followupdateStr}=    Convert To String    ${FollowUpDateFiltered}
                    ${CalculatedDate}=    Is Weekend    ${followupdateStr}    ${client data list}[7]
                    ${CalculatedDateFiltered}=    Coverting Calculated Date
                    ...    ${client data list}
                    ...    ${configlist}
                    ...    ${CalculatedDate}
                    IF    '${FollowUpDateFiltered}' == '${TodayFilteredDate}'
                        Log    Dates are equal
                        Checking Communication Type and sending mail
                        ...    ${configlist}
                        ...    ${client data list}
                        ...    ${rowCount}
                        Updating followup date as Calculated date In excel
                        ...    ${configlist}
                        ...    ${rowCount}
                        ...    ${CalculatedDate}
                        Total Folowup Counter    ${client data list}    ${rowcount}
                        IF    '${TodayFilteredDate}' == '${CalculatedDateFiltered}'
                            ${rowCount}=    Evaluate    ${rowCount}+1
                        ELSE
                            Updating followup date as Calculated date In excel
                            ...    ${configlist}
                            ...    ${rowCount}
                            ...    ${CalculatedDate}
                            ${rowCount}=    Evaluate    ${rowCount}+1
                        END
                    ELSE IF    '${TodayFilteredDate}' < '${FollowUpDateFiltered}'
                        Log    ${TodayFilteredDate}
                        Log    ${FollowUpDateFiltered}
                        Log    FollowUpdate date is greater than current date
                        ${rowCount}=    Evaluate    ${rowCount}+1
                    ELSE
                        Log    FollowUpDate date is smaller than current date
                        IF    '${CalculatedDateFiltered}' == '${TodayFilteredDate}'
                            Updating followup date as Calculated date In excel
                            ...    ${configlist}
                            ...    ${rowCount}
                            ...    ${CalculatedDate}
                            Checking Communication Type and sending mail
                            ...    ${configlist}
                            ...    ${client data list}
                            ...    ${rowCount}
                            Total Folowup Counter    ${client data list}    ${rowCount}
                            ${rowCount}=    Evaluate    ${rowCount}+1
                        ELSE IF    '${CalculatedDateFiltered}' > '${TodayFilteredDate}'
                            Updating followup date as Calculated date In excel
                            ...    ${configlist}
                            ...    ${rowCount}
                            ...    ${CalculatedDate}
                            ${rowCount}=    Evaluate    ${rowCount}+1
                        ELSE
                            Log    Both Followup date and Calculated Followup has been Passed
                            ${rowCount}=    Evaluate    ${rowCount}+1
                        END
                    END
                ELSE
                    Log    Follow up required is No
                    Emptying Cell from excel    ${configlist}    ${rowCount}    ${CalculatedDate}
                    ${rowCount}=    Evaluate    ${rowCount}+1
                END
            ELSE
                Log    No sales Dependancy
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
        ${configlist}=    Create List
        Append To List
        ...    ${configlist}
        ...    ${InputFilePath}
        ...    ${EnquiryEmailBody}
        ...    ${FollowUpEmailBody}
        ...    ${TodayDate}
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
    ${followup_Required}=    Set Variable    ${row}[Follow Up Required]
    ${TotalFollowupMail}=    Set Variable    ${row}[Total no. of followup mail send]
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
    ...    ${followup_Required}
    ...    ${TotalFollowupMail}
    RETURN    ${client data list}

Converting Current,FollowUp and Calculated Dates
    [Arguments]    ${client data list}    ${configlist}    ${CalculatedDate}
    ${FollowUpDateFiltered}=    Convert Date    ${client data list}[6]    result_format=%d-%m-%Y
    Log    ${FollowUpDateFiltered}
    ${TodayFilteredDate}=    Convert Date    ${configlist}[3]    result_format=%d-%m-%Y
    RETURN    ${FollowUpDateFiltered}    ${TodayFilteredDate}

Coverting Calculated Date
    [Arguments]    ${client data list}    ${configlist}    ${CalculatedDate}
    ${CalculatedDateFiltered}=    Convert Date    ${CalculatedDate}    result_format=%d-%m-%Y
    RETURN    ${CalculatedDateFiltered}

checking Communication Type and sending mail
    [Arguments]    ${configlist}    ${client data list}    ${rowCount}
    IF    "${client data list}[5]" == "Enquiry"
        Enquiry Process    ${client data list}    ${configlist}
        Log    enquiry email should be sent
    ELSE IF    "${client data list}[5]" == "Follow Up"
        Follow up Process    ${client data list}    ${configlist}
        Log    follow up email should be sent
    ELSE
        Log    No communication needed
    END

Updating followup date as Calculated date In excel
    [Arguments]    ${configlist}    ${rowCount}    ${CalculatedDate}
    TRY
        Open Workbook    ${configlist}[0]    overwrite=True
    EXCEPT
        Log    Writing Excel Option
    ELSE
        Set Cell Value    ${rowCount}    I    ${CalculatedDate}
        Save Workbook
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
    [Arguments]    ${client data list}    ${configlist}
    ${EnquiryEmailBody}=    Replace String    ${configlist}[2]    <SalesExecName>    ${client data list}[2]
    ${EnquiryEmailBody}=    Replace String    ${EnquiryEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[2]
    Log    ${client data list}[3]

    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${EnquiryEmailBody}

Follow up Process
    [Arguments]    ${client data list}    ${configlist}
    ${FollowUpEmailBody}=    Replace String    ${configlist}[2]    <SalesExecName>    ${client data list}[2]
    ${FollowUpEmailBody}=    Replace String    ${FollowUpEmailBody}    <proposal>    ${client data list}[1]
    Log    ${configlist}[2]
    Log    ${client data list}[3]
    Send Email    recipients=${client data list}[3]
    ...    subject=${client data list}[1]
    ...    body=${FollowUpEmailBody}

Total Folowup Counter
    [Arguments]    ${client data list}    ${rowcount}
    IF    "${client data list}[9]" == "None"
        Open Workbook    Opportunity tracker_Remainder.xlsx
        Set Cell Value    ${rowcount}    K    1
        Save Workbook
        ${rowcount}=    Evaluate    ${rowcount} + 1
    ELSE
        ${TotalFollowupMail}=    Evaluate    ${client data list}[9] + 1
        Set Cell Value    ${rowcount}    K    ${TotalFollowupMail}
        Log    ${TotalFollowupMail}
        Save Workbook
        ${rowcount}=    Evaluate    ${rowcount} + 1
    END
