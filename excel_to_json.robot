*** Settings ***
Documentation       Downloads JSON data from a remote API and writes it
...                 into a local Excel file.

Library        ExcelLibrary
Library        RequestsLibrary
Library        Collections
Library        OperatingSystem
Library        json

*** Keywords ***
Create an Excel file with data from a remote API
    ${document}=    Create Excel Document    comments.xlsx
    Should Be Equal As Strings    comments.xlsx    ${document}
    Write Excel Cell        1        1    Post ID
    Write Excel Cell        1        2    ID
    Write Excel Cell        1        3    Name
    Write Excel Cell        1        4    Email address
    Write Excel Cell        1        5    Body
    ${response}=        Get    https://jsonplaceholder.typicode.com/comments
    ${json_data}=        Set Variable    ${response.json()}
    ${JSON_CONTENT}=      json.dumps  ${response.json()}
    TRY
            Log to console    Data returned successfully
    EXCEPT
            Log to console    Cannot retrieve JSON due to invalid data    console=True
    END
    ${data_for_excel}=    Create List
    ${row_number}=    Set Variable    2   # Start writing from row 2 (header is on row 1)
    FOR    ${item}    IN    @{json_data}
        ${post_id}=    Convert To String    ${item['postId']}
        ${id}=    Convert To String    ${item['id']}
        ${name}=    Set Variable    ${item['name']}
        ${email}=    Set Variable    ${item['email']}
        ${body}=    Set Variable    ${item['body']}
        Write Excel Cell    ${row_number}    1    ${post_id}
        Write Excel Cell    ${row_number}    2    ${id}
        Write Excel Cell    ${row_number}    3    ${name}
        Write Excel Cell    ${row_number}   4    ${email}
        Write Excel Cell    ${row_number}    5    ${body}
        ${row_number}=    Evaluate    ${row_number} + 1
    END
    Save Excel Document    ${document}
    Close Current Excel Document

Convert Excel file to JSON And Make a POST request
    ${document}=    Open Excel Document    comments.xlsx    1
    ${json_data}=    Create List
    FOR    ${row_number}    IN RANGE    2    501  # Set a high range, assuming your data doesn't exceed this number of rows
        ${row_data}=    Read Excel Row    ${row_number}    sheet_name=Sheet
        ${is_empty}=    Run Keyword And Return Status    List Should Be Empty    ${row_data}
        Run Keyword If    ${is_empty}    Exit For Loop
        ${item}=    Create Dictionary    postId=${row_data[0]}    id=${row_data[1]}    name=${row_data[2]}    email=${row_data[3]}    body=${row_data[4]}
        Append To List    ${json_data}    ${item}
    END
    ${json_string}=    Evaluate    json.dumps(${json_data})
    Log To Console    ${json_string}

    # Make a POST request with the JSON data
    ${headers}=    Create Dictionary    Content-Type=application/json
    ${response}=    Post    https://example.com/api/endpoint    data=${json_string}    headers=${headers}

    # Verify the response
    Should Be Equal As Strings    ${response.status_code}    200
    Log To Console    Response Body: ${response.text}


*** Test Cases ***
Convert Excel File to JSON
    Create an Excel file with data from a remote API
    Convert Excel file to JSON And Make a POST request
