# Extracting Data from Remote API to Excel and Making a POST Request

## Overview

This project demonstrates how to use Robot Framework along with the RequestsLibrary and ExcelLibrary to interact with a remote API, extract JSON data, store it in an Excel file, convert it back to JSON, and then make a POST request with the extracted data. The project aims to showcase the understanding of API interactions, data handling, and automation using Robot Framework.

## Project Structure

project_root/
|-- tests/
| |-- excel_to_json.robot
| |-- json_to_excel.robot
| |-- comments.xlsx
|-- README.md


## Requirements

- Python (3.6+)
- Robot Framework
- RequestsLibrary
- ExcelLibrary

## Usage

1. Install the required libraries:

    ```bash
pip install robotframework requests robotframework-requests ExcelLibrary
    ```

2. Place the `json_to_excel.robot` file in the `tests` directory.

3. Update the `json_to_excel.robot` file to replace `https://example.com/api/endpoint` with the actual API endpoint you want to interact with.

4. Execute the test using the following command:

    ```bash
    robot tests/json_to_excel.robot
    ```


## Robot Test Cases

### Create an Excel file with data from a remote API

This test case performs the following steps:

1. Downloads JSON data from the [JSON Placeholder](https://jsonplaceholder.typicode.com/comments) API, which contains information about comments.
2. Writes the extracted data to an Excel file named `comments.xlsx`.
3. The data includes `postId`, `id`, `name`, `email`, and `body` of comments.

### Convert Excel file to JSON And Make a POST request

This test case performs the following steps:

1. Reads data from the `comments.xlsx` Excel file.
2. Converts the data back to JSON format.
3. Makes a POST request to the specified API endpoint (`https://example.com/api/endpoint`) with the JSON data.
4. Verifies that the response status code is `200`.

## Key Concepts

1. **RequestsLibrary**: This library provides keywords to interact with HTTP-based APIs. It allows us to send GET and POST requests and handle the responses.

2. **ExcelLibrary**: This library allows us to read and write data from and to Excel files. It helps in data storage and retrieval.

3. **API Interaction**: We use the `RequestsLibrary` to interact with the JSON Placeholder API, download JSON data, and make a POST request to a custom API endpoint.

4. **Data Handling**: We extract data from the API response, convert it into a suitable format, and write it to an Excel file. Later, we read the Excel data and convert it back to JSON for the POST request.

5. **Automation**: The entire process is automated using Robot Framework, which allows us to perform these operations seamlessly and reliably.

## Conclusion

This project demonstrates a practical application of the Robot Framework in API interactions and data handling. Expand the project to demonstrate additional capabilities and impress your audience further.
