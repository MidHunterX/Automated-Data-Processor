# Scholar CAP

Your All-in-One Solution for Student Scholarship Processing

Scholar CAP (Computer Aided Processing) is a comprehensive toolset designed to simplify the processing of student scholarship forms, specifically focusing on banking details. From initial data extraction and cleaning to verification, correction of typos, and the generation of NEFT formats. Scholar Toolsets included in Scholar CAP ensures a seamless workflow for efficiently handling student information. This versatile project empowers users to enhance accuracy and organization throughout the scholarship processing journey.

## Usage

| Command             | Description                                           |
| ------------------- | ----------------------------------------------------- |
| process forms       | Parse, Clean, Validate and Organize Forms             |
| process database    | Commits organized forms into a Database               |
| process ifsc        | Converts pasted IFSC code into Branch name            |
| process spreadsheet | Converts database into custom styled xlsx spreadsheet |
| process neft        | Converts database into spreadsheet for NEFT transfers |

## How it works

### Process Forms

**Data Processing and Verification**

It is designed for efficient processing and verification of student data collected from various institutions. The code begins by extracting details about the institution and student from files situated in "input" directory. The subsequent data processing steps involve cleaning and normalizing student information for accurate analysis.

One notable feature is the district guessing mechanism based on the IFSC codes associated with student data. The script intelligently suggests a potential district, allowing users to verify or override the district manually. The processed data is then printed for review.

The verification section prompts the user to confirm the accuracy of the student data. Upon successful verification, the script marks the data as correct and organizes it into a directory corresponding to the selected or guessed district. In cases of incorrect formatting, the script identifies errors, allowing for manual investigation and reformatting.

This script streamlines the workflow of handling student information, ensuring data accuracy and facilitating organized storage based on district categorization.

### Process Database

**Data Verification and Database Integration**

It is designed to be used after organizing forms by executing `process forms`. The primary focus is on verifying the correctness of student data and integrating it into a database. The script begins by checking if the file adheres to the correct format. If the format is correct, the user is prompted to specify the district for the organized data, with the option to automatically determine it based on the Indian state if unspecified.

The script then prints details about the institution and the processed student data for review. The verification section allows the user to confirm the accuracy of the data. Upon successful verification, the script marks the data as correct and proceeds to write it into the database. The user is notified of the success or rejection by the database, and the corresponding actions are taken. Verified data is moved to an output directory, while rejected data is moved to a separate directory for further investigation.

In cases of incorrect formatting, the script identifies errors, prompting the user for manual investigation and reformatting. This script provides a comprehensive solution for verifying and integrating student data into a database, ensuring data accuracy and facilitating efficient data management.

### Process IFSC

**IFSC to Branch Converter**

This tool simplifies the process of converting a list of IFSC codes to their corresponding branch names. Users can conveniently paste IFSC codes, and the script retrieves the associated branch names. The results are then copied to the clipboard for easy use. This tool streamlines the conversion task, providing a quick and efficient solution for handling IFSC data.

### Process Spreadsheet

**Excel Spreadsheet Generator**

This tool generates an Excel spreadsheet summarizing school and student information stored in database. The user specifies district name. The script dynamically fetches data based on the district, organizing it into a clear and structured spreadsheet. For each school, the spreadsheet includes institution details such as contact information and a list of students with their names, classes, account numbers, branches, and amounts.

### Process NEFT

**NEFT Format Generator for School Data**

This tool generates a spreadsheet in the NEFT (National Electronic Funds Transfer) format, summarizing student's banking information stored in database. Users provide the district name. The script dynamically fetches data based on the district, organizing it into a structured spreadsheet compatible with NEFT standards. The generated spreadsheet includes essential details such as account numbers, account types, account titles, addresses, IFSC codes, and transaction amounts for each student.
