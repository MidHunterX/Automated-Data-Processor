## Usage

This is a CLI tool so the features are accessed using the following commands:

### Commands

| Command     | Description                                           |
| ----------- | ----------------------------------------------------- |
| forms       | Parse, Clean, Validate and Organize Forms             |
| database    | Commits organized forms into a Database               |
| ifsc        | Converts pasted IFSC code into Branch name            |
| spreadsheet | Converts database into custom styled xlsx spreadsheet |
| neft        | Converts database into spreadsheet for NEFT transfers |

### Syntax

To use the above commands, you use it in your CLI like this:

```
process <command>
```

## How it works

### Process Forms

This Python script is designed for efficient processing and verification of student data collected from various institutions. The code begins by extracting details about the institution and student from files situated in "input" directory. The subsequent data processing steps involve cleaning and normalizing student information for accurate analysis.

One notable feature is the district guessing mechanism based on the IFSC codes associated with student data. The script intelligently suggests a potential district, allowing users to verify or override the district manually. The processed data is then printed for review.

The verification section prompts the user to confirm the accuracy of the student data. Upon successful verification, the script marks the data as correct and organizes it into a directory corresponding to the selected or guessed district. In cases of incorrect formatting, the script identifies errors, allowing for manual investigation and reformatting.

This script streamlines the workflow of handling student information, ensuring data accuracy and facilitating organized storage based on district categorization.
