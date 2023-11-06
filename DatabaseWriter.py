import os           # Directory path support
import shutil       # Copying and Moving files
import sqlite3      # SQLite DB
import datetime     # ISO Date
import Processor as pr

input_dir = "input"
db_file = "data\\database.db"

iso_date = datetime.date.today().isoformat()
ifsc_dataset = pr.loadIfscDataset("data\\IFSC.csv")

district_user = pr.getDistrictFromUser()
print(f"Selected District: {district_user}")

input_dir = pr.initNestedDir(input_dir, district_user)
output_dir = pr.initNestedDir(input_dir, iso_date)
investigation_dir = pr.initNestedDir(input_dir, "for checking")
rejected_dir = pr.initNestedDir(input_dir, "rejected")

# ----------------------------------------------------- [ VARS FOR REPORT ]

files_written = 0
for_checking_count = 0
incorrect_format_count = 0

# Open connection to Database
print("Connecting to Database")
conn = sqlite3.connect(db_file)

# Rechecking Files
file_list = pr.getFileList(input_dir, [".docx", ".pdf"])

for file in file_list:
    file_name, file_extension = os.path.basename(file).split(".")

    # Flags
    proceed = False
    valid_std = False

    # Check for correct formatting in form
    if pr.correctFormat(file):
        proceed = True
        print(f"\n==== {file_name}.{file_extension} ====")
        institution = pr.getInstitutionDetails(file)
        student_data = pr.getStudentDetails(file)

    if proceed is True:

        # Guessing District
        ifsc_list = pr.getStudentIfscList(student_data)
        dist_guess = pr.guessDistrictFromIfscList(ifsc_list, ifsc_dataset)
        print(f"Possible District: {dist_guess}")

        # Deciding User District vs Guessed District
        district = district_user
        if district == "Unknown":
            district = dist_guess
        print(f"Selected District: {district}\n")

        # Normalizing Student Data
        student_data = pr.normalizeStudentStd(student_data)
        student_data = pr.normalizeStudentBranch(student_data, ifsc_dataset)

        if pr.isValidStudentStd(student_data):
            valid_std = True

        # Printing Final Data
        pr.printInstitution(institution)
        print("")
        pr.printStudentDataFrame(student_data)

        if valid_std is True:
            verification = input("\nCorrect? (ret / n): ")
            print("")
        else:
            verification = "n"

        # Write to database
        if verification == "":
            print("Marking as Correct.")
            if pr.writeToDB(conn, district, institution, student_data):
                print("Data Written Successfully!")
                # writeToCSV(csv_file, institution, student_data)
                shutil.move(file, output_dir)
                files_written += 1
            else:
                print("Rejected by Database")
                shutil.move(file, rejected_dir)
                for_checking_count += 1
        else:
            print("Moving for further Investigation.")
            shutil.move(file, investigation_dir)
            for_checking_count += 1

# Close Connection to Database
print("Closing DB")
conn.close()
