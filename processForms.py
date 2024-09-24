import sqlite3  # SQLite DB operations
import shutil  # Copying and Moving files
import function as fn
from function import var


def main():

    db_file = var["db_file"]
    input_dir = var["input_dir"]
    investigation_dir = fn.initNestedDir(input_dir, "for checking")
    formatting_dir = fn.initNestedDir(input_dir, "formatting issues")
    rejected_dir = fn.initNestedDir(input_dir, "rejected")
    district_user = fn.getDistrictFromUser()
    files_written = 0
    for_checking_count = 0
    incorrect_format_count = 0
    rejected_count = 0
    file_list = fn.getFileList(input_dir, [".docx", ".pdf"])

    print("ℹ️ Connecting to Database")
    cursor = sqlite3.connect(db_file).cursor()

    try:
        for file in file_list:

            print(file)

            if fn.correctFormat(file):

                # -------------------------------------------- [ FORM PARSING ]

                institution = fn.getInstitutionDetails(file)
                student_data = fn.getStudentDetails(file)

                # --------------------------------------- [ FILENAME RENAMING ]

                file = fn.renameFilenameToInstitution(file, institution)
                fn.printFileNameHeader(file)

                # ----------------------------------------- [ DATA PROCESSING ]

                # Cleaning up Student Data for processing
                student_data = fn.cleanStudentData(student_data)
                # Normalizing Student Data
                student_data = fn.normalizeStudentData(student_data)

                # Guessing District
                ifsc_list = fn.getStudentIfscList(student_data)
                district_guess = fn.guessDistrictFromIfscList(ifsc_list)
                print(f"💡 Possible District: {district_guess}")

                # Check for duplicate accounts in database
                if fn.checkExistingAccounts(student_data, cursor):
                    print("❌ Duplicate account detected in Database!")
                    duplicate_accounts = fn.getExistingAccounts(student_data, cursor)
                    fn.printExistingAccounts(duplicate_accounts)
                    input("Move to Rejected? (ret) ")
                    shutil.move(file, rejected_dir)
                    rejected_count += 1
                    continue  # Skip to next iteration
                else:
                    pass

                # Deciding User District vs Guessed District
                district = district_user
                if district == "Unknown":
                    district = district_guess
                print(f"✍️ Selected District: {district}\n")

                # ------------------------------------------- [ DATA PRINTING ]

                fn.printInstitution(institution)
                print("")
                fn.printStudentDataFrame(student_data)
                print("")

                # ------------------------------------ [ VERIFICATION SECTION ]

                verification = fn.userVerifyStudentData(student_data)

                # --------------------------------------- [ ACTUATION SECTION ]

                if verification is True:
                    print("✅ Marking as Correct.")
                    # SORTING VERIFIED FORM INTO DISTRICT DIRECTORY
                    output_dir = fn.initNestedDir(input_dir, district)
                    shutil.move(file, output_dir)
                    files_written += 1
                else:
                    input("Move for Investigation? (ret) ")
                    print("❌ Moving for further Investigation.")
                    shutil.move(file, investigation_dir)
                    for_checking_count += 1

            # ---------------------------------------- [ INCORRECT FORMATTING ]

            else:
                fn.printFileNameHeader(file)
                print("⚠️ Formatting error detected!")
                input("Move for checking Format? (ret) ")
                print("❌ Moving for Re-Formatting.")
                shutil.move(file, formatting_dir)
                incorrect_format_count += 1

    except KeyboardInterrupt:
        print("Caught the Keyboard Interrupt ;D")

    # -------------------------------------------------------------- [ REPORT ]

    finally:
        print("ℹ️ Closing DB")
        cursor.close()
        print("")
        horizontal_line = "-" * 80
        print(horizontal_line)
        print("FINAL REPORT".center(80))
        print(horizontal_line)
        print(f"Files Accepted    : {files_written}".center(80))
        print(f"For Checking      : {for_checking_count}".center(80))
        print(f"Formatting Issues : {incorrect_format_count}".center(80))
        print(f"Rejected by DB    : {rejected_count}".center(80))
        print(horizontal_line)
