import shutil       # Copying and Moving files
import function as fn
from function import var


def main():

    input_dir = var["input_dir"]
    investigation_dir = fn.initNestedDir(input_dir, "for checking")
    formatting_dir = fn.initNestedDir(input_dir, "formatting issues")
    district_user = fn.getDistrictFromUser()
    files_written = 0
    for_checking_count = 0
    incorrect_format_count = 0
    file_list = fn.getFileList(input_dir, [".docx", ".pdf"])

    try:
        for file in file_list:

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

    print("")
    horizontal_line = "-"*80
    print(horizontal_line)
    print("FINAL REPORT".center(80))
    print(horizontal_line)
    print(f"Files Accepted    : {files_written}".center(80))
    print(f"For Checking      : {for_checking_count}".center(80))
    print(f"Formatting Issues : {incorrect_format_count}".center(80))
    print(horizontal_line)
