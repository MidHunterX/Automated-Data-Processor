import function as fn
from function import var
from pathlib import Path


def main():
    """
    Gets all files from every District dirs and generates total estimated amount
    """
    total_amt = 0

    district_dataset = var["district_dataset"]
    input_dir = var["input_dir"]

    file_list = []
    input_dir_path = Path(input_dir)

    for district in district_dataset:
        district_path = input_dir_path / district
        dist_file_list = fn.getFileList(str(district_path), [".docx"])

        if dist_file_list:
            file_list.append(dist_file_list)

    for dist_files in file_list:
        for file in dist_files:
            print(file)
            student_data = fn.getStudentDetails(file)

            for _, value in student_data.items():
                standard = value[1]
                standard = fn.convertStdToNum(standard)
                amt = fn.convertStdToAmount(standard)
                total_amt += amt

    print(f"Estimated Amount: {total_amt}")
    return 0


if __name__ == "__main__":
    main()
