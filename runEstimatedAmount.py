import function as fn
from function import var


def main():
    total_amt = 0

    district_dataset = var["district_dataset"]
    input_dir = var["input_dir"]

    file_list = []
    for district in district_dataset:
        dist_list = fn.getFileList(f"{input_dir}\\{district}", [".docx"])
        if dist_list:
            file_list.append(dist_list)

    for dist_files in file_list:
        for file in dist_files:
            print(file)
            student_data = fn.getStudentDetails(file)

            for key, value in student_data.items():
                standard = value[1]
                standard = fn.convertStdToNum(standard)
                amt = fn.convertStdToAmount(standard)
                total_amt += amt

    print(f"Final Amount: {total_amt}")
    return 0


if __name__ == "__main__":
    main()
