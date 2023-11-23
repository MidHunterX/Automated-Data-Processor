import csv          # CSV file manipulation


def initVarCommon():
    """
Returns:
    var = {
        "input_dir": "input",
        "db_file": "data\\database.db",
        "ifsc_dataset": loadIfscDataset("data\\IFSC.csv"),
        "district_dataset": loadDistrictDataset(),
        "excel_file": "output.xlsx",
    }
    """
    var = {
        "input_dir": "input",
        "db_file": "data\\database.db",
        "ifsc_dataset": loadIfscDataset("data\\IFSC.csv"),
        "district_dataset": loadDistrictDataset(),
        "excel_file": "output.xlsx",
    }
    return var


def initVarCmd():
    """
Returns:
    cmd = {
        "db": "database",
        "form": "forms",
        "ifsc": "ifsc",
        "excel": "spreadsheet",
        "bank": "neft",
    }
    """
    cmd = {
        "db": "database",
        "form": "forms",
        "ifsc": "ifsc",
        "excel": "spreadsheet",
        "bank": "neft",
    }
    return cmd


def loadDistrictDataset():
    district_list = [
        "Thiruvananthapuram", "Kollam", "Pathanamthitta", "Alappuzha",
        "Kottayam", "Idukki", "Ernakulam", "Thrissur", "Palakkad",
        "Malappuram", "Kozhikode", "Wayanad", "Kannur", "Kasargod"
    ]
    return district_list


def loadIfscDataset(csv_file):
    """
    Parameter: CSV Dataset from RazorPay
    Returns: Dataset Dictionary loaded into memory

    dataset[row['IFSC']] = {
        'Bank': row['BANK'],
        'Branch': row['BRANCH'],
        'Centre': row['CENTRE'],
        'District': row['DISTRICT'],
        'State': row['STATE'],
        'Address': row['ADDRESS'],
        'City': row['CITY'],
    }
    """
    dataset = {}
    with open(csv_file, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            dataset[row['IFSC']] = {
                'Bank': row['BANK'],
                'Branch': row['BRANCH'],
                'Centre': row['CENTRE'],
                'District': row['DISTRICT'],
                'State': row['STATE'],
                'Address': row['ADDRESS'],
                'City': row['CITY'],
            }
    return dataset
