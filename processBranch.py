import sys          # Command line arguments and exit
import pyperclip    # Clipboard handling


def main(var):
    ifsc_dataset = var["ifsc_dataset"]
    getBranchFromPastedIfsc(ifsc_dataset)


def getBranchFromPastedIfsc(ifsc_dataset):

    print("")
    print("📝 Paste IFSC and press Ctrl+Z")
    print("-------------------------------")
    text = read_pasted_text()
    branch_list = []
    for ifsc in text:
        ifsc = ifsc.strip()
        branch = getBranchFromIfsc(ifsc, ifsc_dataset)
        branch_list.append(branch)

    # Join the list into a single string
    text = '\n'.join(branch_list)

    # Copy to Clipboard
    print("")
    print("✅ Copied to Clipboard")
    print("-----------------------")
    print(text)
    pyperclip.copy(text)


def getBranchFromIfsc(ifsc, ifsc_dataset):
    """
    Arguments: (ifsc, ifsc_dataset)
        - ifsc_list: IFSC Code
        - ifsc_dataset: IFSC Razorpay Dataset from loadIfscDataset()

    Returns:
        - "": If there exists no record of IFSC in dataset
        - branch: If Branch for IFSC is found
    """
    branch = ""
    if type(ifsc) is str:
        ifsc_details = ifsc_dataset.get(ifsc)
        if ifsc_details:
            branch = ifsc_details["Branch"]

    # Clean Up Branch
    if "IMPS" in branch:
        branch = branch.replace("IMPS", "").strip()

    return branch


def read_pasted_text():
    lines = []
    for line in sys.stdin:
        lines.append(line.rstrip('\n'))
    # Join the lines into a single string
    # pasted_text = '\n'.join(lines)
    return lines
