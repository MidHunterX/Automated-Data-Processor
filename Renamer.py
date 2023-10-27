import os
import shutil

input_dir = "test"


def preprocessFiles(input_dir):
    """
    Renames every Supported files into numbers
    Moves Unsupported files into a separate directory
    """
    unsupported_dir = os.path.join(input_dir, "unsupported")
    counter = 1

    # Ensure the unsupported directory exists
    if not os.path.exists(unsupported_dir):
        os.makedirs(unsupported_dir)

    for filename in os.listdir(input_dir):
        file_path = os.path.join(input_dir, filename)

        if os.path.isfile(file_path):
            # Check if it's a PDF or DOCX file
            if filename.lower().endswith(('.pdf', '.docx')):
                base_extension = os.path.splitext(filename)[1]
                new_name = f"{counter:03d}{base_extension}"
                new_path = os.path.join(input_dir, new_name)
                os.rename(file_path, new_path)
                counter += 1
            else:
                # Move unsupported files to the 'unsupported' directory
                unsupported_path = os.path.join(unsupported_dir, filename)
                shutil.move(file_path, unsupported_path)

    print("Files renamed and unsupported files moved.")


preprocessFiles(input_dir)
