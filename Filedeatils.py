import glob
import os
import time

def get_file_details_glob(pattern):
    file_details = []
    for filepath in glob.glob(pattern):
        if os.path.isfile(filepath):
            file_info = {
                'name': os.path.basename(filepath),
                'size': os.path.getsize(filepath),
                'created': time.ctime(os.path.getctime(filepath)),
                'modified': time.ctime(os.path.getmtime(filepath))
            }
            file_details.append(file_info)
    return file_details


# Example usage
pattern = 'C:/Users/PioneerGuest/Desktop/Group tool'
details = get_file_details_glob(pattern)
for detail in details:
    print(detail)

