"""This script changes .docx files author's name in directory where py-file placed"""

import os
import docx


author = input("Input author's name: ")
directory = os.path.dirname(os.path.abspath(__file__))

for filename in os.listdir(directory):
    try:
        name, file_type = filename.split('.')
        print(name)
    except:
        continue
    
    else:
        dir_and_file = os.path.join(directory, filename)
        
        if os.path.isfile(dir_and_file) and file_type == 'docx':
            document = docx.Document(filename)
            core_properties = document.core_properties
            core_properties.author = author
            document.save(f"{name}(copy).docx")


print("Done!")
