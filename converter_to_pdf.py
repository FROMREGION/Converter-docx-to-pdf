import os
import comtypes.client
import re

# Constants
# ______________________________________________________________________________
filter_docx = lambda in_filter: '.docx' in in_filter
wdFormatPDF = 17
CURRECT_DIR = os.getcwd()
FILES = list(filter(filter_docx, os.listdir(CURRECT_DIR)))


# Remove word temp files in files list | example: ~$example.docx
# ______________________________________________________________________________
def remove_temp_files(lst):
    files = lst
    for file in range(len(files)):
        if re.search(r'~.\w+.docx', files[file]) is not None:
            # print(re.search(r'~.\w+.docx', files[file]).group())
            files.remove(re.search(r'~.\w+.docx', files[file]).group())
        # else:
            # print(re.search(r'~.\w+.docx', files[file]))
    print(f'All doc in dir: {files}')
    print(f'Count: {len(files)}')
    return files


# Save .docx to .pdf
# ______________________________________________________________________________
def save_as_pdf(files):
    files = files
    for file in range(len(files)):
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(f"{CURRECT_DIR}\\{files[file]}")
        doc.SaveAs(f"{CURRECT_DIR}\\{files[file][:-5]}.pdf", FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        print(f"[{file+1}] File {files[file]} converted to PDF")
    print()
    print("All file converted!")
    print('===================')


#
# ______________________________________________________________________________
if __name__ == '__main__':
    save_as_pdf(remove_temp_files(FILES))
