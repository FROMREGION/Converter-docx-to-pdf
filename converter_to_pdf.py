import os
import comtypes.client
import re

class Converter:
    def __init__(self):
        self.wdFormatPDF = 17
        self.CURRECT_DIR = os.getcwd()

    def remove_temp_files(self):
        filter_docx = lambda in_filter: '.docx' in in_filter
        files = list(filter(filter_docx, os.listdir(self.CURRECT_DIR)))
        for file in range(len(files)):
            if re.search(r'~.\w+.docx', files[file]) is not None:
                files.remove(re.search(r'~.\w+.docx', files[file]).group())
        print(f'All doc in dir: {files}')
        print(f'Count: {len(files)}')
        return files

    def save_as_pdf(self):
        files = self.remove_temp_files()
        for file in range(len(files)):
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(f"{self.CURRECT_DIR}\\{files[file]}")
            doc.SaveAs(f"{self.CURRECT_DIR}\\{files[file][:-5]}.pdf", FileFormat=self.wdFormatPDF)
            doc.Close()
            word.Quit()
            print(f"[{file+1}] File {files[file]} converted to PDF")
        print()
        print("All file converted!")
        print('===================')


#
# ______________________________________________________________________________
if __name__ == '__main__':
    convert = Converter().save_as_pdf()
