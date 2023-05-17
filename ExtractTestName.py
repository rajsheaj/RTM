from docx.api import Document
import re
import os
from os import listdir
from os.path import isfile, join
import pandas as pd




class ExtractTestName:

    def __init__(self, path):
        self.files_path = [os.path.join(path, f) for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        print(self.files_path)
        self.df = pd.DataFrame(columns = ['Protocol_Name','Test_Name', 'Test_ID'])

    def extract_test_name(self):
        for file in self.files_path:
            head, tail = os.path.split(file)
            doc = Document(file)

            for paragraph in doc.paragraphs:
                if paragraph.style.name == 'Heading 2':
                    if paragraph.text.find("Test Case ID / Title:") != -1:
                        ind = paragraph.text.find("Test Case ID / Title:")
                        testname = paragraph.text[(21 + ind):]
                        testid = paragraph.text[:ind].strip()
                        newRow = {"Protocol_Name": tail, "Test_Name": testname, "Test_ID": testid}
                        row_df = pd.DataFrame([newRow])
                        self.df = pd.concat([self.df, row_df], ignore_index=True)

        return self.df
