from docx.api import Document
import pandas as pd


class SRSGathering:
    def __init__(self, path):
        self.file_path = path
        self.doc = Document(self.file_path)
        self.req_dict = {}
        self.duplicate_list = []

    def gather_srs(self):
        for table in self.doc.tables:
            if len(table.columns) == 3:
                for row in table.rows:
                    req_id = row.cells[0].text
                    req_text = row.cells[1].text
                    if "SRS_" in req_id:
                        if req_id in self.req_dict.keys():
                            print("requirement duplicate conflict observed - ", req_id)
                            self.duplicate_list.append(req_id)
                        else:
                            self.req_dict[req_id] = req_text

        df = pd.DataFrame(self.req_dict.items(), columns=['Req_ID', 'Req_Text'])
        return df
        #return [self.req_dict, self.duplicate_list]
