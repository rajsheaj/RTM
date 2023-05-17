import pandas as pd


class WriteToExcel:
    def __init__(self, file, dict):
        self.file = file
        self.dict = dict

    def writetoexcel_Dict(self):
        df = pd.DataFrame(data=self.dict, index=[0])
        df = df.T
        df.to_excel(self.file)

    def writetoexcel_DF(self):
        self.dict.to_excel(self.file)

