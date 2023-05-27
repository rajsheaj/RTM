import os.path
import threading
from _cffi_backend import callback
from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput

from ExtractTestName import ExtractTestName
from ExtractTestProtocols import ExtractTestProtocols
from SRSGathering import SRSGathering
from WriteToExcel import WriteToExcel

class GridLayout(GridLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.rows = 2
        self.cols = 2
        self.text1 = TextInput(text='Enter SRS Path', multiline=False)
        self.add_widget(self.text1)

        self.text2= TextInput(text='Enter TP Path', multiline=False)
        self.add_widget(self.text2)

        self.btn1 = Button(text = 'Calculate RTM')
        self.btn1.bind(on_press = self.callback)
        self.add_widget(self.btn1)

        self.btn2 = Button(text='Clear text')
        self.btn2.bind(on_press= self.callback2)
        self.add_widget(self.btn2)

        self.popup = Popup(title='RTM',
                      content=Label(text='invalid path'),
                      size_hint=(None, None), size=(200, 200))

    def callback(self, elem):

        if os.path.isfile(self.text1.text) and os.path.exists(self.text2.text):
            threading.Thread(target=self.do_callback).start()
        else:
            self.popup.open()

    def do_callback(self):
        print(self.text1.text)
        print(self.text2.text)
        # logic to call the RTM, SRS first
        srs_object = SRSGathering(self.text1.text)
        srs = srs_object.gather_srs()
        excel_obj = WriteToExcel('SRS.xlsx', srs)
        excel_obj.writetoexcel_DF()
        # extract test names
        tp_object = ExtractTestName(self.text2.text)
        a = tp_object.extract_test_name()
        excel_obj = WriteToExcel('extracted test names.xlsx', a)
        excel_obj.writetoexcel_DF()
        # extract test steps
        ts_object = ExtractTestProtocols(self.text2.text)
        a = ts_object.extract_test_steps()
        excel_obj = WriteToExcel('extracted test steps.xlsx', a)
        excel_obj.writetoexcel_DF()
        # Generate RTM
        rtm_df = srs.merge(a, on=['Req_ID'], how='left')
        rtm_df.drop('Req_Text', axis=1, inplace=True)
        excel_obj = WriteToExcel('rtm.xlsx', rtm_df)
        excel_obj.writetoexcel_DF()

        # find orphans in TP and not in SRS
        print("completed")

    def callback2(self, elem):
        self.text1.text = ""
        self.text2.text = ""

class DemoApp(App):
    def build(self):
        return GridLayout()
