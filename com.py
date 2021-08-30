#%%
class CpStockCode:
    def __init__(self):
        self.stocks = {'유한양행' : 'A000100'}
        
    def GetCount(self):
        return len(self.stocks)
    
    def NameToCode(self, name):
        return self.stocks[name]
# %%
instCpStockCode = CpStockCode()
print(instCpStockCode.GetCount())
print(instCpStockCode.NameToCode('유한양행'))
# %%
import win32com.client

explore = win32com.client.Dispatch("InternetExplorer.Application")
explore.Visible = True
# %%
import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = True
# %%
import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")
ws.Cells(1, 1).Value = "hello world"
wb.SaveAs("e:\\Github\\algorithm_trading\\test.xlsx")
excel.Quit()
# %%
import win32com.client

excle = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open("e:\\Github\\algorithm_trading\\input.xlsx")
ws = wb.ActiveSheet
print(ws.Cells(1,1).Value)
excel.Quit()
# %%
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('e:\\Github\\algorithm_trading\\input.xlsx')
ws = wb.ActiveSheet

ws.cells(1,2).Value = "is"
ws.Range("C1").Value = "good"
ws.Range("C1").Interior.ColorIndex = 10
ws.Range("A2:C2").Interior.ColorIndex = 27

excel.Quit()
# %%


# %%
