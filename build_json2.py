
#!utils/python/3.6.0
#!python3
import sys
import os
import csv
import fileinput
import xlrd
import wx
import collections
print(sys.version)
#print(sys.base_exec_prefix + '\\python.exe')

workbook = xlrd.open_workbook('S:/building_json_for_nemsdv/layin.v1.233.xls')
sheet = workbook.sheet_by_name('layin')
table_name = {}
for row in range(0,160):
    if (sheet.cell(row,2).value in range(0,155)):
    #if (0 < int(sheet.cell(row,2).value) < 155):
        table_name[sheet.cell(row,2).value] = sheet.cell(row,1).value

table_name = {int(k):v.encode('ascii') for k,v in table_name.items() if not '=' in v}
#table_name = {int(k):v.encode('ascii') for k,v in table_name.items()}
#print table_name

with open("S:/building_json_for_nemsdv/ref2019.0906a.api.txt", "r") as f:
    searchlines = f.readlines()

row_name = {}
stub_name = {}
for i, line in enumerate(searchlines):
    if i > 0:
        #print line.split('|')[2]
        row_name[(line.split('|')[2],line.split('|')[3])] = line.split('|')[6]
        stub_name[(line.split('|')[2],line.split('|')[3])] = line.split('|')[5]


row_name = {(l.replace('"', ''),m.replace('"', '')):v.replace('"', '') for (l,m),v in row_name.items()}
row_name = {(int(l),int(m)):v for (l,m),v in row_name.items()}

stub_name = {(l.replace('"', ''),m.replace('"', '')):v.replace('"', '') for (l,m),v in stub_name.items()}
stub_name = {(int(l),int(m)):v for (l,m),v in stub_name.items()}

# q = {}
# rowname_formatted = []
# q = {(l,m): v for (l,m),v in row_name.items() if l == 1}
# for (l,m), v in q.items(): rowname_formatted.append(str(m)+" "+str(v))
# print(rowname_formatted)

#print (stub_name)
tablename_formatted = []
od = collections.OrderedDict(sorted(table_name.items()))
for k, v in od.items(): tablename_formatted.append(str(k)+" "+str(v.decode('utf-8')))


class MyFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        wx.Frame.__init__(self, *args, **kwargs, size=(800, 400))

        self.p1 = wx.Panel(self)
        self.lst1 = tablename_formatted
        self.lst2 = self.get_lst2()

        self.st = wx.ComboBox(self.p1, -1, choices = self.lst1, style=wx.TE_PROCESS_ENTER)
        self.st2 = wx.ComboBox(self.p1, -1, choices = self.lst2, style=wx.TE_PROCESS_ENTER, pos=(-1,100), size=(300,-1))

        self.st.Bind(wx.EVT_COMBOBOX, self.update)

        self.st2.Bind(wx.EVT_COMBOBOX, self.func)



    def get_lst2(self, selectn=None):

        q = {}
        rowname_formatted = []
        q = {(l,m): v for (l,m),v in row_name.items() if l == selectn}
        for (l,m), v in q.items(): rowname_formatted.append(str(m)+" "+str(v))
        return rowname_formatted

    def update(self, event):
        selectn = int(self.st.GetValue().split(' ')[0])
        #print(type(selectn))
        self.lst2 = self.get_lst2(selectn)
        self.st2.Clear()
        for number in self.lst2:
           self.st2.Append(number)
        # if selectn == '1':
        #     self.st.Hide()
    
    # def writ(self, event): 
    #     self.label.SetLabel("You selected "+self.st.GetValue().split(' ')[0]+" from Combobox")

    def OnCombo(self, event):
        table_num = int(self.st.GetValue().split(' ')[0])
        row_num = int(self.st2.GetValue().split(' ')[0])
        return stub_name[(table_num,row_num)]

    def func(self,event):
        x=self.OnCombo(None)
        print (x)
        #wx.MessageBox(x)
        #self.out.WriteText(x)


class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, '20_combobox.py')
        frame.Show()
        self.SetTopWindow(frame)
        return 1


if __name__ == "__main__":
    app = MyApp(1)
#    wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()


