
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

        #self.panel = wx.Panel(self)     
        #self.quote = wx.StaticText(self.p1, label="Your quote:")
        #self.result = wx.StaticText(self.p1, label="")
        #self.result.SetForegroundColour(wx.RED)
        #self.button = wx.Button(self.p1, label="Save")
        self.titlename = wx.StaticText(self.p1,label="Title:")
        self.editname = wx.TextCtrl(self.p1, size=(200, -1))

        self.typename = wx.StaticText(self.p1,label="Type:")
        self.edittype = wx.TextCtrl(self.p1, size=(140, -1))
        # Set sizer for the frame, so we can change frame size to match widgets
        self.windowSizer = wx.BoxSizer()
        self.windowSizer.Add(self.p1, 1, wx.ALL | wx.EXPAND)  


        self.lst1 = tablename_formatted
        self.lst2 = self.get_lst2()

        self.st = wx.ComboBox(self.p1, choices = self.lst1, style=wx.TE_PROCESS_ENTER)
        self.st2 = wx.ComboBox(self.p1, choices = self.lst2, style=wx.TE_PROCESS_ENTER, pos=(-1,100), size=(500,-1))
        
        self.button = wx.Button(self.p1, label="Save")

        # Set sizer for the panel content
        self.sizer = wx.GridBagSizer(8, 8)
        self.sizer.Add(self.titlename, (0, 0))
        self.sizer.Add(self.editname, (0, 1))
        self.sizer.Add(self.typename, (1, 0))
        self.sizer.Add(self.edittype, (1, 1))
        #self.sizer.Add(self.button, (2, 0), (1, 2), flag=wx.EXPAND)
        self.sizer.Add(self.st, (3, 0))
        self.sizer.Add(self.st2, (4, 0))
        self.sizer.Add(self.button, (5, 0), flag=wx.EXPAND)
        # Set simple sizer for a nice border
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        # Use the sizers
        self.p1.SetSizerAndFit(self.border)  
        self.SetSizerAndFit(self.windowSizer)  

        # Set event handlers
        #self.button.Bind(wx.EVT_BUTTON, self.OnButton)

        self.st.Bind(wx.EVT_COMBOBOX, self.update)
        self.st2.Bind(wx.EVT_COMBOBOX, self.func)

        self.button.Bind(wx.EVT_BUTTON, self.OnButton)
    def OnButton(self, e):
        x=self.OnCombo(None)
        #print(x)
        #print(self.editname.GetValue())
        print('[\n{{\n"stubs": "{}",\n"title": "{}",\n"type": "{}"\n}}\n]'.format(x,self.editname.GetValue(),self.edittype.GetValue()))
        #with open("C:\\Users\\foggi\\Documents\\Fred.txt", "w") as fp:
        with open("S:\\building_json_for_nemsdv\\sample1.json", "w") as fp:
            fp.write('[\n{{\n"stubs": "{}",\n"title": "{}",\n"type": "{}"\n}}\n]'.format(x,self.editname.GetValue(),self.edittype.GetValue()))

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

    def OnCombo(self, event):
        table_num = int(self.st.GetValue().split(' ')[0])
        row_num = int(self.st2.GetValue().split(' ')[0])
        return (stub_name[(table_num,row_num)])

    def func(self,event):
        x=self.OnCombo(None)
        #print (x)

class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, -1, 'get_stubs.py')
        frame.Show()
        self.SetTopWindow(frame)
        return 1


if __name__ == "__main__":
    app = MyApp(1)
#    wx.lib.inspection.InspectionTool().Show()
    app.MainLoop()


