
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

years = []
years1 = []
stubgrp = []
# operationanme = []
# operationsymbol = []
operation_name_sym = {}
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

        #optname
        self.optname = wx.StaticText(self.p1,label="Optname:")
        self.editopt = wx.TextCtrl(self.p1, size=(200, -1), style=wx.TE_PROCESS_ENTER)
        
        #optsym
        self.optsymname = wx.StaticText(self.p1,label="Optsym:")
        self.editoptsym = wx.TextCtrl(self.p1, size=(140, -1), style=wx.TE_PROCESS_ENTER)

        #optsym
        self.yearsname = wx.StaticText(self.p1,label="years:")
        self.edityears = wx.TextCtrl(self.p1, size=(140, -1), style=wx.TE_PROCESS_ENTER)

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

        self.sizer.Add(self.optname, (2, 0))
        self.sizer.Add(self.optsymname, (2, 1))
        self.sizer.Add(self.editopt, (3, 0))
        self.sizer.Add(self.editoptsym, (3, 1))

        self.sizer.Add(self.yearsname, (4, 0))
        self.sizer.Add(self.edityears, (4, 1))

        self.sizer.Add(self.st, (5, 0))
        self.sizer.Add(self.st2, (6, 0))
        self.sizer.Add(self.button, (7, 0), flag=wx.EXPAND)
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

        self.editopt.Bind(wx.EVT_TEXT_ENTER, self.update2)
        self.editoptsym.Bind(wx.EVT_TEXT_ENTER, self.func2)
        self.edityears.Bind(wx.EVT_TEXT_ENTER, self.func4)
        self.button.Bind(wx.EVT_BUTTON, self.OnButton)
    def OnButton(self, e):
        x=self.OnCombo(None)
        #print(x)
        #print(self.editname.GetValue())
        #print (operationsymbol)
        print('[\n{\n"stubs": ['+ ',\n '.join('"{0}"'.format(w) for w in stubgrp)+'\n],')
        #print('[\n{{\n"stubs": ['"{0}"'\n],'.format(',\n'.join(stubgrp)))
        print('"'+"names"+'":{')
        print((',\n '.join('"{}":{}'.format(x,y) for x, y in operation_name_sym.items())+'\n},').replace("'",'"'))

        #print(",\n".join("{}: [{}]".format(x,y) for x, y in zip(operationanme, operationsymbol)))



        #print(",\n".join("{}: [{}]".format(x, y) for x, y in zip(operationanme, operationsymbol)))
        #print("},")
        #print('"names": "{}:{}"\n}},'.format(x, y) for x, y in zip(operationanme, operationsymbol))
        #print("\n".join("{}: [{}]\n}},\n".format(x, y) for x, y in zip(operationanme, operationsymbol)))
        #print('"names":{{\n {}:{}\n}}\n'.format('\n'.join(*operationanme,*operationsymbol)))
        #print(years[0])
        print('"years": ['+ ','.join('"{0}"'.format(ww) for ww in years[0])+'],')
        print('"title": "{}",\n"type": "{}"\n}}\n]'.format(self.editname.GetValue(),self.edittype.GetValue()))

        #with open("C:\\Users\\foggi\\Documents\\Fred.txt", "w") as fp:
        with open("S:\\building_json_for_nemsdv\\sample1.json", "w") as fp:
            fp.write('[\n{\n"stubs": ['+ ',\n '.join('"{0}"'.format(w) for w in stubgrp)+'\n],')
            fp.write('"'+"names"+'":{')
            fp.write((',\n '.join('"{}":{}'.format(x,y) for x, y in operation_name_sym.items())+'\n},').replace("'",'"'))
            fp.write('"title": "{}",\n"type": "{}"\n}}\n]'.format(self.editname.GetValue(),self.edittype.GetValue()))
            
            #fp.write('[\n{{\n"stubs": [{}\n],\n'.format('\n'.join(stubgrp)))

            #fp.write('"names":{{\n "{}":["{}"]\n],\n'.format('\n'.join(stubgrp),'\n'.join(stubgrp)))
            #fp.write("names:{{\n".join("{}: [{}]\n}},\n".format(x, y) for x, y in zip(operationanme, operationsymbol)))

            #fp.write('"title": "{}",\n"type": "{}"\n}}\n]'.format(self.editname.GetValue(),self.edittype.GetValue()))

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
        stubgrp.append(x)
        #print (x)

    def update2(self,event):
        self.editoptsym.Clear()

    def OnCombo2(self, event):
        a=self.editopt.GetValue()
        b=list(self.editoptsym.GetValue().split(','))
        #print (b)
        return (a,b)

    def func2(self,event):
        #print (self.OnCombo2(None)[0])
        #p,q =self.OnCombo2(None)
        # operationanme.append(self.OnCombo2(None)[0])
        # operationsymbol.append(self.OnCombo2(None)[1])
        # print (operationsymbol)
        #print (self.OnCombo2(None)[1])
        operation_name_sym[self.OnCombo2(None)[0]] = self.OnCombo2(None)[1]
        #print (operation_name_sym)


    def func3(self,event):
        years11=self.edityears.GetValue()
        # print (years11)
        # print (years11.split(","))
        # print (list(years11.split(",")))
        return (years11)

    def func4(self,event):
        
        years.append(self.func3(None).split(","))
        print (years[0])
        return years[0]

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

#print (years,years1)

