# -*- coding: cp936  -*-
# -*- coding: UTF-8  -*-
import Tkinter, Tkconstants, tkFileDialog  
import xlrd,re
import time,sys

class excel_search(Tkinter.Frame):  
    def __init__(self, root):  
        Tkinter.Frame.__init__(self, root)
        root.title(u'excel search v0.2 by 张珲')
        Tkinter.Button(self, text=u'打开excel', 
                       command=self.askopenfilename).pack(expand='no',anchor='nw')
        Tkinter.Button(self, text=u'检索结果另存为', 
                       command=self.savefile).pack(expand='no',anchor='ne') 
        Tkinter.Button(self, text=u'开始检索', 
                       command=self.searchinexcel).pack(expand='no',anchor='center')
        Tkinter.Button(self, text=u'clear', 
                       command=self.clearText).pack(expand='no',anchor='ne')
        self.sheetObj = {}
        self.searchstr = Tkinter.StringVar()
        self.textinput = Tkinter.Entry(self, width=50, textvariable = self.searchstr)
        self.searchstr.set(u"需搜索内容")
        self.textinput.pack(expand='no')
        self.text = Tkinter.Text(self,width=500,height=480)
        self.text.bind("<Control-Key-a>", self.selectText)
        self.text.pack(fill=Tkinter.BOTH)
        self.file_opt = options = {}  
        options['defaultextension'] = '.txt'  
        options['filetypes'] = [('xls', '.xls'), ('xlsx', '.xlsx')]  
        options['initialdir'] = 'D:\\Work\\Old PC disk E\\山西移动\\EPC\\output\\0730\\'  
        options['initialfile'] = '*.xlsx'  
        options['parent'] = root  
        options['title'] = 'Open excel file'  
    
    def open_excel(self,file = 'test.xlsx'):
        try:
            data = xlrd.open_workbook(file.encode('GBK'))
            data.sheetname = 'file'
            return data
        except Exception,e:
            print str(e)

    def strs(self,row):
        values = "";
        for i in range(len(row)):
            if i == len(row) - 1:
                values = values + str(row[i])
            else:
                values = values + str(row[i]) + " "
        return values
    def selectText(self, event):
        self.text.tag_add('sel', '1.0', 'end')
        return 'break'
    def clearText(self):
        self.text.delete(0.0, Tkinter.END)
        self.sheetObj = {}
    def excel_table_byindex(self,obj,sheetname,index='default',colnameindex=0,by_index=0):
        searchRE = '.*%s.*' %index
        allsheet = obj.sheets()
        for sheet_num in range(len(allsheet)):
            table = allsheet[sheet_num]
            nrows = table.nrows
            #ncols = table.ncols
            for row in range(nrows):
                colnames = table.row_values(row)
                row_value = self.strs(colnames)
                if re.findall(searchRE, row_value):
                    for i in re.findall(searchRE, row_value):
                        self.text.insert(1.0, "%s : %s\n" %(sheetname,i))
                        
    def searchinexcel(self):
        self.text.insert(1.0, u"正在检索"+'\n')
        now = time.time()
        self.file_opt['multiple']=1
        self.file_opt['initialfile'] = '*.txt' 
        self.file_opt['filetypes'] = [('xls', '.xls'), ('xlsx', '.xlsx')]  
        self.file_opt['initialdir'] = 'C:\\' 
        index = self.searchstr.get()
        if self.sheetObj == {}:
            self.text.insert(1.0, 'Please select excel first'+'\n')
        if index is not None:
            for sheet1 in self.sheetObj:
                self.excel_table_byindex(sheet1,self.sheetObj[sheet1],index)
            self.text.insert(1.0, "Total search %s excels;Time cost: %s\n" 
                             %(len(self.sheetObj),time.time()-now))
        self.text.insert(1.0, u"检索完毕"+'\n')
            
    def askopenfilename(self):
        self.file_opt['multiple']=1
        self.file_opt['filetypes'] = [('xls', '.xls'), ('xlsx', '.xlsx')]  
        self.file_opt['initialdir'] = 'D:\\Work\\Old PC disk E\\山西移动\\EPC\\output\\0730\\'  
        self.file_opt['initialfile'] = '*.xlsx' 
        self.filelist = []
        self.text.insert(1.0, u'正在加载中...'+'\n')
        filename = tkFileDialog.askopenfilename(**self.file_opt)
        if isinstance(filename,tuple):
            pass
        else:
            filename = re.findall('\{(.*?)\}',filename)
        for i in range(len(filename)):
            if filename[i]:
                self.filelist.append(filename[i])
                self.text.insert(1.0, filename[i]+'\n')
                obj = self.open_excel(filename[i])
                self.sheetObj[obj] = filename[i]
        self.text.insert(1.0, u'加载完毕,可继续添加或输入检索内容开始检索...目前已添加:'+str(len(self.sheetObj))+'\n')
          
    def savefile(self):
        self.file_opt['multiple']=None
        self.file_opt['initialfile'] = '*.txt' 
        self.file_opt['filetypes'] = [('txt', '.txt'), ('log', '.log')]  
        self.file_opt['initialdir'] = 'C:\\'  
        file_save = self.asksaveasfile()
        if file_save:
            file_save.writelines(self.text.get('0.0',Tkinter.END))
            file_save.close() 
    def asksaveasfile(self):  
        return tkFileDialog.asksaveasfile(mode='w', **self.file_opt)  
    
if __name__ == '__main__':  
    reload(sys)
    sys.setdefaultencoding('utf-8')
    root = Tkinter.Tk()
    root.geometry('1000x500')
    excel_search(root).pack()
    root.mainloop()  
