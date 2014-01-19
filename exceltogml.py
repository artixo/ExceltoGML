import xlrd, Tkinter, tkFileDialog
root = Tkinter.Tk()
excel = tkFileDialog.askopenfile(parent=root,mode='rb',title='Choose a file',
                                 filetypes= [('Excel 2010-2013 Worksheet', '*.xlsx'),('Excel 2003-2007 Worksheet', '*.xls')])
book = xlrd.open_workbook(str(excel.name))
myFormats = [
    ('GameMaker Script','*.gml'),
    ]

root = Tkinter.Tk()
fileName = str(tkFileDialog.asksaveasfilename(parent=root,filetypes=myFormats ,title="Save the image as..."))+".gml"

newfile = open(str(fileName), 'w')

for k in range(0, (book.nsheets)):
    sh = book.sheet_by_index(k)
    for i in range(1, sh.nrows):
        for j in range(0, sh.ncols):
            newfile.write( str(sh.name) + "["+str(i-1)+","+str(j)+"] = " +str(sh.cell_value(i, j))+";")
            if i==1:
                newfile.write("\t//"+str(sh.cell_value(0, j))+"\n")
            else:
                newfile.write("\n")
newfile.close()
newfile = open(str(fileName), 'r')
print newfile.read()
