import xlrd
import matplotlib.pyplot as pt
import Tkinter as tk


wb=xlrd.open_workbook('parsed.xls')
ws=wb.sheet_by_index(0)




#function which plots the graph
def plot(teacher,subjectname):
    name=teacher
    subject=subjectname

    t_score=[]
    year=[]




    #getting all total score values for all teachers
    for i in range(ws.nrows):
        if ws.cell_value(i,0)==name and ws.cell_value(i,1)==subject:
            for j in range(ws.ncols):


                if ws.cell(i, j).ctype == xlrd.XL_CELL_EMPTY:
                    t_score.append(ws.cell_value(i, j - 1))
                    break


                if j==ws.ncols-1:
                    t_score.append(ws.cell_value(i, j))



            year.append(ws.cell_value(i,2))


    #plotting the graph from data extracted
    pt.plot(year,t_score ,marker='o' ,color='r',label=subject)
    pt.axis([2000, 2019 , 70,100])
    pt.ylabel('Total score')
    pt.xlabel('Year')
    pt.title(name)
    pt.legend(loc='upper left')
    pt.show()




#getting the subject name and teacher name from the GUI and calling plot function
def get():
    subject=variable.get()
    teacher_name=var2.get()
    plot(teacher_name,subject)



#getting all teachers and subjects from excel sheet for display in dropdown
sublist =set(ws.cell_value(i,1) for i in range(ws.nrows))
teachers=set(ws.cell_value(i,0) for i in range(ws.nrows))




root=tk.Tk()


var2=tk.StringVar(root)
var2.set('None')
dropdown_1=apply(tk.OptionMenu,(root,var2)+tuple(teachers))
dropdown_1.pack()


variable=tk.StringVar(root)
variable.set('None')
dropdown=apply(tk.OptionMenu,(root,variable)+tuple(sublist))
dropdown.pack()



button=tk.Button(root,text='Submit' ,command=get)
button.pack()


root.mainloop()









