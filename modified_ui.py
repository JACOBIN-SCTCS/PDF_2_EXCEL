
# IMPORTING THE NEEDED MODULES
import matplotlib.pyplot as plt
import xlrd
import sys
import pandas as pd
from PyQt4 import QtCore,QtGui
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt4agg import NavigationToolbar2QT as NavigationToolbar


wb=xlrd.open_workbook('parsed.xls')
ws=wb.sheet_by_index(0)


class Window(QtGui.QMainWindow):

    currentteacher=""
    currentsubject=""
    teacherdataset=""
    subjectdataset=""

    def __init__(self,df):
        super(Window,self).__init__()
        self.df=df
        self.setGeometry(50,50,400,400)
        self.setWindowTitle("FeedBack Plotter")
         
    
        self.tcb=QtGui.QComboBox(self)
        
        teachers_list=list(set(self.df['Teacher'].tolist()))
        self.tcb.addItems(teachers_list)
        currentteacher=self.tcb.currentText()
        self.tcb.currentIndexChanged.connect(self.teacherchange)

        teacherdataset=self.df[self.df.Teacher==currentteacher]

        self.scb=QtGui.QComboBox(self)
        subjects_list=list(set(teacherdataset['Subject'].tolist()))
        self.scb.addItems(subjects_list)
        currentsubject=self.scb.currentText()
       
        btn=QtGui.QPushButton("Submit",self)
        btn.clicked.connect(self.submit)
         
     
        self.figure=plt.figure(figsize=(15,5))
        self.canvas=FigureCanvas(self.figure)
       

        wid=QtGui.QWidget(self)
        self.setCentralWidget(wid)
        grid=QtGui.QGridLayout()
        grid.addWidget(self.tcb,1,1)
        grid.addWidget(self.scb,1,2)
        grid.addWidget(btn,1,3)
        grid.addWidget(self.canvas ,2,1,3,3)
       
        wid.setLayout(grid) 

        self.show()


    def teacherchange(self,i):
       currentteacher=self.tcb.currentText()
       self.scb.clear()
       teacherdataset=self.df[self.df.Teacher==currentteacher]
       subjects_list=list(set(teacherdataset['Subject'].tolist()))
       self.scb.addItems(subjects_list)
       
    def submit(self):
        currentteacher=self.tcb.currentText()
        currentsubject=self.scb.currentText()
        teacherdataset=self.df[self.df.Teacher==currentteacher]
        subjectdataset=teacherdataset[teacherdataset.Subject==currentsubject]
        list_years=subjectdataset['Year'].tolist()
        list_score=subjectdataset['TotalScore'].tolist()

        
        #plt.plot(list_years,list_score,marker='o',color='r')
        
        #plt.show()

        plt.cla()
        ax=self.figure.add_subplot(111)




        
        ax.plot(list_years,list_score,marker='o' ,color='r')
         
        ax.set_xlabel('Year')
        ax.set_ylabel('TotalScore')
        self.canvas.draw() 
    
        






       

t_score=[]
years=[]
teachers=[]
subject=[]



def main():
    for i in range(ws.nrows):
        for j in range(ws.ncols):
            if(ws.cell(i,j).ctype == xlrd.XL_CELL_EMPTY):
                t_score.append(ws.cell_value(i,j-1))
                break

            if(j==ws.ncols-1):
                t_score.append(ws.cell_value(i,j))
    

    for i in range(ws.nrows):
       years.append(ws.cell_value(i,2))
       teachers.append(ws.cell_value(i,0))
       subject.append(ws.cell_value(i,1))
    
    
    
    

    #THE BELOW CODES PRODUCES THE DATAFRAME FROM ALL (PANDAS)
    s_score=pd.Series(t_score)
    s_years=pd.Series(years)
    s_teachers=pd.Series(teachers)
    s_subject=pd.Series(subject)


    dataframe=pd.DataFrame({ 'Teacher':s_teachers , 'Subject':s_subject, 'Year':s_years,'TotalScore':s_score})
        
    app=QtGui.QApplication(sys.argv)
    window=Window(dataframe)
    sys.exit(app.exec_())




if __name__=='__main__':
 main()


