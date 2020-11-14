import xlwings as xw
import numpy as np
import pylab as pl

class f():
    def __init__(self,lsee,rsee):
        self.x1,self.y1=lsee
        self.x2,self.y2=rsee
    def cond1(self):
        return([self.x1**3,self.x1**2,self.x1,1,self.y1])
    def cond2(self):
        return([self.x2**3,self.x2**2,self.x2,1,self.y2])
    def diff1(self):
        return([3*(self.x1**2),2*self.x1,1,0])
    def diff2(self):
        return([6*self.x1,2,0,0])
def neg(a):
    return(-a)
def Print():
    for j in array:
        print(j)

def nodePlot():
    for i in all_see:
        if i in see:
            pl.plot(i[0], i[1], 'o', color = 'red')
        else:
            pl.plot(i[0], i[1], 'o', color = 'blue')

    
wb=xw.Book(r'插值點+斜率+所有點.xlsx')#路徑
sheet1=wb.sheets['插值點']#
sheet3=wb.sheets['全部點']#
sheet4=wb.sheets['結果']#
for sample in range(15, 16):
    min_output = []
    see=[];sp=22#
    while True:
        spline=sheet1.cells(sample,sp).value
        if spline==None:
            break
        sp+=1
        exec('see.append('+spline+')')
    n=len(see)
    all_see=[];sp=1
    while True:
        spline=sheet3.cells(sample,sp).value
        if spline==None:
            break
        sp+=1
        exec('all_see.append('+spline+')')
    for i in range(n-1):
        exec('f'+str(i)+'=f('+str(see[i])+','+str(see[i+1])+')')
    array=[]
    for i in range(4*n-4):
        array.append([])
        for j in range(4*n-3):
            array[i].append(0)
    for i in range(n-1):
        exec('f'+str(i)+'=f('+str(see[i])+','+str(see[i+1])+')')
    for i in range(n-1):
        if i==0:
            array[0][0:4]=f0.cond1()[0:4];array[0][-1]=f0.cond1()[-1]
            array[1][0:4]=f0.cond2()[0:4];array[1][-1]=f0.cond2()[-1]
            array[2][0:4]=[6*see[0][0],2,0,0];array[2][-1]=0
            array[3][-5:-1]=[6*see[-1][0],2,0,0];array[3][-1]=0
        else:
            exec('array[i*4][i*4:i*4+4]=f'+str(i)+'.cond1()[0:4];array[i*4][-1]=f'+str(i)+'.cond1()[-1]')
            exec('array[i*4+1][i*4:i*4+4]=f'+str(i)+'.cond2()[0:4];array[i*4+1][-1]=f'+str(i)+'.cond2()[-1]')
            exec('array[i*4+2][i*4-4:i*4+4]=f'+str(i)+'.diff1()+list(map(neg,f'+str(i)+'.diff1()))')
            exec('array[i*4+3][i*4-4:i*4+4]=f'+str(i)+'.diff2()+list(map(neg,f'+str(i)+'.diff2()))')
    #-------------------------------------------------------------
    del_list=[]
    while True:
        situ=0
        for i in range(n*4-4):
            if array[i][i]==0:
                break
            if i==n*4-5:
                situ=1
                break
        if situ:
            break
        ave=[]
        for i in range(n*4-4):
            ave.append([])
            if i in del_list:
                continue
            for j in range(n*4-4):
                if j in del_list:
                    continue
                if array[j][i]!=0:
                    ave[i].append(j)
        temp=[]
        for i in range(n*4-4):
            if len(ave[i])!=0:
                temp.append(len(ave[i]))
        for i in range(n*4-4):
            if len(ave[i])==min(temp):
                array[i],array[ave[i][0]]=array[ave[i][0]],array[i]
                del_list.append(i)
                break
    #Print()
    #-------------------------------------------------------------
    for i in range(len(array)):
        for j in range(len(array)):
            if i==j:
                continue
            temp=(-array[j][i])/array[i][i]
            for k in range(len(array[0])):
                array[j][k]=array[i][k]*temp+array[j][k]
    #-------------------------------------------------------------
    output=[]
    for i in range(len(array)):
        output.append(array[i][-1]/array[i][i])
    #-------------------------------------------------------------
    accuracy=[]
    for i in all_see:
        for j in range(n-1):
            if see[j][0]<=i[0]<=see[j+1][0]:
                break
        accuracy.append(abs(i[0]**3*output[j*4]+i[0]**2*output[j*4+1]+i[0]**1*output[j*4+2]+output[j*4+3]-i[1]))#cube spline

    #sheet4.cells(sample,25).value=sum(accuracy)#

    for i in range(n - 1):
        a, b = sorted([see[i][0], see[i + 1][0]])
        x = np.linspace(a, b, 100)
        y = output[i * 4] * (x ** 3) + output[i * 4 + 1] * (x ** 2) + output[i * 4 + 2] * x + output[i * 4 + 3]

        pl.plot(x, y, color = 'blue')

    nodePlot()

    pl.show()
