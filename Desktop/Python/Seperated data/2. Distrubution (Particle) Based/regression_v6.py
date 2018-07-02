import numpy as np
import pandas as pd
import statsmodels.api as sm
import sys
import time
import datetime
import xlwt, xlrd
from xlutils.copy import copy as xl_copy
from tkinter.filedialog import askopenfilename

file=askopenfilename()
file_name='Predictions_'+datetime.datetime.now().strftime("%y-%m-%d-%H-%M")+'.xls'

print (file)

df=pd.read_csv(file, low_memory=False)	

#x1 = 'M-M0'
#x2 = 'R-R0'
#x3 = 'I-I0'

x1 = 'Ar'
x2 = 'H2'
x3 = 'I'

#y_n = ['Temp', 'Speed', 'MI', 'KE']

y_n = ['Temp', 'Speed']

x=df[[x1, x2, x3]]
x=sm.add_constant(x)

est=[]
predicted=[]
y=[]
measured=[]
diff=[]
error=[]

#for i in range(0,4):
for i in range(0,2):
	y.append(df[y_n[i]])
	est.append(sm.OLS(y[i],x).fit())	
	predicted.append(est[i].predict())
	diff.append(np.absolute(y[i]-predicted[i]))
	error.append((diff[i])/(y[i]))
	df['predicted '+y_n[i]]=predicted[i]
	df['error'+y_n[i]]=error[i]

writer=pd.ExcelWriter(file_name, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Data')

df2=pd.DataFrame({'Coeff 1': est[0].params})
df2['Coeff 2']=est[1].params

df2.to_excel(writer, sheet_name='Summary')

writer.save()
writer.close()

