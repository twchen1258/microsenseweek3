'''
A python program that takes csv data file as input and output an Excel file of fill with an array
Created on 08/10/2022
Programmed by David Chen
'''
import numpy as np
import xlwings as xw
from math import sqrt
import matplotlib.pyplot as plt


def csv_to_array(fileName):
    with open(fileName) as f:
        n_cols = len(f.readline().split(","))
    temp = np.loadtxt(fileName, delimiter=",", usecols=np.arange(4, n_cols-1))
    row, _ = temp.shape
    try:
        dimension = int(sqrt(row))
    except:
        print("Data does not have square dimension!")
    else:
        result = np.empty((dimension, 0), int)
        valueList = []
        for data in temp:
            top = int(np.argmax(data[0:32]))*16 + int(np.argmax(data[32:48]))
            valueList.append([top])
            if (len(valueList)==dimension):
                valueList.reverse()
                result = np.append(result, np.array(valueList), axis=1)
                valueList = []
        
        return np.flip(result, 1)
    
def postExcel(value, fileName, sheetName = 'Sheet1'):
    wb = xw.Book(fileName)
    sheet = wb.sheets[sheetName]
    
    sheet.range('A1').value = value
    
def saveFig(dataArray, figName):
    depthmap_arry=dataArray[::-1]
    plt.figure(figsize=(9,7))     
    plt.pcolormesh(depthmap_arry,cmap='plasma_r')
    plt.title("Depth Map (64x64)")
    plt.colorbar()
    plt.clim(16, 79)
    plt.savefig(figName+'.jpg')

    
if __name__ == '__main__':
    dataArray = csv_to_array('rawdata.csv')
    postExcel(dataArray, 'Result_Sheet.xlsx')
    saveFig(dataArray, 'My Picture')
    print("Finished")
    print("This is main")

