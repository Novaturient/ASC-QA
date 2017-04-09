# -*- coding: utf-8 -*-
"""
Created on Sun Mar 26 13:36:41 2017

@author: cloud
"""

import pandas as pd
import time
import numpy as np
#import xlsxwriter as xls


def readInpFiles(inFile,surveyRange,sheet=0):
    print "Reading {}...".format(inFile)
    df = pd.read_excel(inFile,sheetname=sheet,skiprows=[0]) # Skip primary header "Furniture Survey Items"
    
    # Get range of surveyors from csv
    surveyRange = pd.read_csv(surveyRange)

    # Get all range per NAME
    dicNames = {}
    for i,value in surveyRange.iterrows():
        allRange = []
        valRange = value.Range.split(';')
        numRange = len(valRange)
        for j in range(numRange):
            allRange += xrange(int(valRange[j].split('-')[0]),int(valRange[j].split('-')[1])+1)
        dicNames[value.Name] = allRange # Save all range to mapping dic
    return df,dicNames

def findMissingPCN(df,dicNames):
    # Use mapping dic
    missingPCN = []
    for name in dicNames:
        nameDf = df[df['Survey Code'].str.contains("{}".format(name))==True]
        
        if nameDf['Property Control Number'].dtype == 'int64':
            for pcn in dicNames[name]: # Get pcn from allRange for each surveyor
                # Check if that pcn is in the data, if not, tag as missing
                if not pcn in list(nameDf['Property Control Number']):
                    missingPCN.append({'Survey Code':name,'Property Control Number':str(pcn).zfill(12),\
                    'ARCHIBUS Comments':'Missing PCN','ARCHIBUS QA Status':'Discrepancy found'})
                    print 'missing: {}'.format(pcn)
        else:
            for pcn in dicNames[name]:
                # Check if that pcn is in the data, if not, tag as missing
                if not str(pcn).zfill(12) in list(nameDf['Property Control Number']):
                    missingPCN.append({'Survey Code':name,'Property Control Number':str(pcn).zfill(12),\
                    'ARCHIBUS Comments':'Missing PCN','ARCHIBUS QA Status':'Discrepancy found'})
                    print 'missing: {}'.format(pcn)
                
            
    if not len(missingPCN) == 0:
        missingDf = pd.DataFrame(missingPCN)
    else:
        missingDf = pd.DataFrame(columns=['Survey Code','Property Control Number','ARCHIBUS Comments','ARCHIBUS QA Status'])
    print "Total missing PCN: {}".format(len(missingPCN))
    return missingDf
    
def findDuplicatePCN(df):
    # Zero pad PCN
    df['Property Control Number'] = df['Property Control Number'].map(lambda x:str(x).zfill(12))
    
    # Check duplicates
    duplicated = df.duplicated(subset=['Property Control Number'],keep=False)
    ddf = df[duplicated == True]
    
    # Add ARCHIBUS Comments based on filters
    fdf = pd.merge(df,ddf,how='left',indicator='ARCHIBUS Comments')
    fdf['ARCHIBUS Comments'] = fdf['ARCHIBUS Comments'].astype('str')
    fdf.loc[pd.isnull(fdf['Building Code'])==True,'ARCHIBUS Comments'] = 'No Location'
    fdf.loc[pd.isnull(fdf['Floor Code'])==True,'ARCHIBUS Comments'] = 'No Location'
    fdf.loc[pd.isnull(fdf['Room Code'])==True,'ARCHIBUS Comments'] = 'No Location'
    fdf.loc[fdf['ARCHIBUS Comments']=='both','ARCHIBUS Comments'] = 'Duplicate PCN'
    fdf.loc[fdf['ARCHIBUS Comments']=='left_only','ARCHIBUS Comments'] = ''
    
    # Add furniture type
#    fdf.loc[pd.isnull(fdf['Unifying Code']),'Furniture Type'] = 'As is Furniture'
    fdf['Furniture Type'] = np.where(pd.isnull(fdf['Unifying Code']),'As is Furniture','Component')
    
    # Add ARCHIBUS QA Status
    fdf.loc[fdf['ARCHIBUS Comments']!='','ARCHIBUS QA Status'] = 'Discrepancy found'
    newCol = fdf.columns.tolist()[:-3] + fdf.columns.tolist()[-2:-1] + fdf.columns.tolist()[-1:] + [fdf.columns.tolist()[-3]]
    fdf = fdf[newCol]
    return fdf
    
def combineDf(fdf,missingDf):
    finalDf = fdf.append(missingDf,ignore_index=True)
    finalDf = finalDf[fdf.columns]
    return finalDf
    
def writeOutput(finalDf,inFile):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('QA_{}'.format(inFile), engine='xlsxwriter')
    
    # Convert the dataframe to an XlsxWriter Excel object.
    finalDf.to_excel(writer,index=False, sheet_name='Sheet1')
    
    
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Format
    format1 = workbook.add_format({'bg_color':   '#FFC7CE', #red
                                   'font_color': '#9C0006'})
    format2 = workbook.add_format({'bg_color':   '#FFEB9C', #yellow
                                   'font_color': '#9C6500'})
    
    # Condition for PCN
    # Change column letter based on the given worksheet
    ld = len(finalDf)+2
    worksheet.conditional_format('B2:B%s'%ld, {'type':   'duplicate',
                                           'format': format2})
    
    # Condition for AC and AQA
    # Change column letter based on the given worksheet
    # ARCHIBUS Comments
    worksheet.conditional_format('S2:S%s'%ld, {'type':     'text',
                                           'criteria': 'containing',
                                           'value':    'No Location',
                                           'format':   format1})
    
    worksheet.conditional_format('S2:S%s'%ld, {'type':     'text',
                                           'criteria': 'containing',
                                           'value':    'Duplicate PCN',
                                           'format':   format1})
    
    worksheet.conditional_format('S2:S%s'%ld, {'type':     'text',
                                           'criteria': 'containing',
                                           'value':    'Missing PCN',
                                           'format':   format1})
    
    # ARCHIBUS QA Status
    worksheet.conditional_format('R2:R%s'%ld, {'type':     'text',
                                           'criteria': 'containing',
                                           'value':    'Discrepancy found',
                                           'format':   format1})
    
    # Add border
    border = workbook.add_format({'border':1})
    
    worksheet.set_column(0,len(finalDf.columns)-1,18,border)
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    print "Saved QA as QA_{}...".format(inFile)

def main(inFile,surveyRange,sheet):
    df,dicNames = readInpFiles(inFile,surveyRange,sheet)
    missingDf = findMissingPCN(df,dicNames)
    fdf = findDuplicatePCN(df)
    finalDf = combineDf(fdf,missingDf)
    writeOutput(finalDf,inFile)
    return df,dicNames,missingDf
    
if __name__ == '__main__':
    inFile = 'MAR3-P4.xlsx'
    surveyRange = 'surveyRangeMar25.csv'
    # inFile = 'Book1.xlsx'
    sheet = 0
    t0 = time.time() 
    df,dicNames,missingDf=main(inFile, surveyRange, sheet)
#    f,m = main(inFile, surveyRange, sheet)
    t1 = time.time()
    dt = t1-t0
    print "Finished processing in {} seconds.".format(round(dt,4))
    raw_input('Press Enter to exit...')

