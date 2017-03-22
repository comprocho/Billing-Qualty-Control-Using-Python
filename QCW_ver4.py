
import pandas as pd
import numpy as np 
import glob as glob
import re

file_name = [] #Creates an empty list that will contain names of the file
prices = [] # Creates an empty list that will contain prices
for f in glob.glob("documents/*.xlsx"):
    if f.endswith('HTB.xlsx'): #Condition that will filter the file that has a name ending with HTB
        df_htb = pd.read_excel(f, skiprows=2)
        total_price = df_htb['Total'].loc[26]
        file_name.append(f)
        prices.append(total_price)
    elif f.endswith('IDM.xlsx'): #Condition that will filter the file that has a name ending with IDM
        df_idm = pd.read_excel(f, skiprows = 6)
        total_price = df_idm['Total'].loc[31]
        file_name.append(f)
        prices.append(total_price)
    else:
        df = pd.read_excel(f,skiprows = 3) #Filter the remaining files that has the name with no HTB or IDM
        total_price = df['Total'].loc[40]
        file_name.append(f)
        prices.append(total_price)

worktags = pd.DataFrame({'Worktag': file_name, 'Price': prices})
worktags = worktags[['Worktag', 'Price']]

worktags_split = worktags['Worktag'].str.replace('(.xlsx)', '').str.split(" ", 1, expand = True)
worktags['Worktag'] = worktags_split[1]
worktags['Price'] = worktags['Price'].astype(float).round(2)

replacements = {'Worktag' : {r'(HTB|IDM|)': ''}} #Create a dictionary that will be used for dropping substrings 'HTB' and 'IDM'
worktags.replace(replacements, regex= True, inplace=True) #Drop substrings 'HTB' and 'IDM' from 'Worktag' column

worktags.Worktag = worktags.Worktag.str.split().str.join(' ')

invoices = pd.read_excel('documents/Invoice Details.xls', skiprows=8)
invoices = invoices[['Worktags', 'Fee']]
invoices.rename(columns={'Worktags': 'Invoice Worktags', 'Fee': 'Price'}, inplace=True)
replacements =  {'Invoice Worktags' :
                    {
                    r'(GMS Center-|Fund-|Program-|Purpose Code-|Assignee-|Grant-|Gift-|Project-|Lab Support|-|\(|\))': '',
                    r'(,)': ' '
                    }
                }
invoices.replace(replacements, regex= True, inplace=True)
invoices = invoices.apply(lambda x: x.astype(str).str.upper())

sep = 'CC'
rest = invoices['Invoice Worktags'].str.split(sep, 1, expand = True)
invoices['Invoice Worktags'] = rest[1]
invoices['Invoice Worktags'] = 'CC' + invoices['Invoice Worktags'].astype(str)
invoices = invoices[:-1]
invoices['Price'] = invoices['Price'].astype(float).round(2)

frames = [worktags, invoices]
result = pd.concat(frames, axis=1)

wt = worktags.set_index('Price').Worktag.str.split(expand=True).stack().to_frame('Worktag').reset_index().rename(columns={'level_1':'idx'})
iv = invoices.set_index('Price')['Invoice Worktags'].str.split(expand=True).stack().to_frame('Invoice Worktags').reset_index().rename(columns={'level_1':'idx'})

no_match_idx = wt.loc[~wt.Worktag.isin(iv['Invoice Worktags']), 'idx'].unique()
no_match_idx

worktags['Matched_ID'] = ~worktags.index.isin(no_match_idx)
worktags['Matched_Value'] = worktags.Price.isin(invoices.Price)

worktags.to_excel('results.xlsx')
