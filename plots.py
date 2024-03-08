import pandas as pd 
import matplotlib.pyplot as plt
import openpyxl

#Print Statements
print("Welcome to the Analysis by Nawall. ")
ticker= input("Enter the Required Ratio for Histogram: ")
ticker = str(ticker)
date = input("Enter the Required Year for Histogram: ")
date=str(date)
ticker1= input("Enter the Required Ratio for Line Plot: ")
ticker1= str(ticker1)
company_name= input("Enter the Required Company for Line Plot: ")
company_name=str(company_name)

#Reading the Data from Excel
data_AC = pd.read_excel('Financial Statements (AC.MX).xlsx', sheet_name='Ratio')
data_AEFES= pd.read_excel('Financial Statements (AEFES.IS).xlsx', sheet_name='Ratio')
data_ANDINA= pd.read_excel('Financial Statements (AKOA).xlsx', sheet_name='Ratio')
data_BUD = pd.read_excel('Financial Statements (BUD).xlsx', sheet_name='Ratio')
data_CCEP = pd.read_excel('Financial Statements (CCEP.A).xlsx', sheet_name='Ratio')
data_CCHGY = pd.read_excel('Financial Statements (CCHGY.PK).xlsx', sheet_name='Ratio')
data_CCOLA = pd.read_excel('Financial Statements (CCOLA.IS).xlsx', sheet_name='Ratio')
data_KO = pd.read_excel('Financial Statements (KO).xlsx', sheet_name='Ratio')
data_KOF = pd.read_excel('Financial Statements (KOF).xlsx', sheet_name='Ratio')
data_PEP = pd.read_excel('Financial Statements (PEP.O).xlsx', sheet_name='Ratio')

#Setting the Values in the Analysis Ratio Column as Indexes
data_AC.set_index('Analysis Ratios', inplace=True)
data_AEFES.set_index('Analysis Ratios', inplace=True)
data_ANDINA.set_index('Analysis Ratios', inplace=True)
data_BUD.set_index('Analysis Ratios', inplace=True)
data_CCEP.set_index('Analysis Ratios', inplace=True)
data_CCHGY.set_index('Analysis Ratios', inplace=True)
data_CCOLA.set_index('Analysis Ratios', inplace=True)
data_KO.set_index('Analysis Ratios', inplace=True)
data_KOF.set_index('Analysis Ratios', inplace=True)
data_PEP.set_index('Analysis Ratios', inplace=True)

#Variables for Histogram and Converting them into INT
ac= data_AC.loc[f'{ticker}',f'{date}']
aefes= data_AEFES.loc[f'{ticker}',f'{date}']
andina= data_ANDINA.loc[f'{ticker}',f'{date}']
bud= data_BUD.loc[f'{ticker}',f'{date}']
ccep= data_CCEP.loc[f'{ticker}',f'{date}']
cchgy= data_CCHGY.loc[f'{ticker}',f'{date}']
ccola= data_CCOLA.loc[f'{ticker}',f'{date}']
ko= data_KO.loc[f'{ticker}',f'{date}']
kof= data_KOF.loc[f'{ticker}',f'{date}']
pep= data_PEP.loc[f'{ticker}',f'{date}']


#Variables for Line Plot and Converting them into INT 
ac1= data_AC.loc[f'{ticker1}',f'2022']
ac2= data_AC.loc[f'{ticker1}',f'2022']
ac3= data_AC.loc[f'{ticker1}',f'2022']

aefes1= data_AEFES.loc[f'{ticker1}',f'2022']
aefes2= data_AEFES.loc[f'{ticker1}',f'2021']
aefes3= data_AEFES.loc[f'{ticker1}',f'2020']

andina1= data_ANDINA.loc[f'{ticker1}',f'2022']
andina2= data_ANDINA.loc[f'{ticker1}',f'2021']
andina3= data_ANDINA.loc[f'{ticker1}',f'2020']

bud1= data_BUD.loc[f'{ticker1}',f'2022']
bud2= data_BUD.loc[f'{ticker1}',f'2021']
bud3= data_BUD.loc[f'{ticker1}',f'2020']

ccep1= data_CCEP.loc[f'{ticker1}',f'2022']
ccep2= data_CCEP.loc[f'{ticker1}',f'2021']
ccep3= data_CCEP.loc[f'{ticker1}',f'2020']

cchgy1= data_CCHGY.loc[f'{ticker1}',f'2022']
cchgy2= data_CCHGY.loc[f'{ticker1}',f'2021']
cchgy3= data_CCHGY.loc[f'{ticker1}',f'2020']

ccola1= data_CCOLA.loc[f'{ticker1}',f'2022']
ccola2= data_CCOLA.loc[f'{ticker1}',f'2021']
ccola3= data_CCOLA.loc[f'{ticker1}',f'2020']

ko1= data_KO.loc[f'{ticker1}',f'2022']
ko2= data_KO.loc[f'{ticker1}',f'2021']
ko3= data_KO.loc[f'{ticker1}',f'2020']

kof1= data_KOF.loc[f'{ticker1}',f'2022']
kof2= data_KOF.loc[f'{ticker1}',f'2021']
kof3= data_KOF.loc[f'{ticker1}',f'2020']

pep1= data_PEP.loc[f'{ticker1}',f'2022']
pep2= data_PEP.loc[f'{ticker1}',f'2021']
pep3= data_PEP.loc[f'{ticker1}',f'2020']


#List for Histogram
list1= [ac,aefes,andina,bud,ccep,cchgy,ccola,ko,kof,pep]
#List for Line Plot
if company_name== 'AC.MX':
    list2=[ac1,ac2,ac3]
elif company_name== 'AEFES.IS':
    list2=[aefes1,aefes2,aefes3]
elif company_name== 'AKOA':
    list2=[andina1,andina2,andina3]
elif company_name== 'BUD':
    list2=[bud1,bud2,bud3]
elif company_name== 'CCEP.A':
    list2=[ccep1,ccep2,ccep3]
elif company_name== 'CCHGY.PK':
    list2=[cchgy1,cchgy2,cchgy3]
elif company_name== 'CCOLA.IS':
    list2=[ccola1,ccola2,ccola3]
elif company_name== 'KO':
    list2=[ko1,ko2,ko3]
elif company_name== 'KOF':
    list2=[kof1,kof2,kof3]
elif company_name== 'PEP.O':
    list2=[pep1,pep2,pep3]

#Plotting HISTOGRAM
companies = ['Arca Continental', 'Anadolu EFES', 'Andina', 'Anheuser-Busch', 'Euro-Pacific', 'Hallanic', 'Icecek', 'The Coca Cola Comp', 'Femsa', 'Pepsi']
plt.bar(companies, list1)
plt.title(f'{ticker} for Fiscal Year {date}')
plt.xlabel('Company')
plt.ylabel('Values')
plt.show()

#Plotting LINE PLOT
years = [2022, 2021, 2020, 2019, 2018]
plt.plot(years, list2, marker='o')
plt.xlabel('Year')
plt.ylabel('Ratios')
plt.title(f"{company_name}'s {ticker1} over Time")
plt.show()