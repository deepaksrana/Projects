import pandas as pd
import numpy as np
import numpy_financial
import openpyxl
import xlsxwriter

class Invest():
    '''
    Docstring: This class reprersents each investment vehicle, i.e., potential investment
    '''
    def __init__(self,investment,tax_rate='21%',lifetime_years=20):
        self.investment = investment
        self.tax_rate = tax_rate
        self.lifetime_years = lifetime_years
        self.sales_growth = 0
        self.opex_growth = 0
        
        self.capex = [0]*(self.lifetime_years+1)
        self.sales = [0]*(self.lifetime_years+1)
        self.opex = [0]*(self.lifetime_years+1)
        self.ebit = [0]*(self.lifetime_years+1)
        self.nopat = [0]*(self.lifetime_years+1)
        self.income = [0]*(self.lifetime_years+1)
           
    def __str__(self):
        print(self.investment)
        
myfile = input("Please specify the file path including the file name: ") 
myworksheet  = input("Please specify the worksheet which has the inputs ")
    
with pd.ExcelFile(myfile) as xls:
    df_capex = pd.read_excel(xls, myworksheet, usecols = 'B:D', index_col=None,skiprows=1,keep_default_na=False)
    df_sales = pd.read_excel(xls, myworksheet, usecols = 'F:H', index_col=None,skiprows=1,keep_default_na=False)
    df_opex = pd.read_excel(xls, myworksheet, usecols = 'J:L', index_col=None,skiprows=1,keep_default_na=False)
    df_others = pd.read_excel(xls, myworksheet, usecols = 'N:O', index_col=None,skiprows=1,keep_default_na=False)
    
# To ensure the program doens't pick up values from the resulting income statement (columns below capex) if it is rerun 
for i,name in enumerate(df_capex.loc[ : , 'Capital Expenditure' ]): 
    if name:
        pass
    else:
        df_capex.drop(df_capex.index[i:len(df_capex)], inplace=True)
        break
 
 # To ensure the program doens't pick up values from the resulting income statement (columns below revenue)  if it is rerun
for i,name in enumerate(df_sales.loc[:,'Revenue']):
    if name:
        pass
    else:
        df_sales.drop(df_sales.index[i:len(df_sales)], inplace=True)
        break

# To ensure the program doens't pick up values from the resulting income statement (columns below opex)  if it is rerun
for i,name in enumerate(df_opex.loc[:,'Operating Expense']):
    if name:
        pass
    else:
        df_opex.drop(df_opex.index[i:len(df_opex)], inplace=True)
        break

# To ensure the program doens't pick up values from the resulting income statement (columns below others)  if it is rerun
for i,name in enumerate(df_others.loc[:,'Others']):
    if name:
        pass
    else:
        df_others.drop(df_others.index[i:len(df_others)], inplace=True)
        break
 

# to get the the ending row of inputs in excel so the resulting income statement can be plotted a few rows below
ending_row_num = max(len(df_capex),len(df_sales),len(df_opex),len(df_others)) + 1 

#filter Others dataframe to get value for tax rate and then assign it
tax_rate = (df_others.loc[df_others['Others'] == 'Federal Income Tax']) 
tax_rate = float(tax_rate.iloc[0,[1]])
#tax_rate = "{:.2%}".format(tax_rate) # to represent the tax rate in percetage terms

#filter Others dataframe to read and set the project lifetime in years: currently set as integer
lifetime_years = (df_others.loc[df_others['Others'] == 'Project Lifetime'])
lifetime_years = int(lifetime_years.iloc[0,[1]])

#Create a set of investment vehicles with no duplicate values and has all investments
veh_set = set()
# the commmeted code is no longer needed since the rerun problem has been taken care of in the code at the starting part of this program
##for i,name in enumerate(df_capex.Vehicle): ## So the program doesn't pick values from result matrix if it the program is run multiple times
##   if name:
##        veh_set.add(name)
##    else:
##        break
########## filter Remove blank values so they aren't counted as investments
filtered_vehicles = filter(None, df_capex.Vehicle)
for v in filtered_vehicles:
    veh_set.add(v)

# to ensure the correct order of investment names while creating the income stmt so correct paramaters can be mapped
# Current investment are presented in the income statement in reverse alphabetical order
veh_set_sorted = sorted(veh_set,reverse=True) 
    
#Convert set to tuple to make this is an unmutable list so as to avoid reference errors
v_tup=tuple(veh_set_sorted)

v1 = Invest(v_tup[0],tax_rate, lifetime_years)
v2 = Invest(v_tup[1],tax_rate, lifetime_years)

# Filter CapEx dataframe for Vehicle 1 values, i.e., Vehicle = 'Solar PV'
v1_capex = (df_capex.loc[df_capex['Vehicle'] == v1.investment]) 

#Iterate through the capex inputs and set the appropciarte capex values for Solar PV
v1_capex_dict = v1_capex.to_dict(orient='split')
for i,info in enumerate(v1_capex_dict['data']):
    v1.capex[info[1]] = info[2]

# Filter CapEx dataframe for Vehicle 2 values , i.e., Vehicle = 'BESS' 
v2_capex = (df_capex.loc[df_capex['Vehicle'] == v2.investment])

#Iterate through the capex inputs and set the appropciarte capex values for Solar PV
v2_capex_dict = v2_capex.to_dict(orient='split')
for i,info in enumerate(v2_capex_dict['data']):
    v2.capex[info[1]] = info[2]

# Covert capital expenditure numbers into negative to signify cash outflow for both investments  
v1.capex = [ -x for x in v1.capex]   
v2.capex = [ -x for x in v2.capex]   

# Filter sales dataframe for Vehicle 1 values, i.e., Vehicle = 'Solar PV'
v1_sales = (df_sales.loc[df_sales['Vehicle.1'] == v1.investment]) 
# Filter sales dataframe for Vehicle 2 values , i.e., Vehicle = 'BESS'
v2_sales = (df_sales.loc[df_sales['Vehicle.1'] == v2.investment]) 


# Iterate throught the sales inputs of Solar PV and project them into the project lifetime
# The inputs are set till the year they are provided, once the growth rate is encountered, the rest is projected
v1_sales_dict = v1_sales.to_dict(orient='split')
for i,value in enumerate(v1_sales_dict['data']):
    if (isinstance(value[1],int) or isinstance(value[1],float)) :
        v1.sales[value[1]] = value[2]
        growth_year_base=value[1]
    elif isinstance(value[1],str):
        v1.sales_growth = value[2]
        for num in range((growth_year_base+1),(v1.lifetime_years+1)):
            v1.sales[num] = round(v1.sales[num-1] * (1+ v1.sales_growth))

# Iterate throught the sales inputs of BESS and project them into the project lifetime
# The inputs are set till the year they are provided, once the growth rate is encountered, the rest is projected
v2_sales_dict = v2_sales.to_dict(orient='split')
for i,value in enumerate(v2_sales_dict['data']):
    if (isinstance(value[1],int) or isinstance(value[1],float)) :
        v2.sales[value[1]] = value[2]
        growth_year_base=value[1]
    elif isinstance(value[1],str):
        v2.sales_growth = value[2]
        for num in range((growth_year_base+1),(v2.lifetime_years+1)):
            v2.sales[num] = round(v2.sales[num-1] * (1+ v2.sales_growth))
            

# Filter operating expense dataframe for Vehicle 1 values, i.e., Vehicle = 'Solar PV'
v1_opex = (df_opex.loc[df_opex['Vehicle.2'] == v1.investment]) 
# Filter operating expense dataframe for Vehicle 2 values, i.e., Vehicle = 'BESS'
v2_opex = (df_opex.loc[df_opex['Vehicle.2'] == v2.investment]) 

# Iterate throught the opex inputs of Solar PV and project them into the project lifetime
# The inputs are set till the year they are provided, once the growth rate is encountered, the rest is projected
v1_opex_dict = v1_opex.to_dict(orient='split')
for i,value in enumerate(v1_opex_dict['data']):
    if (isinstance(value[1],int) or isinstance(value[1],float)) :
        v1.opex[value[1]] = value[2]
        growth_year_base=value[1]
    elif isinstance(value[1],str):
        v1.opex_growth = value[2]
        for num in range((growth_year_base+1),(v1.lifetime_years+1)):
            v1.opex[num] = round(v1.opex[num-1] * (1+ v1.opex_growth))

# Iterate throught the opex inputs of BESS and project them into the project lifetime
# The inputs are set till the year they are provided, once the growth rate is encountered, the rest is projected
v2_opex_dict = v2_opex.to_dict(orient='split')
for i,value in enumerate(v2_opex_dict['data']):
    if (isinstance(value[1],int) or isinstance(value[1],float)) :
        v2.opex[value[1]] = value[2]
        growth_year_base=value[1]
    elif isinstance(value[1],str):
        v2.opex_growth = value[2]
        for num in range((growth_year_base+1),(v2.lifetime_years+1)):
            v2.opex[num] = round(v2.opex[num-1] * (1+ v2.opex_growth))
            

# Create the columns for for the income statement in accordance with the projected lifetime of investtments (in years)
# Current Assumption is that both investments have the same lifetime
# If it that is not the case, then a function is to be written which would find the longest lifetime of amongst the investments
# then accordingly the columns would be created
years = []
for x in range(lifetime_years+1): years.append(x)

# Create the income statement dataframe using the years list created in the line above
income_stmt = pd.DataFrame(index=['Year','','Capital Expenditure','Solar PV','BESS','Total CapEx','','Revenue','Solar PV','BESS','Total Revenue','','Operating Expense','Solar PV', 'BESS', 'Total OpEx','','EBIT', 'Solar PV','BESS','Tax Rate','Solar PV: Tax','BESS: Tax','','NOPAT','Solar PV','BESS','Total NOPAT','','Net Cash Flow','Solar PV','BESS','Project net cash flow','','IRR'], columns=years)

income_stmt = income_stmt.fillna('') # Remove all NaN from the dataframe

income_stmt.iloc[0] = years  # Set years for columns headers
income_stmt.iloc[3] = v1.capex # Solar capex
income_stmt.iloc[4] = v2.capex # BESS capex
income_stmt.iloc[5] = income_stmt.iloc[3] + income_stmt.iloc[4] # Total Capex

income_stmt.iloc[8] = v1.sales # Solar Sales
income_stmt.iloc[9] = v2.sales # BESS Sales
income_stmt.iloc[10] = income_stmt.iloc[8] + income_stmt.iloc[9] # Total Sales

income_stmt.iloc[13] = v1.opex # Solar opex
income_stmt.iloc[14] = v2.opex # BESS opex
income_stmt.iloc[15] = income_stmt.iloc[13] + income_stmt.iloc[14] # Total opex

income_stmt.iloc[18] = income_stmt.iloc[8] - income_stmt.iloc[13] # Solar EBIT
income_stmt.iloc[19] = income_stmt.iloc[9] - income_stmt.iloc[14] # BESS EBIT
income_stmt.iloc[20] = tax_rate # Federal Tax Rate
income_stmt.iloc[21] = np.asarray(income_stmt.iloc[20]) * income_stmt.iloc[18] # Solar: Tax
income_stmt.iloc[22] = income_stmt.iloc[20] * income_stmt.iloc[19] # Bess: Tax

income_stmt.iloc[25] = income_stmt.iloc[18] - income_stmt.iloc[21] # Solar NOPAT
income_stmt.iloc[26] = income_stmt.iloc[19] - income_stmt.iloc[22] # BESS NOPAT
income_stmt.iloc[27] = income_stmt.iloc[25] + income_stmt.iloc[26] # Total NOPAT

income_stmt.iloc[30] = income_stmt.iloc[25] + income_stmt.iloc[3] # Solar - Net Cash Flow
income_stmt.iloc[31] = income_stmt.iloc[26] + income_stmt.iloc[4] # BESS - Net Cash Flow
income_stmt.iloc[32] = income_stmt.iloc[30] + income_stmt.iloc[31] # total Net Cash Flow

IRR = round(numpy_financial.irr(income_stmt.iloc[32]),4); # Calc IRR; numpy.irr is deprecated and will be removed from NumPy 1.20. Used numpy_financial.irr instead
income_stmt.iloc[34][0] = "{:.2%}".format(IRR) # Represent IRR in percentage

# Starting row is row from which income statement will be printed. It started 4 rows after the last of input rows
starting_row = ending_row_num + 4


book = openpyxl.load_workbook(myfile)
writer = pd.ExcelWriter(myfile, engine='openpyxl')  # Important to set this engine otherwise you'll encouner an error
writer.book = book  
# ExcelWriter uses writer.sheets to access the sheet. 
# If it is left empty, it will not know that sheet myworksheet is already there and will create a new sheet.
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
income_stmt.to_excel(writer,sheet_name=myworksheet,header=False,startrow=starting_row,startcol=1)
writer.save()

