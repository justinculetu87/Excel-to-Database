import pandas as pd
import numpy as np

df = pd.read_csv ('P:\Housing Development Share\Development Cost Initiative\CDLAC_TCAC Competition Analysis\Application Analysis\Raw Files\Revised Linc_CDLAC Analysis Workbook_Personal.csv')
print(df.head())

df_reshape = df.T.reset_index() #transpose the dataframe
df_reshape.columns = df_reshape.iloc[0] #first row as header
df_reshape = df_reshape.drop(df_reshape.index[0]) #drop first row

#reformat some observations
df_reshape['Net Rentable'] = df_reshape['Net Rentable'].apply(lambda x: x if str(x).startswith('$') else f'${x}')
df_reshape['Total Gross Sq Ft'] = df_reshape['Total Gross Sq Ft'].str.replace('$', '', regex=False).str.replace(',', '', regex=False)
df_reshape['Total Gross Sq Ft'] = pd.to_numeric(df_reshape['Total Gross Sq Ft'])

df_reshape = df_reshape.rename(columns= {'Category' : 'Project Name' })   #rename the project row

df_reshape = df_reshape.drop(df_reshape.index[0]) #drop the current avg row

df_reshape.to_csv('P:\Housing Development Share\Development Cost Initiative\CDLAC_TCAC Competition Analysis\Application Analysis\Output\Interim Data\CDLAC_analysis.csv', index=False, encoding = 'utf-8-sig') #Save as interim data to make sure new additions save














