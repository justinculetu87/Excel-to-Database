import pandas as pd
import openpyxl
import numpy as np

#load the master csv each time
df_main = pd.read_csv('P:\Housing Development Share\Development Cost Initiative\CDLAC_TCAC Competition Analysis\Application Analysis\Output\Master File\CDLAC_analysis_master.csv')

#add new columns if needed (uncomment)
#df_main.insert[# column, 'new name, ''] 

#add averages for later
df_main = df_main.iloc[:-1]

def dollar_convert(value): #convert to integers to do calculations in cell transfer 
        if value is not None and isinstance (value, str) and '$' in value:
            return float(value.replace('$', '').replace (',', ''))
        return value or 0

#add new application info
add_application = input('Load in new application? (yes/no): ')

#if yes, paste the file path of the application in the terminal
if add_application.lower() == 'yes':
    while True:
        file_path = input('Enter the path to the application file: ')
        work_book = openpyxl.load_workbook(file_path, data_only=True)

        new_obs = [] #make a new obervation to add to the master file
        
    #define worksheets
        ws1 = work_book['Application'] 
        ws2 = work_book['Sources and Uses Budget']
        ws3 = work_book['Basis & Credits']
        ws4 = work_book['Sources and Basis Breakdown']
        ws5 = work_book['Tie Breaker']

        new_row = { #call the cell values
            'Project Name': ws1['H17'].value if ws1['H17'].value is not None else '', 
            'Project Type': ws1['D214']. value if ws1['D214'].value is not None else '',
            'Geographic':  ws1['D223'].value if ws1['D223'].value is not None else '',
            'Total Units': ws1['AG438'].value if ws1['AG438'].value is not None else '',
            'Acreage': ws1['P419'].value if ws1['P419'].value is not None else '',
            'Density (DU/AC)': ws1['AF419'].value if ws1['AF419'].value is not None else '',
            'Number of Stories': ws1['AD412'].value if ws1['AD412'].value is not None else '', 
            'Net Rentable': f"${dollar_convert(ws1['Y788'].value)}" if ws1['Y788'].value not in (None, 0) else '',
            #'Parking Structure': ,
            'Total Gross Sq Ft': ws1['AG451'].value if ws1['AG451'].value is not None else '',
            'Land Cost – Total ($)': f"${dollar_convert(ws2['B12'].value)}",
            'Land Cost / Unit ($)': f"${dollar_convert(ws2['B12'].value) / dollar_convert(ws1['AG438'].value)}" if ws1['AG438'].value not in (None, 0) else '', 
            'Land Cost / Acre ($)': f"${dollar_convert(ws2['B12'].value) / dollar_convert(ws1['P419'].value)}" if ws1['P419'].value not in (None, 0) else '', 
            'Parking Spaces': (ws1['M499'].value or 0)+(ws1['AH499'].value or 0), 
            #'Parking Ratio': , 
            #'Parking Type': , 
            #'1BR Units (%)': ,
            #'2BR Units (%)': , 
            #'3BR Units (%)': , 
            #'4BR Units (%)': , 
            'AVG AMI': f"{ws1['AH738'].value * 100:.1f}%", 
            'Hard Costs incl. Contingency ($)': f"${(dollar_convert(ws2['B38'].value) - dollar_convert(ws2['B35'].value)) + dollar_convert(ws2['B78'].value)}",  ### 
            #'Hard Cost / Unit ($)': , 
            #'Hard Cost Net Rentable': , 
            #'Hard Cost Gross SF': , 
            'Prevailing Wage': ws1['AG977'].value if ws1['AG977'].value is not None else '', ###
            'OPX / Unit ($)': f"${dollar_convert(ws1['AC872'].value)}", 
            #'Soft Funding ($)': , 
            #'Soft Funding / Unit ($)': , 
            'Construction Interest': f"${dollar_convert(ws2['B45'].value)}" if ws2['B45'].value is not None else '', 
            #'Safehold Proceeds ($)': labeled as ground lease or safehold, III. Project Financing
            #'B- Bond ': subordinate bond (behind the perm loan 3rd lean positon) , 
            #'GP Note': , 
            #'Deferred Developer Fee': , AO631
            'Total Development Cost ($)': f"${dollar_convert(ws2['B104'].value)}", 
            'TDC / Unit ': f"${dollar_convert(ws2['B104'].value) / dollar_convert(ws1['AG438'].value)}",
            #'State Tax Credits': ,
            'Perm Lender': ws1['C629'].value if ws1['C629'].value is not None else '',  ###
            'Perm Loan Amount': ws1['AO641'].value if ws1['AO641'].value is not None else '' ,  ###
            'Rate': ws1['W629'].value if ws1['W629'].value is not None else ''  , ###
            #'Equity Investor': , 
            'Architect': ws1['AA289'].value if ws1['AA289'].value is not None else ''  , 
            'General Contractor': ws1['AA297'].value if ws1['AA297'].value is not None else '' , 
            'Resource Area': ws1['R199'].value if ws1['R199'].value is not None else '' , 
            #'CUAC Analyst': , 
            'Tie Breaker': f"{ws5['R75'].value * 100 :.1f}%"
            #total dev fee
            #cash fee = total dev - deffered dev 
            }
        new_obs.append(new_row)

        df_main = pd.concat([df_main, pd.DataFrame(new_obs)], ignore_index=True)
        print(f'{len(new_row)} row added successfully.')
    
        #adding new row after running this once
        another = input('Add another application (yes/no): ')
        if another.lower() != 'yes':
            print('Proceeding to calculations')
            break 

        
#########################################################################################################

#get averages
average = {}
for col in df_main.columns:
    try:
        strings = df_main[col].astype(str) #cast all to strings

        #determine formats
        has_dollar = strings.str.contains(r'\$', regex=True).any()
        has_percent = strings.str.contains(r'\%', regex=True).any()

        cleaned = pd.to_numeric(
            strings.str.replace(r'[$,%]', '', regex=True),
            errors='coerce'
        )
        if cleaned.notna().any():
            mean_val = cleaned.mean()
            if has_dollar:
                average[col] = f'${mean_val:,.2f}'
            elif has_percent:
                average[col] = f'%{mean_val:,.2f}'
            else: 
                average[col] = round(mean_val,2)
        else:
            average[col] = ''
    except Exception:
        average[col] = np.nan #empty cell if not applicable

df_main = pd.concat([df_main, pd.DataFrame([average])], ignore_index = True)
df_main.iloc[-1, 0] = 'Averages'


df_main.to_csv('P:\Housing Development Share\Development Cost Initiative\CDLAC_TCAC Competition Analysis\Application Analysis\Output\Master File\CDLAC_analysis_master.csv', index=False, encoding='utf-8-sig')  
print('File saved successfully')

    






