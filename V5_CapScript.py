import pandas as pd
from datetime import datetime, timedelta

def load_data():
    xls = pd.ExcelFile(r"C:\Users\AliS\Desktop\test1\test4.xlsx")
    df = pd.read_excel(xls, 'Sheet1')
    df2 = pd.read_csv(r"C:\Users\AliS\Desktop\test1\Hour_map_test.csv")

    return df, df2

def Solution(df, df2, Solution):
    df = df.loc[df['Study Type'] == Solution]
    df2 = df2.loc[df2['Study Type'] == Solution]
    
    return df, df2

def preprocess_data(df):
    
    df = df[['Campaign', 'Study Type', 'FW Start Date', 'FW End Date', 'Presentation Date', 'Project Lead', 'Exec/CM']]
    dates = ['FW Start Date', 'FW End Date', 'Presentation Date']
    
    # Assuming 'FW End Date' and 'Presentation Date' are datetime columns
    df['Presentation Date'] = pd.to_datetime(df['Presentation Date'])
    df['FW End Date'] = pd.to_datetime(df['FW End Date'])
    
    # Calculate new date and assign it to empty rows in 'Presentation Date'
    mask = df['Presentation Date'].isnull()
    df.loc[mask, 'Presentation Date'] = df.loc[mask, 'FW End Date'] + timedelta(weeks=4)
    
    for i in dates:
        df[i] = pd.to_datetime(df[i], errors='coerce')
        df = df.dropna(subset=i, axis=0)
        
    # Split ppl into a seperate column and count no. of ppl on each project
    df['Exec1'] = None
    df['Exec2'] = None
    try:
        df[['Exec1', 'Exec2']] = df['Exec/CM'].str.split(r"/", expand=True)
    except ValueError:
        df['Exec1'] = df['Exec/CM']
    df = df.drop('Exec/CM', axis=1)
    df['N'] = df[['Project Lead', 'Exec1', 'Exec2']].count(axis=1)

    return df

def calculate_stage(row, i):
    if i > row['Presentation Date']:
        return 'Complete'
    elif i < row['FW Start Date']:
        return 'PreFW'
    elif i > row['FW End Date']:
        return 'PostFW'
    else:
        return 'Infield'

def calculate_hours(stage, df2):
    df2["Hours_1"] = 0.2*df2.Hours # Calculates "Maintanence Hours"
    df2["Hours_2"] = 0.8*df2.Hours # Calculates "Last week Hours"
    
    # Here we match the Stage of a project to their corresponding hours
    
    if stage.endswith('_x'):
        stage = stage[:-2]  # Remove the suffix "_x"
        hours = df2.loc[df2['Stage'] == stage, 'Hours_2'].values[0]
    elif stage == 'Complete':
        hours = 0
    else:
        hours = df2.loc[df2['Stage'] == stage, 'Hours_1'].values[0]
    return hours

def generate_summary_table(df, df2):
    
    today = datetime.today().date() # Starting point for current week (today)
    W2_Date = today + timedelta(weeks=1) # Week 1
    W3_Date = today + timedelta(weeks=2) # Week 2
    W4_Date = today + timedelta(weeks=3) # Week 3 

    SP = [today, W2_Date, W3_Date, W4_Date]
    
    # From each starting point we determine what stage of a project we are in (PreFW, Infield or PostFW)
    count = 0
    for i in SP:
        col_name = f'Stage_{count}'
        df[col_name] = df.apply(lambda row: calculate_stage(row, i), axis=1)

    # Now we look through the stages for each week. If there is a change in stage between weeks we know that it is the last week of that stage. We denote this by adding a "_x" at the end of the value.
        if count > 0:
            prev_col_name = f'Stage_{count-1}'
            df.loc[df[col_name] != df[prev_col_name], prev_col_name] += "_x"

        count += 1
        
    stages = ['Stage_0', 'Stage_1', 'Stage_2', 'Stage_3']
    stage_hours = [stage + '_Hours' for stage in stages]
    
    df = df[~(df[stages] == 'Complete').all(axis=1)]
    
    for col in stages:
        df[col + '_Hours'] = df[col].apply(calculate_hours, args=(df2,))
        df[col + '_Hours'] /= df['N']
    
    summary_table = pd.DataFrame(columns=['Person'] + stage_hours)
    persons = set(df['Project Lead'].tolist() + df['Exec1'].tolist() + df['Exec2'].tolist())
    
    for person in persons:
        person_rows = df[(df['Project Lead'] == person) | (df['Exec1'] == person) | (df['Exec2'] == person)]
        person_hours = person_rows[stage_hours].sum()
        summary_table = summary_table.append({'Person': person, 'Stage_0_Hours': person_hours['Stage_0_Hours'],
                                              'Stage_1_Hours': person_hours['Stage_1_Hours'],
                                              'Stage_2_Hours': person_hours['Stage_2_Hours'],
                                              'Stage_3_Hours': person_hours['Stage_3_Hours']}, ignore_index=True)

    summary_table = summary_table.dropna()

    return summary_table, df

def siuuu(df, df2, sol):
    df, df2 = load_data()
    df = df.replace('\s', '', regex=True)
    df2 = df2.replace('\s', '', regex=True)
    df['Exec1'] = None
    df['Exec2'] = None
    df, df2 = Solution(df, df2, sol)
    df = preprocess_data(df)
    sum_table, df = generate_summary_table(df, df2)
    
    return sum_table, df

combined_df = pd.DataFrame()
combined_df1 = pd.DataFrame()

for i in ["XM", "ContextLab", "BLI"]:
    df, df2 = load_data()
    table, project_stages = siuuu(df,df2,i)
    combined_df = pd.concat([combined_df, table], ignore_index=True)
    combined_df1 = pd.concat([combined_df1, project_stages], ignore_index=True)

s = combined_df.groupby(by="Person").sum()
s /= 30
s *= 5
    

def write_():
    # Create an ExcelWriter object
    writer = pd.ExcelWriter('combined_data.xlsx')
    
    # Write each DataFrame to a separate sheet
    s.to_excel(writer, sheet_name='summary', index=True)
    combined_df1.to_excel(writer, sheet_name='detailed', index=False)
    
    # Save and close the ExcelWriter object
    writer.save()
    writer.close()

write_()
