# This code calculate the transition matrix on an end of month basis.

import pandas as pd
import numpy as np
import os

statuses = ['חיובי', 'שלילי', 'יציב','בחינת דירוג עם השלכות שליליות', 'בחינת דירוג עם השלכות חיובויות', 'בחינת דירוג ללא כיוון וודאי']
months = [1, 2, 3, 6, 9, 12]

def review_normalization(df):
    full_df = pd.DataFrame(columns=df.columns)
    for j in range(0, len(df)):
        i = 33
        row = df.iloc[j].tolist()
        new_row = row[:33]
        while i < len(row):
            if (row[i] == row[i-1]) or (row[i] == 0):
                row[i] = row[i-1]
            i += 1
        full_df.loc[len(full_df)] = row
        
    return full_df

def merged_for_inspection(df):
    mat = []
    norm_df = review_normalization(df)
    for i in range(len(df)):
        mat.append(df.iloc[i])
        mat.append(norm_df.iloc[i])
    merged_df = pd.DataFrame(mat, columns=df.columns)
    with pd.ExcelWriter('merge_test.xlsx', mode='w') as writer:  
        merged_df.to_excel(writer, sheet_name="1", index=False)
    return merged_df

# indicator is the variable of the status current status that is being cheked in the matrix, the status variable refers to the transtions itslef
def row_status_counter(indicator, status, row, counter, months): # count for each row
    for i in range(32, len(row)-months):
        if row[i] == indicator:
            if row[i+months] == status:
                counter += 1
    return counter

def get_transition_matrix(df, months):
    cols = ['CR'] + statuses + ['sum']
    df_count = pd.DataFrame(columns = cols)
    df_percentage = pd.DataFrame(columns = cols)
    for indicator in statuses:
        row = [0] * 6
        for s in range(0, len(statuses)):
            for i in range(0, len(df)):
               row[s] =  row_status_counter(indicator, statuses[s], df.iloc[i], row[s], months)
        if(sum(row) > 0):
            prow = [f'{(num / sum(row)) * 100:.2f}%' for num in row]
        else:
            prow = row[0:]
        df_count.loc[len(df_count)] = [indicator] + row + [sum(row)]
        df_percentage.loc[len(df_percentage)] = [indicator] + prow + [sum(row)]
    return(df_count, df_percentage)

def get_time_period_matrix(df, department):
    norm_df = review_normalization(df)
    for period in months:
        count_df, percentage_df = get_transition_matrix(norm_df, period)
        file_path = 'transition_matrix.xlsx'
        sheet_count_name = f"מספר שינוי דירוג {period} חודשים"
        sheet_percentage_name = f"אחוז שינוי דירוג {period} חודשים"
        
        if os.path.exists(file_path):
            with pd.ExcelWriter(file_path, mode='a') as writer:  
                count_df.to_excel(writer, sheet_name=sheet_count_name, index=False)
        else:
            with pd.ExcelWriter(file_path, mode='w') as writer:  
                count_df.to_excel(writer, sheet_name=sheet_count_name, index=False)
        
        with pd.ExcelWriter(file_path, mode='a') as writer:  
            percentage_df.to_excel(writer, sheet_name=sheet_percentage_name, index=False) 
    return
