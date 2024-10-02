#this code calculates for every outlook over a course of a year how many ratings changed and in which direction 

import pandas as pd
import numpy as np

up = {}
down = {}
no_change = {}
stoped = {}
wr = {}
pdf = {} #dictinary for failiours

def merge_df(df1, df2):
    mat = []
    for i in range(len(df1)):
        mat.zappend(df1.iloc[i])
        mat.append(df2.iloc[i])
    mergred_df = pd.DataFrame(mat, columns = df1.columns)
    return mergred_df  

def get_exlusion_status(status):
    # אופק דירוג
    overview = ['חיובי', 'שלילי', 'יציב', 'WR', 'PD']
    # בחינת דירוג 
    outlook = ['בחינת דירוג עם השלכות שליליות', 'בחינת דירוג עם השלכות חיובויות', 'בחינת דירוג ללא כיוון וודאי']
    if status in overview:
        return  [s for s in overview if s != status]
    return [s for s in outlook if s != status]

def get_change_direction(original_rating, new_rating):
    rating_dic = {
        'Aaa.il' : 1, 'Aa1.il': 2, 'Aa2.il': 3, 'Aa3.il': 4, 'A1.il': 5, 'A2.il': 6, 'A3.il': 7,
        'Baa1.il': 8, 'Baa2.il': 9, 'Baa3.il': 10, 'Ba1.il': 11, 'Ba2.il': 12, 'Ba3.il': 13, 'B1.il': 1, 'B2.il': 15,
        'B3.il': 16, 'Caa1.il': 17, 'Caa2.il':18, 'Caa3.il': 19, 'Ca.il': 20, 'C.il': 21

    }
    os = rating_dic[original_rating]
    ns = rating_dic[new_rating]
    #reutrns positive if review improve, zero if not changed, negative if decreased
    return os - ns

def get_num_months(j, current_rating, rating):
    global up, down, no_change, stoped, wr
    months = 0
    new_rating = ''
    while (j < len(rating) and ((rating[j] == current_rating) or rating[j] == 0)):
        new_rating = rating[j]
        months += 1
        j += 1
    if(j < len(rating)):
        new_rating = rating[j]
    else:
        new_rating = 0

    if (current_rating == new_rating) or (new_rating == 0):
        no_change[months] = no_change.get(months, 0) + 1
    elif new_rating == 'WR':
        wr[months] = wr.get(months, 0) + 1
    elif new_rating == 'PD' or new_rating == 'FP':
        pdf[months] = pdf.get(months, 0) + 1
    elif get_change_direction(current_rating, new_rating) == 1:
        up[months] = up.get(months, 0) + 1
    else:
        down[months] = down.get(months, 0) + 1
    return j

def get_rating_duration(rating, review, status, not_status):
    i = 32
    while i < len(rating):
        if review[i] != status:
            i += 1
        else:
            current_rating = rating[i]
            i = get_num_months(i, current_rating, rating)
    return

def get_ratings_count_stauts(ratings, reviews,status):
    not_status = get_exlusion_status(status)
    for i in range(0, len(ratings)):
        get_rating_duration(ratings.iloc[i], reviews.iloc[i], status, not_status)
    return


def build_dataframes(status):
    global up, down, no_change, stoped, wr, pdf
    col_dict = {
        'up': 'מספר דירוגים שדירוגם עלה',
        'down': 'מספר דירוגים שדירוגם ירד', 
        'no_change': 'מספר דירוגים שדירוגם לא השתנה', 
        'stopped': 'מספר דירוגים שהגיעו לחדלות פרעון',
        'wr': 'מספר דירוגים שהופסקו',
        'pdf': 'מספר כשלים'
    }
    columns = ['בחינת דירוג עם ' + status] + list(range(1,13)) + ['else']
    df = pd.DataFrame(columns = columns)
    ups = [col_dict['up']] + [0] * 13
    downs = [col_dict['down']] + [0] * 13
    not_changed = [col_dict['no_change']] + [0] * 13
    wrs = [col_dict['wr']] + [0] * 13
    pds = [col_dict['pdf']] + [0] * 13
    
    for i in range(1, 13):
        ups[i] = up.get(i, 0)
    for i in range(1, 13):
        downs[i] = down.get(i, 0)
    for i in range(1, 13):
        not_changed[i] = no_change.get(i, 0)
    for i in range(1, 13):
        wrs[i] = wr.get(i, 0)
    for i in range(1, 13):
        pds[i] = pdf.get(i, 0)
    ups[13] = sum(value for key, value in up.items() if key < 1 or key > 12)
    downs[13] = sum(value for key, value in down.items() if key < 1 or key > 12)
    not_changed[13] = sum(value for key, value in no_change.items() if key < 1 or key > 12)
    wrs[13] = sum(value for key, value in wr.items() if key < 1 or key > 12)
        
    df.loc[0] = ups
    df.loc[1] = downs
    df.loc[2] = not_changed
    df.loc[3] = wrs
    df.loc[4] = pds
    return df


def clean_dicts():
    global up, down, no_change, stoped, wr, pdf 
    up = {}
    down = {}
    no_change = {}
    stoped = {}
    wr = {}
    pdf = {}
    return

def get_review_change(reviews, ratings, department):
    statuses = ['חיובי', 'שלילי', 'יציב', 'בחינת דירוג עם השלכות שליליות', 'בחינת דירוג עם השלכות חיובויות', 'בחינת דירוג ללא כיוון וודאי','PD']
    for status in statuses: 
        get_ratings_count_stauts(ratings, reviews, status)
        df = build_dataframes(status)
        with pd.ExcelWriter(f'research.xlsx', mode='a') as writer:  
            df.to_excel(writer, sheet_name=status, index=False)
        clean_dicts()
    return