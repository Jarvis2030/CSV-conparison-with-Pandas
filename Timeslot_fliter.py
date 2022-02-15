"""
Name: google sheet timeslot fliter
Author: Lin Yu-Chun, June
'Just for personal usage, still got some logical errors in this code, feel free to contact me
for any comment or discussion'
"""
import pandas as pd
Mon = pd.DataFrame()
Tue1 = pd.DataFrame()
Tue2 = pd.DataFrame()
Tue3 = pd.DataFrame()
Wed1 = pd.DataFrame()
Wed2 = pd.DataFrame()
Thu = pd.DataFrame() #create several dataframes for different timeslot


def OutputExcel():

    Result_PATH = r'C:\Users\user\OneDrive\桌面\hkust\UROP material\Total.xlsx'

    writer = pd.ExcelWriter(Result_PATH, engine='xlsxwriter')

    # SHEET1
    df_S1 = pd.DataFrame.from_dict(total)
    df_S1.to_excel(writer, sheet_name='SHEET1', index=False) 

    writer.save()
    print('成功產出'+Result_PATH)


reg = pd.read_csv(
    r'C:\Users\user\OneDrive\桌面\hkust\UROP material\Registration form_LOOP_August 4, 2021_23.24.csv')

Newreg = reg["Q5"].apply(lambda x: pd.Series(x.split(','))) 
# google sheet collect the MCQ data by putting them in one block and split with ','
                            

for i in range(87):
    for j in range(14): #Organize all the data to their timeslot dataframe
        if Newreg.iloc[i, j] == "Monday August 9":
            Mon = Mon.append(reg.iloc[i, 17:26], ignore_index=True)

        if Newreg.loc[i, j] == "Tuesday August 10":
            if Newreg.loc[i, j+1] == " 9am - 12nn":
                Tue1 = Tue1.append(reg.iloc[i, 17:26], ignore_index=True)
            elif Newreg.loc[i, j+1] == " 12nn - 3pm":
                Tue2 = Tue2.append(reg.iloc[i, 17:26], ignore_index=True)
            elif Newreg.loc[i, j+1] == " 3pm - 6pm":
                Tue3 = Tue3.append(reg.iloc[i, 17:26], ignore_index=True)

        if Newreg.loc[i, j] == "Wednesday August 11":
            if Newreg.loc[i, j+1] == " 10am - 1pm":
                Wed1 = Wed1.append(reg.iloc[i, 17:26], ignore_index=True)

            elif Newreg.loc[i, j+1] == " 1pm - 4pm":
                Wed2 = Wed2.append(reg.iloc[i, 17:26], ignore_index=True)

        if Newreg.loc[i, j] == "Thursday August 12":
            Thu = Thu.append(reg.iloc[i, 17:26], ignore_index=True)

Mon = Mon.reindex(Mon.Q5.str.len().sort_values().index) #sort all the dataframe and combined them into a sheet
Tue1 = Tue1.reindex(Tue1.Q5.str.len().sort_values().index)
Tue2 = Tue2.reindex(Tue2.Q5.str.len().sort_values().index)
Tue3 = Tue3.reindex(Tue3.Q5.str.len().sort_values().index)
Wed1 = Wed1.reindex(Wed1.Q5.str.len().sort_values().index)
Wed2 = Wed2.reindex(Wed2.Q5.str.len().sort_values().index)
Thu = Thu.reindex(Thu.Q5.str.len().sort_values().index)

total = pd.concat([Mon, Tue1, Tue2, Tue3, Wed1, Wed2, Thu], keys=[
                  'Monday', 'Tuesday morning', 'Tuesday noon', 'Tuesday afternoon', 'Wednesday noon', 'Wednesday afternoon', 'Thursday'],
                  names=['Session', 'Row ID'])

total = total.drop(total[(total['Q4'] == "Yes") | (total['Q9'] == "No")].index) #fliter those who disagree with the terms and conditions
print(total)
OutputExcel()
