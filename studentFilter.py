#import pandas lb as pd
import pandas as pd

df = pd.read_excel("students.xlsx")

#filter by grade and year
filtered_df = df[(df['Grade']=='4') & (df['Last Attendance Date'].str.contains('2023'))
                 | (df['Grade']=='3') & (df['Last Attendance Date'].str.contains('2023'))
                 | (df['Grade']=='2') & (df['Last Attendance Date'].str.contains('2023'))
                 |(df['Grade']=='1') & (df['Last Attendance Date'].str.contains('2023'))
                 |(df['Grade']=='3') & (df['Last Attendance Date'].str.contains('2022'))
                 |(df['Grade']=='2') & (df['Last Attendance Date'].str.contains('2022'))
                 |(df['Grade']=='1') & (df['Last Attendance Date'].str.contains('2022'))
                 |(df['Grade']=='K') & (df['Last Attendance Date'].str.contains('2022'))
                 |(df['Grade']=='2') & (df['Last Attendance Date'].str.contains('2021'))
                 |(df['Grade']=='1') & (df['Last Attendance Date'].str.contains('2021'))
                 |(df['Grade']=='K') & (df['Last Attendance Date'].str.contains('2021'))
                 |(df['Grade']=='1') & (df['Last Attendance Date'].str.contains('2020'))
                 |(df['Grade']=='K') & (df['Last Attendance Date'].str.contains('2020'))
                 |(df['Grade']=='K') & (df['Last Attendance Date'].str.contains('2019'))]
filtered_df.to_excel('file.xlsx', sheet_name='Filtered Students')

#create student dataframe and parent dataframe
sdf = pd.read_excel('file.xlsx')
pdf = pd.read_excel('parents.xlsx')

#create a series based off of guardian last names in the filtered student dataframe
ser = sdf.iloc[:,3]
ser = ser.str.split(',').str[0]
print(ser)

#create a new parent dataframe by matching last names between parent and filtered student dataframes
npdf = pdf[pdf['Last Name'].isin(ser)]
print(npdf)

#copy first name, last name, and email addresses to FINAL dataframe
fdf = pd.DataFrame()
fdf['First Name'] = npdf['First Name']
fdf['Last Name'] = npdf['Last Name']
fdf['Email Address'] = npdf['Email Address']

fdf.to_excel('master.xlsx', sheet_name='Names and Emails')





