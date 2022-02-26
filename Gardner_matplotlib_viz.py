# -*- coding: utf-8 -*-
"""
Created on Sat Feb 19 10:14:26 2022

@author: Charlie Gardner
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.ticker import FuncFormatter
import os

os.chdir('G:/Python/.jupyter/MSDS670/')

#%%

#Read multiple sheets from the same file
xlsx = pd.ExcelFile('wsf_2021_scholarships.xlsx' )
df_apps = pd.read_excel(xlsx, 'Applicants')
df_awards = pd.read_excel(xlsx, 'Awards')
df_appdate = pd.read_excel(xlsx, 'AppDate')
df_race = pd.read_excel(xlsx, 'Race')


#%%

# Horizontal bar plot counting miscellaneous characteristics of Applicants to WSF Scholarships

# Query the Miscellaneous column to pull out and count different characterisitcs of applicants

columns_apps = list(df_apps.columns)
crosby_count = df_apps['MISCELLANEOUS'].str.contains('Crosby', na=False).sum()
arts_count = df_apps['MISCELLANEOUS'].str.contains('arts', na=False).sum()
audobon_count = df_apps['MISCELLANEOUS'].str.contains('wildlife', na=False).sum()
firstgen_count = df_apps['MISCELLANEOUS'].str.contains('generation', na=False).sum()
disability_count = df_apps['MISCELLANEOUS'].str.contains('disability', na=False).sum()
housing_count = df_apps['MISCELLANEOUS'].str.contains('Housing Authority', na=False).sum()
veteran_count = df_apps['MISCELLANEOUS'].str.contains('Veteran', na=False).sum()
degree_count = df_apps['MISCELLANEOUS'].str.contains('degree', na=False).sum()
prison_count = df_apps['MISCELLANEOUS'].str.contains('prison', na=False).sum()

# Create a dataframe with the characteristic counts
df_misc = pd.DataFrame({'Characteristics': ['Crosby Scholar', 
                                       'Artist',
                                       'First Generation Student',
                                       'Already has a Degree',
                                       "Person with Disability",
                                       'Audobon Society Member',
                                       'Child of Veteran',
                                       'Public Housing Resident',
                                       'Formerly Incarcerated'], 
                        'Count': [crosby_count, 
                             arts_count, 
                             firstgen_count,
                             degree_count,
                             disability_count,
                             audobon_count,
                             veteran_count,
                             housing_count,
                             prison_count
                             ]})

df_misc = df_misc.sort_values(by=['Count'], ascending=True)


# function to add value labels
x = df_misc['Characteristics']
y = df_misc['Count']

# create horizontal bar chart
fig, ax = plt.subplots(figsize=(12,6))

ax.barh(df_misc['Characteristics'], df_misc['Count'])
ax.set_title('Count of Applicants by Characteristics', fontweight='bold', fontsize=20)
ax.set_xlabel('Applicant Count')
ax.set_ylabel('Characteristics') 

#Add labels for each bar
for i, v in enumerate(y):
    ax.text(v, i, str(v), color='black', fontweight='bold', fontsize=14, ha='left', va='center')



#%%

# Line plot comparing scholarship award recicipients and non-recipients by dates

# create line plot
fig, ax = plt.subplots(figsize=(16,8))

ax.plot(df_appdate['Not Offered Count'].values, linewidth=7, label='Non Award Recipient', color='skyblue')
ax.plot(df_appdate['Offered Count'].values, linewidth=7, label='Award Recipient', color='midnightblue')

# Formatting Axes
ax.set_title('Date of Application by Students (2021)', fontsize=24)
ax.set_xlabel('Application Date', fontsize=12)
ax.set_ylabel('Number of Applicants', fontsize=12, labelpad=1.5)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

#Label the ticks on x axis
ax.set_xticks(np.arange(0, 160, 14))
ax.set_xticklabels(['Jan 1', 'Jan 15', 'Jan 29', 'Feb 12', 'Feb 26',
                    'Mar 12', 'March 26', 'April 9', 'April 23',
                    'May 7', 'May 21', 'Jun 4'])

# Adding deadline for merit based awards.
plt.axvline(x=75, color='gold', linewidth =5, label= 'March 15: Deadline for Merit Based Awards')

# Add legend
ax.legend(fontsize=14)

 





#%%

# Line plot of scholarship rate by racial demogrpahics with bar chart count in background

x = df_race['Race']
y = df_race['Award Rate']

# Formatting axes
fig, ax1 = plt.subplots(figsize=(16,8))
ax1.set_ylabel('Award Rate')
ax1.set_xlabel('Race')
ax1.set_yticklabels(['0%', '20%', '40%', '60%', '80%', '100%'], fontsize=12)
ax1.set_title('Scholarship Award Rate by Racial Demographic', fontweight='bold', fontsize=24)


# Adding line plot with award rate
ax1.plot(x, df_race['Award Rate'], 'D-g', linewidth=3, color='gold', label = "Award Rate")
plt.ylim(0, 1)

# Create secondary axis
ax = ax1.twinx()

# Add transparent bar plots with the count of award recipients and non-recipients
ax.bar(x, df_race['Non-Recipient'], bottom=df_race['Aid Recipient'], label='Non Award Recipient', alpha=0.50, color='skyblue')
ax.bar(x, df_race['Aid Recipient'],  label='Award Recipient', color='midnightblue', alpha=0.50)

ax.set_ylabel('Number of Applicants')
ax.legend(fontsize=14)

# Label the line plot with rate
for i, v in enumerate(y):
    ax.text(i, v + 260, str(v*100)+'%', fontweight='bold', color = 'darkgoldenrod' , fontsize=20, ha = 'center')




