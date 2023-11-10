# -*- coding: utf-8 -*-
"""
Created on Wed Sep 22 11:04:12 2021

@author: TurnerJ
"""
# This is a python script to pull in data - nothing to be changed 
import pandas as pd
pd.set_option('display.max_colwidth', None)

import win32com.client


# This is creating a format to format the numerics of the report to specify if it is in the Thousands, millions etc. 
def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])


#This is formatting the colors of the pacing portion - green is we are pacing above 25% and gold if we are #1 but pacing below 25%
def color_code_ahead(value):
            if value > 25:
                return """<b><p style="color:green;">"""
            else:
                return """<b><p style="color:gold;">"""
            
  
# This is highlighting the Fox News specific lines in blue to stand out 
def color_code_fox_news(value):
         if value=="Fox News":
            return """<b><p style="color:blue;">"""
         else:
            return """</b><p style="color:black;">"""

#Total Interactions
# This is only pulling in the one sheet from the workbook to consolidate the report and simplify the code
#r'C:\Users\VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230701.xlsx
overview  = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Overview')


overview = overview.set_index('Rank')
Total_Interactions = overview.iloc[0:5].copy()



Total_Interactions = Total_Interactions[['Site', 'Total Interactions', 'Interactions per 1k Followers']]
Total_Interactions['Interactions per 1k Followers'] = round(Total_Interactions["Interactions per 1k Followers"]).astype(int)
Total_Interactions['Interactions per 1k Followers'] = Total_Interactions['Interactions per 1k Followers'].map('{:,d}'.format)



Total_Interactions['Total Interactions'].iloc[0] = human_format(Total_Interactions['Total Interactions'].iloc[0])
Total_Interactions['Total Interactions'].iloc[1] = human_format(Total_Interactions['Total Interactions'].iloc[1])
Total_Interactions['Total Interactions'].iloc[2] = human_format(Total_Interactions['Total Interactions'].iloc[2])
Total_Interactions['Total Interactions'].iloc[3] = human_format(Total_Interactions['Total Interactions'].iloc[3])
Total_Interactions['Total Interactions'].iloc[4] = human_format(Total_Interactions['Total Interactions'].iloc[4])

# Creating the first section of the email - intro and combined total interactions 
print('\n Total Interactions (Facebook/Instagram/Twitter combined):')
print("1. " + Total_Interactions['Site'].iloc[0] + " – "  + Total_Interactions['Total Interactions'].iloc[0] +
      ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[0] + " interactions per 1k followers/fans.")
print("2. " + Total_Interactions['Site'].iloc[1] + " – "  + Total_Interactions['Total Interactions'].iloc[1] +
      ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[1] + " interactions per 1k followers/fans.")
print("3. " + Total_Interactions['Site'].iloc[2] + " – "  + Total_Interactions['Total Interactions'].iloc[2] +
      ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[2] + " interactions per 1k followers/fans.")
print("4. " + Total_Interactions['Site'].iloc[3] + " – "  + Total_Interactions['Total Interactions'].iloc[3] +
      ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[3] + " interactions per 1k followers/fans.")
print("5. " + Total_Interactions['Site'].iloc[4] + " – "  + Total_Interactions['Total Interactions'].iloc[4] +
      ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[4] + " interactions per 1k followers/fans.")

Total_Interactions_Pacing = overview.iloc[0:5].copy()
if Total_Interactions_Pacing.Site.iloc[0] == 'Fox News':
        fox_news = Total_Interactions_Pacing['Total Interactions'].iloc[0]
        site_competitor = Total_Interactions_Pacing['Total Interactions'].iloc[1]
        pacing_value = ((fox_news/site_competitor)-1)
        pacing_value = (round(pacing_value,2) * 100).astype(int)
        pacing_statement = "Fox News is pacing " + str(pacing_value) + "% ahead of the next most interacted news publisher."
        pacing_statement_color = color_code_ahead(pacing_value) + pacing_statement + """</p></b>"""
            
else:
        site_competitor =  Total_Interactions_Pacing['Total Interactions'].iloc[0]
        fox_news = Total_Interactions_Pacing.loc[(Total_Interactions_Pacing['Site'] == 'Fox News')].iloc[0][1]
        pacing_value = ((site_competitor/fox_news)-1)
        pacing_value = (round(pacing_value,2) *100).astype(int)
        pacing_statement = "Fox News is pacing " + str(pacing_value) + "% behind of the most interacted news publisher."
        pacing_statement_color = """<b><p style="color:red;">""" + pacing_statement + """</p></b>"""

#Facebook Interactions

Facebook = overview[['Site.1', 'Facebook Interactions', 'Facebook Interacts per 1k Fans']].sort_values('Facebook Interactions', ascending=False)
Facebook.iloc[0:5]

Facebook['Facebook Interacts per 1k Fans'] = round(Facebook['Facebook Interacts per 1k Fans']).fillna(0).astype(int)
Facebook['Facebook Interacts per 1k Fans'] = Facebook['Facebook Interacts per 1k Fans'].map('{:,d}'.format)


Facebook['Facebook Interactions'].iloc[0] = human_format(Facebook['Facebook Interactions'].iloc[0])
Facebook['Facebook Interactions'].iloc[1] = human_format(Facebook['Facebook Interactions'].iloc[1])
Facebook['Facebook Interactions'].iloc[2] = human_format(Facebook['Facebook Interactions'].iloc[2])
Facebook['Facebook Interactions'].iloc[3] = human_format(Facebook['Facebook Interactions'].iloc[3])
Facebook['Facebook Interactions'].iloc[4] = human_format(Facebook['Facebook Interactions'].iloc[4])

print('\n Facebook:')
print("1. " + Facebook['Site.1'].iloc[0] + " – "  + Facebook['Facebook Interactions'].iloc[0] +
      ' total interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[0] + " interactions per 1k followers/fans.")
print("2. " + Facebook['Site.1'].iloc[1] + " – "  + Facebook['Facebook Interactions'].iloc[1] +
      ' total interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[1] + " interactions per 1k followers/fans.")
print("3. " + Facebook['Site.1'].iloc[2] + " – "  + Facebook['Facebook Interactions'].iloc[2] +
      ' total interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[2] + " interactions per 1k followers/fans.")
print("4. " + Facebook['Site.1'].iloc[3] + " – "  + Facebook['Facebook Interactions'].iloc[3] +
      ' total interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[3] + " interactions per 1k followers/fans.")
print("5. " + Facebook['Site.1'].iloc[4] + " – "  + Facebook['Facebook Interactions'].iloc[4] +
      ' total interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[4] + " interactions per 1k followers/fans.")


Facebook_Interactions_Pacing = overview.copy()
Facebook_Interactions_Pacing = Facebook_Interactions_Pacing[['Site.1','Facebook Interactions']]
Facebook_Interactions_Pacing = Facebook_Interactions_Pacing.sort_values(by=['Facebook Interactions'], ascending=False)
Facebook_Interactions_Pacing.reset_index(drop=True)
if Facebook_Interactions_Pacing['Site.1'].iloc[0] == 'Fox News':
        fb_fox_news = Facebook_Interactions_Pacing['Facebook Interactions'].iloc[0]
        fb_site_competitor = Facebook_Interactions_Pacing['Facebook Interactions'].iloc[1]
        fb_pacing_value = ((fb_fox_news/fb_site_competitor)-1)
        fb_pacing_value = (round(fb_pacing_value,2) * 100).astype(int)
        fb_pacing_statement = "Fox News is pacing " + str(fb_pacing_value) + "% ahead of the next most interacted news publisher."
        fb_pacing_statement_color = color_code_ahead(fb_pacing_value) + fb_pacing_statement + """</p></b>"""
else:
        fb_site_competitor =  Facebook_Interactions_Pacing['Facebook Interactions'].iloc[0]
        fb_fox_news = Facebook_Interactions_Pacing.loc[(Facebook_Interactions_Pacing['Site.1'] == 'Fox News')].iloc[0][1]
        fb_pacing_value = ((fb_site_competitor/fb_fox_news)-1)
        fb_pacing_value = (round(fb_pacing_value,2) *100).astype(int)
        fb_pacing_statement = "Fox News is pacing " + str(fb_pacing_value) + "% behind of the most interacted news publisher."
        fb_pacing_statement_color = """<b><p style="color:red;">""" + fb_pacing_statement + """</p></b>"""



#Instagram Interactions

Instagram = overview[['Site.2', 'Instagram Interactions', 'Instagram Interacts per 1k Fans']].sort_values('Instagram Interactions', ascending=False)
Instagram.iloc[0:5]

Instagram['Instagram Interacts per 1k Fans'] = round(Instagram['Instagram Interacts per 1k Fans']).fillna(0).astype(int)
Instagram['Instagram Interacts per 1k Fans'] = Instagram['Instagram Interacts per 1k Fans'].map('{:,d}'.format)


Instagram['Instagram Interactions'].iloc[0] = human_format(Instagram['Instagram Interactions'].iloc[0])
Instagram['Instagram Interactions'].iloc[1] = human_format(Instagram['Instagram Interactions'].iloc[1])
Instagram['Instagram Interactions'].iloc[2] = human_format(Instagram['Instagram Interactions'].iloc[2])
Instagram['Instagram Interactions'].iloc[3] = human_format(Instagram['Instagram Interactions'].iloc[3])
Instagram['Instagram Interactions'].iloc[4] = human_format(Instagram['Instagram Interactions'].iloc[4])

print('\n Instagram:')
print("1. " + Instagram['Site.2'].iloc[0] + " – "  + Instagram['Instagram Interactions'].iloc[0] +
      ' total interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[0] + " interactions per 1k followers/fans.")
print("2. " + Instagram['Site.2'].iloc[1] + " – "  + Instagram['Instagram Interactions'].iloc[1] +
      ' total interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[1] + " interactions per 1k followers/fans.")
print("3. " + Instagram['Site.2'].iloc[2] + " – "  + Instagram['Instagram Interactions'].iloc[2] +
      ' total interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[2] + " interactions per 1k followers/fans.")
print("4. " + Instagram['Site.2'].iloc[3] + " – "  + Instagram['Instagram Interactions'].iloc[3] +
      ' total interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[3] + " interactions per 1k followers/fans.")
print("5. " + Instagram['Site.2'].iloc[4] + " – "  + Instagram['Instagram Interactions'].iloc[4] +
      ' total interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[4] + " interactions per 1k followers/fans.")


Instagram_Interactions_Pacing = overview.copy()
Instagram_Interactions_Pacing = Instagram_Interactions_Pacing[['Site.2','Instagram Interactions']]
Instagram_Interactions_Pacing = Instagram_Interactions_Pacing.sort_values(by=['Instagram Interactions'], ascending=False)
Instagram_Interactions_Pacing.reset_index(drop=True)
if Instagram_Interactions_Pacing['Site.2'].iloc[0] == 'Fox News':
        ig_fox_news = Instagram_Interactions_Pacing['Instagram Interactions'].iloc[0]
        ig_site_competitor = Instagram_Interactions_Pacing['Instagram Interactions'].iloc[1]
        ig_pacing_value = ((ig_fox_news/ig_site_competitor)-1)
        ig_pacing_value = (round(ig_pacing_value,2) * 100).astype(int)
        ig_pacing_statement = "Fox News is pacing " + str(ig_pacing_value) + "% ahead of the next most interacted news publisher."
        ig_pacing_statement_color = color_code_ahead(ig_pacing_value) + ig_pacing_statement + """</p></b>"""
else:
        ig_site_competitor =  Instagram_Interactions_Pacing['Instagram Interactions'].iloc[0]
        ig_fox_news = Instagram_Interactions_Pacing.loc[(Instagram_Interactions_Pacing['Site.2'] == 'Fox News')].iloc[0][1]
        ig_pacing_value = ((ig_site_competitor/ig_fox_news)-1)
        ig_pacing_value = (round(ig_pacing_value,2) *100).astype(int)
        ig_pacing_statement = "Fox News is pacing " + str(ig_pacing_value) + "% behind of the most interacted news publisher."
        ig_pacing_statement_color = """<b><p style="color:red;">""" + ig_pacing_statement + """</p></b>"""

#Twitter Interactions

Twitter = overview[['Site.3', 'Twitter Interactions', 'Twitter Interacts per 1k Fans']].sort_values('Twitter Interactions', ascending=False)
Twitter.iloc[0:5]

Twitter['Twitter Interacts per 1k Fans'] = round(Twitter['Twitter Interacts per 1k Fans']).fillna(0).astype(int)
Twitter['Twitter Interacts per 1k Fans'] = Twitter['Twitter Interacts per 1k Fans'].map('{:,d}'.format)


Twitter['Twitter Interactions'].iloc[0] = human_format(Twitter['Twitter Interactions'].iloc[0])
Twitter['Twitter Interactions'].iloc[1] = human_format(Twitter['Twitter Interactions'].iloc[1])
Twitter['Twitter Interactions'].iloc[2] = human_format(Twitter['Twitter Interactions'].iloc[2])
Twitter['Twitter Interactions'].iloc[3] = human_format(Twitter['Twitter Interactions'].iloc[3])
Twitter['Twitter Interactions'].iloc[4] = human_format(Twitter['Twitter Interactions'].iloc[4])

print('\n Twitter:')
print("1. " + Twitter['Site.3'].iloc[0] + " – "  + Twitter['Twitter Interactions'].iloc[0] +
      ' total interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[0] + " interactions per 1k followers/fans.")
print("2. " + Twitter['Site.3'].iloc[1] + " – "  + Twitter['Twitter Interactions'].iloc[1] +
      ' total interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[1] + " interactions per 1k followers/fans.")
print("3. " + Twitter['Site.3'].iloc[2] + " – "  + Twitter['Twitter Interactions'].iloc[2] +
      ' total interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[2] + " interactions per 1k followers/fans.")
print("4. " + Twitter['Site.3'].iloc[3] + " – "  + Twitter['Twitter Interactions'].iloc[3] +
      ' total interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[3] + " interactions per 1k followers/fans.")
print("5. " + Twitter['Site.3'].iloc[4] + " – "  + Twitter['Twitter Interactions'].iloc[4] +
      ' total interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[4] + " interactions per 1k followers/fans.")


Twitter_Interactions_Pacing = overview.copy()
Twitter_Interactions_Pacing = Twitter_Interactions_Pacing[['Site.3','Twitter Interactions']]
Twitter_Interactions_Pacing = Twitter_Interactions_Pacing.sort_values(by=['Twitter Interactions'], ascending=False)
Twitter_Interactions_Pacing.reset_index(drop=True)
if Twitter_Interactions_Pacing['Site.3'].iloc[0] == 'Fox News':
        tw_fox_news = Twitter_Interactions_Pacing['Twitter Interactions'].iloc[0]
        tw_site_competitor = Twitter_Interactions_Pacing['Twitter Interactions'].iloc[1]
        tw_pacing_value = ((tw_fox_news/tw_site_competitor)-1)
        tw_pacing_value = (round(tw_pacing_value,2) * 100).astype(int)
        tw_pacing_statement = "Fox News is pacing " + str(tw_pacing_value) + "% ahead of the next most interacted news publisher."
        tw_pacing_statement_color = color_code_ahead(tw_pacing_value) + tw_pacing_statement + """</p></b>"""
else:
        tw_site_competitor =  Twitter_Interactions_Pacing['Twitter Interactions'].iloc[0]
        tw_fox_news = Twitter_Interactions_Pacing.loc[(Twitter_Interactions_Pacing['Site.3'] == 'Fox News')].iloc[0][1]
        tw_pacing_value = ((tw_site_competitor/tw_fox_news)-1)
        tw_pacing_value = (round(tw_pacing_value,2) *100).astype(int)
        tw_pacing_statement = "Fox News is pacing " + str(tw_pacing_value) + "% behind of the most interacted news publisher."
        tw_pacing_statement_color = """<b><p style="color:red;">""" + tw_pacing_statement + """</p></b>"""


#Uniques
uniques = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Uniques')

uniques[['Total', 'Facebook & FBIA', 'Instagram', 'Twitter']]


#Uniques Percent Change
uniques_PC = uniques.iloc[2:3].copy()
uniques_PC['Total'] = uniques_PC['Total'].round(4) * 100
uniques_PC['Facebook & FBIA'] = uniques_PC['Facebook & FBIA'].round(4) * 100
uniques_PC['Instagram'] = uniques_PC['Instagram'].round(4) * 100
uniques_PC['Twitter'] = uniques_PC['Twitter'].round(4) * 100

uniques_PC['Total_Trend'] = uniques_PC['Total'] > 0
uniques_PC['Facebook_Trend'] = uniques_PC['Facebook & FBIA'] > 0
uniques_PC['Instagram_Trend'] = uniques_PC['Instagram'] > 0
uniques_PC['Twitter_Trend'] = uniques_PC['Twitter'] > 0

uniques_PC.replace(True, 'up', inplace=True)
uniques_PC.replace(False, 'down', inplace=True)

uniques_PC['Total'] = abs(round(uniques_PC['Total']))
uniques_PC['Facebook & FBIA'] = abs(round(uniques_PC['Facebook & FBIA']))
uniques_PC['Instagram'] = abs(round(uniques_PC['Instagram']))
uniques_PC['Twitter'] = abs(round(uniques_PC['Twitter']))

uniques_PC['Total'] = uniques_PC['Total'].astype(int).astype(str) + "%"
uniques_PC['Facebook & FBIA'] = uniques_PC['Facebook & FBIA'].astype(int).astype(str) + "%"
uniques_PC['Instagram'] = uniques_PC['Instagram'].astype(int).astype(str) + "%"
uniques_PC['Twitter'] = uniques_PC['Twitter'].astype(int).astype(str) + "%"


#Page Views
pageviews = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Page Views')

pageviews[['Total', 'Facebook & FBIA', 'Instagram', 'Twitter']]


#Page Views Percent Change
pageviews_PC = pageviews.iloc[2:3].copy()
pageviews_PC['Total'] = pageviews_PC['Total'].round(4) * 100
pageviews_PC['Facebook & FBIA'] = pageviews_PC['Facebook & FBIA'].round(4) * 100
pageviews_PC['Instagram'] = pageviews_PC['Instagram'].round(4) * 100
pageviews_PC['Twitter'] = pageviews_PC['Twitter'].round(4) * 100

pageviews_PC['Total_Trend'] = pageviews_PC['Total'] > 0
pageviews_PC['Facebook_Trend'] = pageviews_PC['Facebook & FBIA'] > 0
pageviews_PC['Instagram_Trend'] = pageviews_PC['Instagram'] > 0
pageviews_PC['Twitter_Trend'] = pageviews_PC['Twitter'] > 0

pageviews_PC.replace(True, 'up', inplace=True)
pageviews_PC.replace(False, 'down', inplace=True)

pageviews_PC['Total'] = abs(round(pageviews_PC['Total']))
pageviews_PC['Facebook & FBIA'] = abs(round(pageviews_PC['Facebook & FBIA']))
pageviews_PC['Instagram'] = abs(round(pageviews_PC['Instagram']))
pageviews_PC['Twitter'] = abs(round(pageviews_PC['Twitter']))

pageviews_PC['Total'] = pageviews_PC['Total'].astype(int).astype(str) + "%"
pageviews_PC['Facebook & FBIA'] = pageviews_PC['Facebook & FBIA'].astype(int).astype(str) + "%"
pageviews_PC['Instagram'] = pageviews_PC['Instagram'].astype(int).astype(str) + "%"
pageviews_PC['Twitter'] = pageviews_PC['Twitter'].astype(int).astype(str) + "%"


#Minutes Spent


minutesspent = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Minutes Spent')

minutesspent[['Total', 'Facebook & FBIA', 'Instagram', 'Twitter']]

#Minutes Spent Percent Change
minutesspent_PC = minutesspent.iloc[2:3].copy()
minutesspent_PC['Total'] = minutesspent_PC['Total'].round(4) * 100
minutesspent_PC['Facebook & FBIA'] = minutesspent_PC['Facebook & FBIA'].round(4) * 100
minutesspent_PC['Instagram'] = minutesspent_PC['Instagram'].round(4) * 100
minutesspent_PC['Twitter'] = minutesspent_PC['Twitter'].round(4) * 100

minutesspent_PC['Total_Trend'] = minutesspent_PC['Total'] > 0
minutesspent_PC['Facebook_Trend'] = minutesspent_PC['Facebook & FBIA'] > 0
minutesspent_PC['Instagram_Trend'] = minutesspent_PC['Instagram'] > 0
minutesspent_PC['Twitter_Trend'] = minutesspent_PC['Twitter'] > 0

minutesspent_PC.replace(True, 'up', inplace=True)
minutesspent_PC.replace(False, 'down', inplace=True)

minutesspent_PC['Total'] = abs(round(minutesspent_PC['Total']))
minutesspent_PC['Facebook & FBIA'] = abs(round(minutesspent_PC['Facebook & FBIA']))
minutesspent_PC['Instagram'] = abs(round(minutesspent_PC['Instagram']))
minutesspent_PC['Twitter'] = abs(round(minutesspent_PC['Twitter']))

minutesspent_PC['Total'] = minutesspent_PC['Total'].astype(int).astype(str) + "%"
minutesspent_PC['Facebook & FBIA'] = minutesspent_PC['Facebook & FBIA'].astype(int).astype(str) + "%"
minutesspent_PC['Instagram'] = minutesspent_PC['Instagram'].astype(int).astype(str) + "%"
minutesspent_PC['Twitter'] = minutesspent_PC['Twitter'].astype(int).astype(str) + "%"

# Date function here will need to be reformatted to just "the month of" 
date = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Uniques')
date = date.iloc[0].values[0]

# This is from the Adobe Data Sheet in the weekly - a formula would need to be created in the monthly for this to populate correctly 
# We do not need a day for this header, just the month and year 
date_header = pd.read_excel(r'C:\Users\VayosKa\Desktop\Social Reporting\Social_Reporting_071723.xlsx', sheet_name='Adobe Data')
date_header = date_header['Email Header'].iloc[0:3].astype(int)
date_header = str(date_header[0]) + "." + str(date_header[2]) + "." + str(date_header[1])


#Final Results
print('\n Total:')
print(human_format(pageviews['Total'].iloc[0]) + " page views, " + pageviews_PC['Total_Trend'] + " " + pageviews_PC['Total'] + " from the same time last month")
print(human_format(uniques['Total'].iloc[0]) + " unique devices, " + uniques_PC['Total_Trend'] + " " + uniques_PC['Total'] + " from the same time last month")
print(human_format(minutesspent['Total'].iloc[0]) + " minutes spent, " + minutesspent_PC['Total_Trend'] + " " + minutesspent_PC['Total'] + " from the same time last month")

print('\n Facebook & FBIA:')
print(human_format(pageviews['Facebook & FBIA'].iloc[0]) + " page views, " + pageviews_PC['Facebook_Trend'] + " " + pageviews_PC['Facebook & FBIA'] + " from the same time last month")
print(human_format(uniques['Facebook & FBIA'].iloc[0]) + " unique devices, " + uniques_PC['Facebook_Trend'] + " " + uniques_PC['Facebook & FBIA'] + " from the same time last month")
print(human_format(minutesspent['Facebook & FBIA'].iloc[0]) + " minutes spent, " + minutesspent_PC['Facebook_Trend'] + " " + minutesspent_PC['Facebook & FBIA'] + " from the same time last month")

print('\n Instagram:')
print(human_format(pageviews['Instagram'].iloc[0]) + " page views, " + pageviews_PC['Instagram_Trend'] + " " + pageviews_PC['Instagram'] + " from the same time last month")
print(human_format(uniques['Instagram'].iloc[0]) + " unique devices, " + uniques_PC['Instagram_Trend'] + " " + uniques_PC['Instagram'] + " from the same time last month")
print(human_format(minutesspent['Instagram'].iloc[0]) + " minutes spent, " + minutesspent_PC['Instagram_Trend'] + " " + minutesspent_PC['Instagram'] + " from the same time last month")

print('\n Twitter:')
print(human_format(pageviews['Twitter'].iloc[0]) + " page views, " + pageviews_PC['Twitter_Trend'] + " " + pageviews_PC['Twitter'] + " from the same time last month")
print(human_format(uniques['Twitter'].iloc[0])+ " unique devices, " + uniques_PC['Twitter_Trend'] + " " + uniques_PC['Twitter'] + " from the same time last month")
print(human_format(minutesspent['Twitter'].iloc[0])+ " minutes spent, " + minutesspent_PC['Twitter_Trend'] + " " + minutesspent_PC['Twitter'] + " from the same time last month")


def determine_html_color (text):
    return """<p style="color:red;">""" + text + """</p>"""


#Create Emails
#Line 383 change to ">For the month of, """ + date + """....etc 
const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)

# Change this to Monthly Social Update - 
newMail.Subject = "Weekly Social Update - " + date_header
newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
newMail.HTMLBody = """ 
<h4 style="font-weight: normal;">Hi all - </h4>
<h4 style="font-weight: normal;">For  the date range, """ + date  + """, below are the top performers on <b>Social Media among the news competitive set.</b></h4>
<h4>Total Interactions (Facebook/Instagram/Twitter combined): </h4>""" + pacing_statement_color + """<ol>""" \
+ """<li>""" + color_code_fox_news(Total_Interactions['Site'].iloc[0]) + Total_Interactions['Site'].iloc[0] + " - " + Total_Interactions['Total Interactions'].iloc[0] + ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[0] + " interactions per 1k followers/fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Total_Interactions['Site'].iloc[1]) + Total_Interactions['Site'].iloc[1] + " - " + Total_Interactions['Total Interactions'].iloc[1] + ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[1] + " interactions per 1k followers/fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Total_Interactions['Site'].iloc[2]) + Total_Interactions['Site'].iloc[2] + " - " + Total_Interactions['Total Interactions'].iloc[2] + ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[2] + " interactions per 1k followers/fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Total_Interactions['Site'].iloc[3]) + Total_Interactions['Site'].iloc[3] + " - " + Total_Interactions['Total Interactions'].iloc[3] + ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[3] + " interactions per 1k followers/fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Total_Interactions['Site'].iloc[4]) + Total_Interactions['Site'].iloc[4] + " - " + Total_Interactions['Total Interactions'].iloc[4] + ' total interactions, with ' + Total_Interactions['Interactions per 1k Followers'].iloc[4] + " interactions per 1k followers/fans." """</li>""" \
    + """</oL>""" \
+ """<h4>Facebook Interactions: </h4>""" + fb_pacing_statement_color + """<ol>""" \
+ """<li>""" + color_code_fox_news(Facebook['Site.1'].iloc[0]) + Facebook['Site.1'].iloc[0]  + " - " + Facebook['Facebook Interactions'].iloc[0] + ' Facebook interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[0] + " interactions per 1k fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Facebook['Site.1'].iloc[1]) + Facebook['Site.1'].iloc[1] + " - " + Facebook['Facebook Interactions'].iloc[1] + ' Facebook interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[1] + " interactions per 1k fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Facebook['Site.1'].iloc[2]) + Facebook['Site.1'].iloc[2] + " - " + Facebook['Facebook Interactions'].iloc[2] + ' Facebook interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[2] + " interactions per 1k fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Facebook['Site.1'].iloc[3]) + Facebook['Site.1'].iloc[3] + " - " + Facebook['Facebook Interactions'].iloc[3] + ' Facebook interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[3] + " interactions per 1k fans." """</li>""" \
+ """<li>""" + color_code_fox_news(Facebook['Site.1'].iloc[4]) + Facebook['Site.1'].iloc[4] + " - " + Facebook['Facebook Interactions'].iloc[4] + ' Facebook interactions, with ' + Facebook['Facebook Interacts per 1k Fans'].iloc[4] + " interactions per 1k fans." """</li>""" \
    + """</oL>"""   \
+ """<h4>Instagram Interactions: </h4>""" + ig_pacing_statement_color + """<ol>""" \
+ """<li>""" + color_code_fox_news(Instagram['Site.2'].iloc[0]) + Instagram['Site.2'].iloc[0] + " - " + Instagram['Instagram Interactions'].iloc[0] + ' Instagram interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[0] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Instagram['Site.2'].iloc[1]) + Instagram['Site.2'].iloc[1] + " - " + Instagram['Instagram Interactions'].iloc[1] + ' Instagram interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[1] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Instagram['Site.2'].iloc[2]) + Instagram['Site.2'].iloc[2] + " - " + Instagram['Instagram Interactions'].iloc[2] + ' Instagram interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[2] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Instagram['Site.2'].iloc[3]) + Instagram['Site.2'].iloc[3] + " - " + Instagram['Instagram Interactions'].iloc[3] + ' Instagram interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[3] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Instagram['Site.2'].iloc[4]) + Instagram['Site.2'].iloc[4] + " - " + Instagram['Instagram Interactions'].iloc[4] + ' Instagram interactions, with ' + Instagram['Instagram Interacts per 1k Fans'].iloc[4] + " interactions per 1k followers." """</li>""" \
    + """</oL>"""   \
 + """<h4>Twitter Interactions: </h4>""" + tw_pacing_statement_color + """<ol>""" \
+ """<li>""" + color_code_fox_news(Twitter['Site.3'].iloc[0]) + Twitter['Site.3'].iloc[0] + " - " + Twitter['Twitter Interactions'].iloc[0] + ' Twitter interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[0] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Twitter['Site.3'].iloc[1]) + Twitter['Site.3'].iloc[1] + " - " + Twitter['Twitter Interactions'].iloc[1] + ' Twitter interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[1] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Twitter['Site.3'].iloc[2]) + Twitter['Site.3'].iloc[2] + " - " + Twitter['Twitter Interactions'].iloc[2] + ' Twitter interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[2] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Twitter['Site.3'].iloc[3]) + Twitter['Site.3'].iloc[3] + " - " + Twitter['Twitter Interactions'].iloc[3] + ' Twitter interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[3] + " interactions per 1k followers." """</li>""" \
+ """<li>""" + color_code_fox_news(Twitter['Site.3'].iloc[4]) + Twitter['Site.3'].iloc[4] + " - " + Twitter['Twitter Interactions'].iloc[4] + ' Twitter interactions, with ' + Twitter['Twitter Interacts per 1k Fans'].iloc[4] + " interactions per 1k followers." """</li>""" \
    + """</oL>"""  \
+"""<h6>Source: </b>Emplifi</h6>""" \
+ """<h4> Referral traffic from Facebook, Instagram, and Twitter as well as FBIA traffic onto FoxNews.com: </h4>""" \
+ """<h4> Total:</h4>""" \
+ """<ul>""" \
   +  """<li>""" + str(human_format(pageviews['Total'].iloc[0])) + " page views, " + pageviews_PC['Total_Trend'].iloc[0] + " " + pageviews_PC['Total'].iloc[0] + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(uniques['Total'].iloc[0])) + " unique devices, " + uniques_PC['Total_Trend'].iloc[0] + " " + uniques_PC['Total'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(minutesspent['Total'].iloc[0])) + " minutes spent, " + minutesspent_PC['Total_Trend'].iloc[0] + " " + minutesspent_PC['Total'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """</ul>"""  \
+ """<h4> Facebook & FBIA:</h4>""" \
+ """<ul>""" \
   +  """<li>""" + str(human_format(pageviews['Facebook & FBIA'].iloc[0])) + " page views, " + pageviews_PC['Facebook_Trend'].iloc[0] + " " + pageviews_PC['Facebook & FBIA'].iloc[0] + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(uniques['Facebook & FBIA'].iloc[0])) + " unique devices, " + uniques_PC['Facebook_Trend'].iloc[0] + " " + uniques_PC['Facebook & FBIA'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(minutesspent['Facebook & FBIA'].iloc[0])) + " minutes spent, " + minutesspent_PC['Facebook_Trend'].iloc[0] + " " + minutesspent_PC['Facebook & FBIA'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """</ul>""" \
+ """<h4> Instagram:</h4>""" \
+ """<ul>""" \
   +  """<li>""" + str(human_format(pageviews['Instagram'].iloc[0])) + " page views, " + pageviews_PC['Instagram_Trend'].iloc[0] + " " + pageviews_PC['Instagram'].iloc[0] + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(uniques['Instagram'].iloc[0])) + " unique devices, " + uniques_PC['Instagram_Trend'].iloc[0] + " " + uniques_PC['Instagram'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(minutesspent['Instagram'].iloc[0])) + " minutes spent, " + minutesspent_PC['Instagram_Trend'].iloc[0] + " " + minutesspent_PC['Instagram'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """</ul>""" \
+ """<h4> Twitter:</h4>""" \
+ """<ul>""" \
   +  """<li>""" + str(human_format(pageviews['Twitter'].iloc[0])) + " page views, " + pageviews_PC['Twitter_Trend'].iloc[0] + " " + pageviews_PC['Twitter'].iloc[0] + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(uniques['Twitter'].iloc[0])) + " unique devices, " + uniques_PC['Twitter_Trend'].iloc[0] + " " + uniques_PC['Twitter'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """<li>""" + str(human_format(minutesspent['Twitter'].iloc[0])) + " minutes spent, " + minutesspent_PC['Twitter_Trend'].iloc[0] + " " + minutesspent_PC['Twitter'].iloc[0]  + " from the same time last month" + """</li>""" \
   + """</ul>""" \
   +"""<h6>Source: </b>Adobe Analytics</h6>""" \
   + """<h4 style="font-weight: normal;">Thank you,</h4>"""   \
   + """<h4 style="font-weight: normal;">Kayla</h4>"""
 

newMail.To = "kayla.vayos@fox.com"

#newMail.Body = "Use this for non-HTML body"
#newMail.Attachments.Add(Source='C:\Users\HarrelsonJ\Documents\NewsBiz\Social Reporting\Social_Reporting_07052022.xlsx')
newMail.display() #if you want to display the email. If not, just comment this out/delete
#newMail.send() #Send the email out


#newMail.HTMLBody = """<h3>Please find data attached. Out of the """ +  str(len(fox_adobe_redshift_qa_final)) + """ days measured,
#the number of discrepancies found between the dates """ + date_chosen1.get("""redshift_date_time""") + """ to """ + date_until_incl1.get("""redshift_date_time""") + """ is 
#""" + str(len(fox_adobe_redshift_qa_final[fox_adobe_redshift_qa_final['news_web_match']==False]) + len(fox_adobe_redshift_qa_final[fox_adobe_redshift_qa_final['news_app_match']==False]) \
   # + len(fox_adobe_redshift_qa_final[fox_adobe_redshift_qa_final['business_web_match']==False]) + len(fox_adobe_redshift_qa_final[fox_adobe_redshift_qa_final['business_app_match']==False]))