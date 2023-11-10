import pandas as pd
import numpy as np
import datetime as dt
import math

pd.set_option('display.max_colwidth', None)

def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])

# Date
date = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Outkick Automation/20230918_Outkick_Sports_Comscore_Aug2023.xlsx', sheet_name='Email Auto')

dateform = date['Date'].iloc[0]
print(dateform)

datetext = date['Date To Text'].loc[0]
print(datetext)

# Data Pull
Oktrend = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Outkick Automation/20230918_Outkick_Sports_Comscore_Aug2023.xlsx', sheet_name='Outkick Trend', skiprows=10)
print(Oktrend)

KPI = Oktrend[[dateform]]
print(KPI)

MultiUniq = KPI[dateform].iloc[1]
MultiUniq = MultiUniq * 1000
print(MultiUniq)

MultiViews = KPI[dateform].iloc[2]
MultiViews = MultiViews * 1000000
print(MultiViews)

MultiMinutes = KPI[dateform].iloc[3]
MultiMinutes = MultiMinutes * 1000000
print(MultiMinutes)

Rank = Oktrend[['Rank']]
print(Rank)

RankUniq = Oktrend['Rank'].iloc[1]
print(RankUniq)

RankViews = Oktrend['Rank'].iloc[2]
print(RankViews)

RankMinutes = Oktrend['Rank'].iloc[3]
print(RankMinutes)

Oktrend['vs. Prior Month'] = Oktrend['vs. Prior Month'].transform(lambda x: '{:,.2%}'.format(x))
PriorMonth = Oktrend[['vs. Prior Month']]
print(PriorMonth)

PriorMonthUniq = Oktrend['vs. Prior Month'].iloc[1]
print(PriorMonthUniq)

PriorMonthViews = Oktrend['vs. Prior Month'].iloc[2]
print(PriorMonthViews)

PriorMonthMinutes = Oktrend['vs. Prior Month'].iloc[3]
print(PriorMonthMinutes)

Oktrend['vs. Prior Year'] = Oktrend['vs. Prior Year'].transform(lambda x: '{:,.2%}'.format(x))
PriorYear = Oktrend[['vs. Prior Year']]
print(PriorYear)

PriorYearUniq = Oktrend['vs. Prior Year'].iloc[1]
print(PriorYearUniq)

PriorYearViews = Oktrend['vs. Prior Year'].iloc[2]
print(PriorYearViews)

PriorYearMinutes = Oktrend['vs. Prior Year'].iloc[3]
print(PriorYearMinutes)

# Competitive set rankings
PrevMonthCompRank = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Outkick Automation/20230918_Outkick_Sports_Comscore_Aug2023.xlsx', sheet_name='Jul23', skiprows=9, usecols= [0,1,2,3,4,5]) # Need to manually change sheet name before running each month
CurrMonthCompRank = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Outkick Automation/20230918_Outkick_Sports_Comscore_Aug2023.xlsx', sheet_name='Aug23', skiprows=9, usecols= [0,1,2,3,4,5]) # Need to manually change sheet name before running each month
PrevMonthCompRank = PrevMonthCompRank.loc[PrevMonthCompRank['Unnamed: 5'] == "OUTKICK.COM"]
print(PrevMonthCompRank)
PrevUniqRank = PrevMonthCompRank['Total Unique Visitors/Viewers (000)'].iloc[0]
print(PrevUniqRank)
PrevViewRank = PrevMonthCompRank['Total Views (MM)'].iloc[0]
print(PrevViewRank)
PrevMinRank = PrevMonthCompRank['Total Minutes (MM)'].iloc[0]
print(PrevMinRank)

CurrMonthCompRank = CurrMonthCompRank.loc[CurrMonthCompRank['Unnamed: 5'] == "OUTKICK.COM"]
print(CurrMonthCompRank)
CurrUniqRank = CurrMonthCompRank['Total Unique Visitors/Viewers (000)'].iloc[0]
print(CurrUniqRank)
CurrViewRank = CurrMonthCompRank['Total Views (MM)'].iloc[0]
print(CurrViewRank)
CurrMinRank = CurrMonthCompRank['Total Minutes (MM)'].iloc[0]
print(CurrMinRank)

# Helper Functions
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

def difference(prev, curr):
    return(abs(curr - prev))

def rise_drop(prev, curr):
    if prev > curr:
        return("""moved down """ + str(difference(prev,curr)) + """ spots  to #""" + str(curr))
    elif prev == curr:
        return("""remained in the """ + str(curr) + """spot""")
    else:
        return("""moved up """ + str(difference(prev,curr)) + """ spots  to #""" + str(curr))


# Test HTML Page
Function_Name = open("GFG-1.html","w")
Function_Name.write(""" 
<h4 style="font-weight: normal;">Hi all - </h4>
<h4 style="font-weight: normal;"> Attached is the updated file with """ + str(datetext) + """ Comscore data added to Outkick.com </h4>
<p><strong><u>Highlights:</u></strong></p>
<p><strong>Current vs. Prior Month Variances</strong></p>
<ul>
    <li>Outkick.com drove in """ + human_format(MultiUniq) + """ multiplatform unique visitors, """ + up_down(PriorMonthUniq) + ' ' + str(PriorMonthUniq) + """ vs. prior month<ul>
            <li>"""+ str(math.trunc(RankUniq)) + """th highest month of unique visitors according to Comscore</li>
        </ul>
    </li>
    <li>Outkick.com drove in """ + human_format(MultiViews) + """ multiplatform views, """ + up_down(PriorMonthViews) + ' ' + str(PriorMonthViews) + """ vs. prior month<ul>
            <li>"""+ str(math.trunc(RankViews)) + """th highest month of views according to Comscore</li>
        </ul>
    </li>
    <li>Outkick.com drove in """ + human_format(MultiMinutes) + """ multiplatform unique visitors, """ + up_down(PriorMonthMinutes) + ' ' + str(PriorMonthMinutes) + """ vs. prior month<ul>
            <li>"""+ str(math.trunc(RankMinutes)) + """th highest month of minutes according to Comscore</li>
        </ul>
    </li>
</ul>
<p><strong>Current vs. Prior Year Variances</strong></p>
<ul>
    <li>Multiplatform unique visitors """ + up_down(PriorYearUniq) + ' ' + str(PriorYearUniq) + """ vs. prior year</li>
    <li>Multiplatform views """ + up_down(PriorYearViews) + ' ' + str(PriorYearViews) + """ vs. prior year</li>
    <li>Multiplatform minutes """ + up_down(PriorYearMinutes) + ' ' + str(PriorYearMinutes) + """ vs. prior year</li>
</ul>
<p><strong>Ranker Status</strong></p>
<ul>
    <li>Outkick.com """ + rise_drop(PrevUniqRank,CurrUniqRank) + """ in terms of unique visitors</li>
    <li>Outkick.com """ + rise_drop(PrevViewRank,CurrViewRank) + """ in terms of views</li>
    <li>Outkick.com """ + rise_drop(PrevMinRank,CurrMinRank) + """ in terms of minutes</li>
</ul>

""")

Function_Name.close()

# # Create Email 
# # ** Note: win32client is not available on MacOS so the following code has not been tested

# ol = win32com.client.Dispatch('Outlook.Application')
# olmailitem = 0x0
# newmail = ol.CreateItem(olmailitem)

# newmail.Subject = 'Monthly Outkick Update - ' + str(datetext) + ''''''
# newmail.HTMLBody = """<h4 style="font-weight: normal;">Hi all - </h4>
# <h4 style="font-weight: normal;"> Attached is the updated file with """ + str(datetext) + """ Comscore data added to Outkick.com </h4>
# <p><strong><u>Highlights:</u></strong></p>
# <p><strong>Current vs. Prior Month Variances</strong></p>
# <ul>
#     <li>Outkick.com drove in """ + human_format(MultiUniq) + """ multiplatform unique visitors, """ + up_down(PriorMonthUniq) + ' ' + str(PriorMonthUniq) + """ vs. prior month<ul>
#             <li>"""+ str(math.trunc(RankUniq)) + """th highest month of unique visitors according to Comscore</li>
#         </ul>
#     </li>
#     <li>Outkick.com drove in """ + human_format(MultiViews) + """ multiplatform views, """ + up_down(PriorMonthViews) + ' ' + str(PriorMonthViews) + """ vs. prior month<ul>
#             <li>"""+ str(math.trunc(RankViews)) + """th highest month of views according to Comscore</li>
#         </ul>
#     </li>
#     <li>Outkick.com drove in """ + human_format(MultiMinutes) + """ multiplatform unique visitors, """ + up_down(PriorMonthMinutes) + ' ' + str(PriorMonthMinutes) + """ vs. prior month<ul>
#             <li>"""+ str(math.trunc(RankMinutes)) + """th highest month of minutes according to Comscore</li>
#         </ul>
#     </li>
# </ul>
# <p><strong>Current vs. Prior Year Variances</strong></p>
# <ul>
#     <li>Multiplatform unique visitors """ + up_down(PriorYearUniq) + ' ' + str(PriorYearUniq) + """ vs. prior year</li>
#     <li>Multiplatform views """ + up_down(PriorYearViews) + ' ' + str(PriorYearViews) + """ vs. prior year</li>
#     <li>Multiplatform minutes """ + up_down(PriorYearMinutes) + ' ' + str(PriorYearMinutes) + """ vs. prior year</li>
# </ul>
# <p><strong>Ranker Status</strong></p>
# <ul>
#     <li>Outkick.com """ + rise_drop(PrevUniqRank,CurrUniqRank) + """ in terms of unique visitors</li>
#     <li>Outkick.com """ + rise_drop(PrevViewRank,CurrViewRank) + """ in terms of views</li>
#     <li>Outkick.com """ + rise_drop(PrevMinRank,CurrMinRank) + """ in terms of minutes</li>
# </ul>"""

# newmail.To = "taylor.caruso@fox.com"
# newmail.display()
