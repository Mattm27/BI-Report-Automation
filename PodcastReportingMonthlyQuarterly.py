import pandas as pd
import numpy as np
import datetime as dt

# Important!!! Make sure to update any lines that contain FYxx Qx to contain proper year and quarter (i.e line 82, line 219, line 234)
pd.set_option('display.max_colwidth', None)

# This function below converts numbers to an easier to read string format
def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])

# Date
overview = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Email Auto') # Reads the Email Auto sheet from Podcast Report excel file and saves as overview
Date = overview[['Date']] # Pulls in the column labeled date as Date
print(Date['Date'])


# Fox News Hourly Update
FNHour = overview[['Downloads', 'Listeners', 'DMoM Change', 'DYoY Change', 'LMoM Change', 'LYoY Change']] # Selects columns in sheet with given names

FNHour['DMoM Change'] = FNHour['DMoM Change'].transform(lambda x: '{:,.2%}'.format(x)) # Converts decimal number to percentage (%)
FNHour['DYoY Change'] = FNHour['DYoY Change'].transform(lambda x: '{:,.2%}'.format(x))
FNHour['LMoM Change'] = FNHour['LMoM Change'].transform(lambda x: '{:,.2%}'.format(x))
FNHour['LYoY Change'] = FNHour['LYoY Change'].transform(lambda x: '{:,.2%}'.format(x))

# Fox Business Hourly report
FBNHour = overview[['Downloads1', 'Listeners1', 'DMoM Change1', 'DYoY Change1', 'LMoM Change1', 'LYoY Change1']]

FBNHour['DMoM Change1'] = FBNHour['DMoM Change1'].transform(lambda x: '{:,.2%}'.format(x))
FBNHour['DYoY Change1'] = FBNHour['DYoY Change1'].transform(lambda x: '{:,.2%}'.format(x))
FBNHour['LMoM Change1'] = FBNHour['LMoM Change1'].transform(lambda x: '{:,.2%}'.format(x))
FBNHour['LYoY Change1'] = FBNHour['LYoY Change1'].transform(lambda x: '{:,.2%}'.format(x))

# Complete Table
TotalDown = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Deliverable', skiprows=7, usecols=[2,4,6,8]) # Brings in new sheet from file and names it 
TotalDown = TotalDown.rename(columns={TotalDown.columns[1]: 'Unique Downloads'}) # Renames first column to Unique Downloads
TotalDown = TotalDown[TotalDown['Unique Downloads'] > 100000] # Only selects observations with unique downloads greater than 100000

MaxDown = TotalDown.loc[TotalDown['% Change'].idxmax()] # Selects podcast with largest percent change
MaxDown['% Change'] = format(MaxDown['% Change'],'.2%')
print(MaxDown)

TotalList = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Deliverable', skiprows=7, usecols=[14,16,18,20])
TotalList = TotalList.rename(columns={TotalList.columns[1]: 'Unique Listeners'})  
TotalList = TotalList[TotalList['Unique Listeners'] > 100000]


MaxList = TotalList.loc[TotalList['% Change.2'].idxmax()]
MaxList['% Change.2'] = format(MaxList['% Change.2'],'.2%')
print(MaxList)

# Max Downloads
MaxDown1 = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Unique Downloads', skiprows=6)
MaxDown1 = MaxDown1[['Podcast Name', 'Is the latest month the max?', 'Latest Month Downloads']]

MaxDown1 = MaxDown1.loc[MaxDown1['Is the latest month the max?'] == True]
print(MaxDown1)

# Max Listeners
MaxList1 = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Unique Listeners', skiprows=8)
MaxList1 = MaxList1[['Podcast Name', 'Is the latest month the max?', 'Latest Month Downloads']]

MaxList1 = MaxList1.loc[MaxList1['Is the latest month the max?'] == True]
print(MaxList1)

# Quarter End
QuarterDown = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Quarterly Downloads')
QuarterTotalDown = QuarterDown.loc[QuarterDown['Podcast'] == 'Total'] # Grabs total number of podcast downloads
QuarterTotalDown['Quarter Difference Percent'] = QuarterTotalDown['Quarter Difference Percent'].transform(lambda x: '{:,.2%}'.format(x)) # Changed column names in excel sheet
QuarterTotalDown['Year Difference Percent'] = QuarterTotalDown['Year Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))
print(QuarterTotalDown)
HourDown = QuarterDown.loc[QuarterDown['Podcast'] == 'Fox Business Hourly Report'] # Selects just the Fox Business Hourly Report
HourDown['Quarter Difference Percent'] = HourDown['Quarter Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))
HourDown['Year Difference Percent'] = HourDown['Year Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))
print(HourDown)

Notable3 = QuarterDown[(QuarterDown['Podcast'] != 'Fox Business Hourly Report') & (QuarterDown['Podcast'] != 'Fox News Hourly Update') & (QuarterDown['Podcast'] != 'Total')]
Notable3 = Notable3[Notable3['FY23 Q4'] > 100000]
Notable3 = Notable3.nlargest(3,['Quarter Difference Percent'])
Notable3['Quarter Difference Percent'] = Notable3['Quarter Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))
print(Notable3)



QuarterList = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Quarterly Listens')
QuarterTotalList = QuarterList.loc[QuarterList['Podcast'] == 'Total']
QuarterTotalList['Quarter Difference Percent'] = QuarterTotalList['Quarter Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))
QuarterTotalList['Year Difference Percent'] = QuarterTotalList['Year Difference Percent'].transform(lambda x: '{:,.2%}'.format(x))

QuarterRanks = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/Podcast Report Monthly:Quarterly Automation/Podcast Report v2 100123.xlsm', sheet_name='Top Performing Podcasts by Uniq')
print(QuarterRanks)



# Helper Functions used for drafting email depending on data values
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

def sim(str1,str2,num2):
    if str1 == str2:
        return(""" and in Unique Listeners with """ + human_format(num2) + """ during""")
    else:
        return(""", while <u>""" + str2 + """</u> set a new historic record for Unique Listeners with """ + human_format(num2) + """ during""")

def FBNHelp(str1,str2):
    index1 = str1.find('-')
    index2 = str2.find('-')
    if index1 == -1 and index2 == -1:
        return("""<u> The Fox Business Hourly Report </u> saw an increase in both Unique Downloads and Unique Listeners in """)
    elif index1 == -1 and index2 != -1:
        return(""" <u> The Fox Business Hourly Report </u> saw an increase in Unique Downloads but a decrease in Unique Listeners in """)
    elif index1 != -1 and index2 == -1:
        return(""" <u> The Fox Business Hourly Report </u> saw a decrease in Unique Downloads but a increase in Unique Listeners in """)
    else:
        return(""" <u> The Fox Business Hourly report </u> saw a decrease in both Unique Downloads and Unique Listeners in """)

def QuarterEndHelp(NumDown, NumList, PerDown, PerList):
    index1 = PerDown.find('-')
    index2 = PerList.find('-')
    if index1 == -1 and index2 == -1:
        return("""Unique Downloads and Unique Listeners were up vs. the prior quarter with """ + human_format(NumDown) + """ Unique Downloads, and """ + human_format(NumList) + """ Unique Listeners""")
    elif index1 == -1 and index2 != -1:
        return("""Unique Downloads were up vs. the prior quarter with """ + human_format(NumDown) + """ downloads, while Unique Listeners were down vs. the prior quarter with """ + human_format(NumList) + """ listeners""")
    elif index1 != -1 and index2 == -1:
        return("""Unique Listeners were up vs. the prior quarter with """ + human_format(NumList) + """ listeners, while Unique Downloads were down vs. the prior quarter with """ + human_format(NumDown) + """ downloads""")
    else:
        return("""Unique Downloads and Unique Listeners were down vs. the prior quarter with """ + human_format(NumDown) + """ Unique Downloads, and """ + human_format(NumList) + """ Unique Listeners """)

def QuarterEndHelp1(PerDownYear, PerListYear):
    index1 = PerDownYear.find('-')
    index2 = PerListYear.find('-')
    if index1 == -1 and index2 == -1:
        return("""Both were up vs. prior year, with unique downloads up """ + PerDownYear + """ and unique listeners up """ + PerListYear)
    elif index1 == -1 and index2 != -1:
        return("""Unique Downloads were up """ + PerDownYear  + """ vs. the prior year, while Unique Listeners were down """ + PerListYear + """ vs. the prior year""")
    elif index1 != -1 and index2 == -1:
        return("""Unique Listeners were up """ + PerListYear  + """ vs. the prior year, while Downloads were down """ + PerDownYear + """ vs. the prior year""")
    else:
        return("""Both were down vs. prior year, with unique downloads down """ + PerDownYear + """ and unique listeners down """ + PerListYear)

def whichQ(date):
    indexQ3 = date.find('March')
    indexQ4 = date.find('June')
    indexQ1 = date.find('September')
    indexQ2 = date.find('December')
    if indexQ3 != -1:
        return("""Q3""")
    elif indexQ4 != -1:
        return("""Q4""")
    elif indexQ1 != -1:
        return("""Q1""")
    elif indexQ2 != -1:
        return("""Q2""")

def QuarterEnd(date, QList,QDown,PQList,PQDown):
    index = date.find('March')
    index1 = date.find('June')
    index2 = date.find('September')
    index3 = date.find('December')
    if index == -1 and index1 == -1 and index2 == -1 and index3 == -1:
        return("""<h4 style="font-weight: normal;">For the month of """ + str(Date['Date'].iloc[0]) + """, below are key takeaways from Fox News and Fox Business podcast performance. Please let me know if you have any questions. </h4>
                  <p style = "margin-bottom:0;">Thanks!</p>
                  <p style = "margin :0; padding-top:0;">Kayla</p>""")
    else:
        return("""<h4 style="font-weight: normal;">Attached are podcast unique downloads and unique audiences for month end (""" + date + """) and quarter-end (""" + whichQ(date) + """) with a summary of key takeaways below. Please let me know if you have any questions.</h4>
                  <p style = "margin-bottom:0;">Thanks!</p>
                  <p style = "margin :0; padding-top:0;">Kayla</p>
                  <p style = "margin-bottom:0;"><u><b>Quarter:</b></u></p>
                  <p style = "margin :0; padding-top:0;">Unique Listeners: """ + human_format(QList) + """</p>
                  <p style = "margin :0; padding-top:0;">Downloads: """ + human_format(QDown) + """</p>
                  <p style = "margin-bottom:0;"><u><b>Last Year Quarter:</b></u></p>
                  <p style = "margin :0; padding-top:0;">Unique Listeners: """ + human_format(PQList) + """</p>
                  <p style = "margin :0; padding-top:0;">Downloads: """ + human_format(PQDown) + """</p>""")

def QuarterEnd1(date, Down, List, DownChange, ListChange, Down1Name, Down1Count, Down2Name, Down2Count, Down3Name, Down3Count, DownChangeYear, ListChangeYear, HourDown, HourPerc, Note1Name, Note1Perc, Note1Down, Note2Name, Note2Perc, Note2Down, Note3Name, Note3Perc, Note3Down):
    index = date.find('March')
    index1 = date.find('June')
    index2 = date.find('September')
    index3 = date.find('December')
    if index == -1 and index1 == -1 and index2 == -1 and index3 == -1:
        return("""""")           
    else:
        return("""<p><u>Quarter-End Highlights:</u></p>
                    <ul>
                        <li> """ + QuarterEndHelp(Down,List,DownChange,ListChange) + """ <ul>
                            <li> """ + QuarterEndHelp1(DownChangeYear, ListChangeYear) + """</li>
                            </ul>
                        <li> The most downloaded programs include the <u>""" + Down1Name + """</u> (""" + human_format(Down1Count) + """), <u>""" + Down2Name + """</u> (""" + human_format(Down2Count) + """), and the <u>""" + Down3Name + """</u> (""" + human_format(Down3Count) + """) <ul>
                        <li> <u> The Fox News Business Report </u> reached """ + human_format(HourDown) + """ Unique Downlaods this Quarter, """ + up_down(HourPerc) +""" """ + HourPerc + """ vs. prior year. <ul>
                            <li> Other notable increases in Unique Downloads vs. the prior quarter: <ul>
                                <li><u>""" + Note1Name +"""</u> """ + up_down(Note1Perc) + """ """ + Note1Perc + """ with """ + human_format(Note1Down) + """ downloads </li>
                                <li><u>""" + Note2Name +"""</u> """ + up_down(Note2Perc) + """ """ + Note2Perc + """ with """ + human_format(Note2Down) + """ downloads </li>
                                <li><u>""" + Note3Name +"""</u> """ + up_down(Note3Perc) + """ """ + Note3Perc +""" with """ + human_format(Note3Down) + """ downloads </li>
                            </ul>
                        </ul>""")

    


# Create Email 
# ** Note: win32client is not available on MacOS so the following code has not been tested

# ol = win32com.client.Dispatch('Outlook.Application')
# olmailitem = 0x0
# newmail = ol.CreateItem(olmailitem)

# newmail.Subject = 'Monthly Podcast Update - ' + str(Date['Date'].iloc[0])
# newmail.HTMLBody = """ 
# <h4 style="font-weight: normal;">Hi all - </h4>
# """ + QuarterEnd(Date['Date'].iloc[0],QuarterTotalDown['FY23 Q4'].iloc[0],QuarterTotalList['FY23 Q4'].iloc[0],QuarterTotalDown['FY22 Q4'].iloc[0], QuarterTotalList['FY22 Q4'].iloc[0]) + """
# <p><u>Key Takeaways:</u></p>
# <ul>
#     <li>""" + str(Date['Date'].iloc[0]) + """ saw """ + human_format(FNHour['Downloads'].iloc[0]) + """ Unique Downloads for <u>Fox News Hourly Update</u>, """ + up_down(FNHour['DMoM Change'].iloc[0]) + """ """ + FNHour['DMoM Change'].iloc[0] + """ vs. the prior month and """ + up_down(FNHour['DYoY Change'].iloc[0]) + """ """ + FNHour['DYoY Change'].iloc[0] + """ vs. the prior year (""" + str(Date['Date'].iloc[2]) + """-""" + str(Date['Date'].iloc[1]) + """) <ul>
#         <li> <u>Fox News Hourly Update</u> reached """ + human_format(FNHour['Listeners'].iloc[0]) + """ Unique Listeners for """ + str(Date['Date'].iloc[0]) + """, """ + up_down(FNHour['LMoM Change'].iloc[0]) + """ """ + FNHour['LMoM Change'].iloc[0] + """ vs. the prior month and """ + up_down(FNHour['LYoY Change'].iloc[0]) + """ """ + FNHour['LYoY Change'].iloc[0] + """ vs. the prior year </li>
#             </ul>
#         </li>
#     <li> <u>""" + MaxDown['Podcast Name']+"""</u> Podcasts saw the largest increase in Unique Downloads vs. the prior month being up """ + MaxDown['% Change'] + """, while <u>""" + MaxList['Podcast Name.1'] + """</u> saw the largest increase in Unique Listeners vs. the prior month being up """ + MaxList['% Change.2'] + """</li>
#     <li> <u>""" + MaxDown1['Podcast Name'].iloc[0] +"""</u>Podcasts reached a historic high in Unique Downloads with """ + human_format(MaxDown1['Latest Month Downloads'].iloc[0]) + sim(MaxDown1['Podcast Name'].iloc[0],MaxList1['Podcast Name'].iloc[0],MaxList1['Latest Month Downloads'].iloc[0]) + """ """ + str(Date['Date'].iloc[0]) + """ </li>
#     <li> """ + FBNHelp(FBNHour['DMoM Change1'].iloc[0],FBNHour['LMoM Change1'].iloc[0]) + str(Date['Date'].iloc[0]) + """<ul>
#         <li> In """ + str(Date['Date'].iloc[0]) + """, The Fox Business Hourly Report generated """ + human_format(FBNHour['Downloads1'].iloc[0]) + """ Unique Downloads and """ + human_format(FBNHour['Listeners1'].iloc[0]) + """ Unique Listeners </li>
#             </ul>
#         </li>
#     </li>
# </ul>
# """ + QuarterEnd1(str(Date['Date'].iloc[0]), QuarterTotalDown['FY23 Q4'].iloc[0], QuarterTotalList['FY23 Q4'].iloc[0], QuarterTotalDown['Quarter Difference Percent'].iloc[0], QuarterTotalList['Quarter Difference Percent'].iloc[0], QuarterRanks['PODCAST'].iloc[0], QuarterRanks['TOTAL DOWNLOADS'].iloc[0],QuarterRanks['PODCAST'].iloc[1], QuarterRanks['TOTAL DOWNLOADS'].iloc[1], QuarterRanks['PODCAST'].iloc[2], QuarterRanks['TOTAL DOWNLOADS'].iloc[2], QuarterTotalDown['Year Difference Percent'].iloc[0], QuarterTotalList['Year Difference Percent'].iloc[0], HourDown['FY23 Q4'].iloc[0], HourDown['Year Difference Percent'].iloc[0], Notable3['Podcast'].iloc[0], Notable3['Quarter Difference Percent'].iloc[0], Notable3['FY23 Q4'].iloc[0], Notable3['Podcast'].iloc[1], Notable3['Quarter Difference Percent'].iloc[1], Notable3['FY23 Q4'].iloc[1], Notable3['Podcast'].iloc[2], Notable3['Quarter Difference Percent'].iloc[2], Notable3['FY23 Q4'].iloc[2]) + """
# <p><br></p>
# """

# newMail.To = "kayla.vayos@fox.com"
# newMail.display()
# newMail.send()

# Test HTML Page
Function_Name = open("PODTEST.html","w")

# For the QuarterEnd Function call, manually update Fiscal Year and Quarter each time report is ran
Function_Name.write(""" 
<h4 style="font-weight: normal;">Hi all - </h4>
""" + QuarterEnd(Date['Date'].iloc[0],QuarterTotalDown['FY24 Q1'].iloc[0],QuarterTotalList['FY22 Q1'].iloc[0],QuarterTotalDown['FY23 Q1'].iloc[0], QuarterTotalList['FY23 Q1'].iloc[0]) + """ 
<p><u>Key Takeaways:</u></p>
<ul>
    <li>""" + str(Date['Date'].iloc[0]) + """ saw """ + human_format(FNHour['Downloads'].iloc[0]) + """ Unique Downloads for <u>Fox News Hourly Update</u>, """ + up_down(FNHour['DMoM Change'].iloc[0]) + """ """ + FNHour['DMoM Change'].iloc[0] + """ vs. the prior month and """ + up_down(FNHour['DYoY Change'].iloc[0]) + """ """ + FNHour['DYoY Change'].iloc[0] + """ vs. the prior year (""" + str(Date['Date'].iloc[2]) + """-""" + str(Date['Date'].iloc[1]) + """) <ul>
        <li> <u>Fox News Hourly Update</u> reached """ + human_format(FNHour['Listeners'].iloc[0]) + """ Unique Listeners for """ + str(Date['Date'].iloc[0]) + """, """ + up_down(FNHour['LMoM Change'].iloc[0]) + """ """ + FNHour['LMoM Change'].iloc[0] + """ vs. the prior month and """ + up_down(FNHour['LYoY Change'].iloc[0]) + """ """ + FNHour['LYoY Change'].iloc[0] + """ vs. the prior year </li>
            </ul>
        </li>
    <li> <u>""" + MaxDown['Podcast Name']+"""</u> Podcasts saw the largest increase in Unique Downloads vs. the prior month being up """ + MaxDown['% Change'] + """, while <u>""" + MaxList['Podcast Name.1'] + """</u> saw the largest increase in Unique Listeners vs. the prior month being up """ + MaxList['% Change.2'] + """</li>
    <li> <u>""" + MaxDown1['Podcast Name'].iloc[0] +"""</u>Podcasts reached a historic high in Unique Downloads with """ + human_format(MaxDown1['Latest Month Downloads'].iloc[0]) + sim(MaxDown1['Podcast Name'].iloc[0],MaxList1['Podcast Name'].iloc[0],MaxList1['Latest Month Downloads'].iloc[0]) + """ """ + str(Date['Date'].iloc[0]) + """ </li>
    <li> """ + FBNHelp(FBNHour['DMoM Change1'].iloc[0],FBNHour['LMoM Change1'].iloc[0]) + str(Date['Date'].iloc[0]) + """<ul>
        <li> In """ + str(Date['Date'].iloc[0]) + """, The Fox Business Hourly Report generated """ + human_format(FBNHour['Downloads1'].iloc[0]) + """ Unique Downloads and """ + human_format(FBNHour['Listeners1'].iloc[0]) + """ Unique Listeners </li>
            </ul>
        </li>
    </li>
</ul>
""" + QuarterEnd1(str(Date['Date'].iloc[0]), QuarterTotalDown['FY24 Q1'].iloc[0], QuarterTotalList['FY24 Q1'].iloc[0], QuarterTotalDown['Quarter Difference Percent'].iloc[0], QuarterTotalList['Quarter Difference Percent'].iloc[0], QuarterRanks['PODCAST'].iloc[0], QuarterRanks['TOTAL DOWNLOADS'].iloc[0],QuarterRanks['PODCAST'].iloc[1], QuarterRanks['TOTAL DOWNLOADS'].iloc[1], QuarterRanks['PODCAST'].iloc[2], QuarterRanks['TOTAL DOWNLOADS'].iloc[2], QuarterTotalDown['Year Difference Percent'].iloc[0], QuarterTotalList['Year Difference Percent'].iloc[0], HourDown['FY24 Q1'].iloc[0], HourDown['Year Difference Percent'].iloc[0], Notable3['Podcast'].iloc[0], Notable3['Quarter Difference Percent'].iloc[0], Notable3['FY24 Q1'].iloc[0], Notable3['Podcast'].iloc[1], Notable3['Quarter Difference Percent'].iloc[1], Notable3['FY24 Q1'].iloc[1], Notable3['Podcast'].iloc[2], Notable3['Quarter Difference Percent'].iloc[2], Notable3['FY24 Q1'].iloc[2]) + """
<p><br></p>
""")
Function_Name.close()