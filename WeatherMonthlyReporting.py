import pandas as pd
import numpy as np
from datetime import datetime

pd.set_option('display.max_colwidth', None)

def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])

# Pull in sheets
Summary = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230930_Fox_Weather_Monthly_dashboard_final.xlsx', sheet_name='Summary - Weekly')
LKM = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230930_Fox_Weather_Monthly_dashboard_final.xlsx', sheet_name='LKM_data_month (2)')
Auto = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230930_Fox_Weather_Monthly_dashboard_final.xlsx', sheet_name='Email Auto')
LKM2 = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230930_Fox_Weather_Monthly_dashboard_final.xlsx', sheet_name='LKM_data_month (2)', skiprows=50)

# Rename Columns
LKM = LKM.rename(columns={LKM.columns[0]: 'Col1'})
LKM = LKM.rename(columns={LKM.columns[1]: 'Col2'})
LKM = LKM.rename(columns={LKM.columns[2]: 'Col3'})
LKM = LKM.rename(columns={LKM.columns[3]: 'Col4'})
LKM = LKM.rename(columns={LKM.columns[4]: 'Col5'})
LKM = LKM.rename(columns={LKM.columns[5]: 'Col6'})
LKM = LKM.rename(columns={LKM.columns[6]: 'Col7'})
LKM = LKM.rename(columns={LKM.columns[7]: 'Col8'})
LKM = LKM.rename(columns={LKM.columns[8]: 'Col9'})
print(LKM['Col7'])

Summary = Summary.rename(columns={Summary.columns[0]: 'Col1'})
Summary = Summary.rename(columns={Summary.columns[1]: 'Col2'})
Summary = Summary.rename(columns={Summary.columns[2]: 'Col3'})
Summary = Summary.rename(columns={Summary.columns[3]: 'Col4'})
Summary = Summary.rename(columns={Summary.columns[4]: 'Col5'})
Summary = Summary.rename(columns={Summary.columns[5]: 'Col6'})
Summary = Summary.rename(columns={Summary.columns[6]: 'Col7'})
Summary = Summary.rename(columns={Summary.columns[7]: 'Col8'})
Summary = Summary.rename(columns={Summary.columns[8]: 'Col9'})
Summary = Summary.rename(columns={Summary.columns[9]: 'Col10'})
Summary = Summary.rename(columns={Summary.columns[10]: 'Col11'})
Summary = Summary.rename(columns={Summary.columns[11]: 'Col12'})
Summary = Summary.rename(columns={Summary.columns[12]: 'Col13'})
Summary = Summary.rename(columns={Summary.columns[13]: 'Col14'})
Summary = Summary.rename(columns={Summary.columns[14]: 'Col15'})
Summary = Summary.rename(columns={Summary.columns[15]: 'Col16'})
Summary = Summary.rename(columns={Summary.columns[16]: 'Col17'})
Summary = Summary.rename(columns={Summary.columns[17]: 'Col18'})
Summary = Summary.rename(columns={Summary.columns[18]: 'Col19'})
Summary = Summary.rename(columns={Summary.columns[19]: 'Col20'})

# Extract Date
Date = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230930_Fox_Weather_Monthly_dashboard_final.xlsx', sheet_name='Summary - Month', header=None)
Date1= Date[18].iloc[0]
print(Date1)
stringdate = str(Date1)
date_obj = datetime.strptime(stringdate,"%Y-%m-%d %H:%M:%S")
formatted_date = date_obj.strftime("%m/%Y")
print(formatted_date)


# Stream Engagement
Unique_Viewing = LKM.loc[LKM['Col2'] == 'Unique Viewers']
Unique_Viewing['Col4'] = Unique_Viewing['Col4'].transform(lambda x: '{:,.0%}'.format(x))
Unique_Viewing['Col5'] = Unique_Viewing['Col5'].transform(lambda x: '{:,.0%}'.format(x))
Unique_Viewing['Col7'] = Unique_Viewing['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(Unique_Viewing)
Session_Freq = LKM.loc[LKM['Col2'] == 'Session Frequency']
Session_Freq['Col4'] = Session_Freq['Col4'].transform(lambda x: '{:,.0%}'.format(x))
Session_Freq['Col5'] = Session_Freq['Col5'].transform(lambda x: '{:,.0%}'.format(x))
Session_Freq['Col7'] = Session_Freq['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(Session_Freq)
Dwell = LKM.loc[LKM['Col2'] == 'Dwell Time']
Dwell['Col4'] = Dwell['Col4'].transform(lambda x: '{:,.0%}'.format(x))
Dwell['Col5'] = Dwell['Col5'].transform(lambda x: '{:,.0%}'.format(x))
Dwell['Col7'] = Dwell['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(Dwell)
Total_View = LKM.loc[LKM['Col2'] == 'Total View Time']
Total_View['Col4'] = Total_View['Col4'].transform(lambda x: '{:,.0%}'.format(x))
Total_View['Col5'] = Total_View['Col5'].transform(lambda x: '{:,.0%}'.format(x))
Total_View['Col7'] = Total_View['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(Total_View['Col7'])
print(Total_View)
AMA = LKM.loc[LKM['Col2'] == 'AMA']
AMA['Col4'] = AMA['Col4'].transform(lambda x: '{:,.0%}'.format(x))
AMA['Col5'] = AMA['Col5'].transform(lambda x: '{:,.0%}'.format(x))
AMA['Col7'] = AMA['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(AMA)
# Minutes Spent (Must select 1st row)
AppMin = LKM.loc[LKM['Col2'] == 'FOX Weather App']
AppMin['Col4'] = AppMin['Col4'].transform(lambda x: '{:,.0%}'.format(x))
AppMin['Col5'] = AppMin['Col5'].transform(lambda x: '{:,.0%}'.format(x))
AppMin['Col6'] = AppMin['Col6'].transform(lambda x: '{:,.0%}'.format(x))
PercentOfTotal = LKM2['% of Total']
PercentOfTotal = PercentOfTotal.transform(lambda x: '{:,.0%}'.format(x))

WebMin = LKM.loc[LKM['Col2'] == 'FOXWeather.com']
WebMin['Col4'] = WebMin['Col4'].transform(lambda x: '{:,.0%}'.format(x))
WebMin['Col5'] = WebMin['Col5'].transform(lambda x: '{:,.0%}'.format(x))
WebMin['Col7'] = WebMin['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(WebMin)
SEOMin = LKM.loc[LKM['Col2'] == 'SEO ']
SEOMin['Col4'] = SEOMin['Col4'].transform(lambda x: '{:,.0%}'.format(x))
SEOMin['Col5'] = SEOMin['Col5'].transform(lambda x: '{:,.0%}'.format(x))
SEOMin['Col7'] = SEOMin['Col7'].transform(lambda x: '{:,.0%}'.format(x))
print(SEOMin)

# Page Views (Must select 2nd row)
AppPage = LKM.loc[LKM['Col2'] == 'FOX Weather App1']
AppPage['Col4'] = AppPage['Col4'].transform(lambda x: '{:,.0%}'.format(x))
AppPage['Col5'] = AppPage['Col5'].transform(lambda x: '{:,.0%}'.format(x))
AppPage['Col6'] = AppPage['Col6'].transform(lambda x: '{:,.0%}'.format(x))
AppPage['Col7'] = AppPage['Col7'].transform(lambda x: '{:,.0%}'.format(x))
WebPage = LKM.loc[LKM['Col2'] == 'Foxweather.com1']
WebPage['Col4'] = WebPage['Col4'].transform(lambda x: '{:,.0%}'.format(x))
WebPage['Col5'] = WebPage['Col5'].transform(lambda x: '{:,.0%}'.format(x))
print(WebPage)
SEOPage = LKM.loc[LKM['Col2'] == 'SEO1']
SEOPage['Col4'] = SEOPage['Col4'].transform(lambda x: '{:,.0%}'.format(x))
SEOPage['Col5'] = SEOPage['Col5'].transform(lambda x: '{:,.0%}'.format(x))

# # Unique Devices (Must select 3rd row)
AppUniq = LKM.loc[LKM['Col2'] == 'FOX Weather App Unique devices']
AppUniq['Col4'] = AppUniq['Col4'].transform(lambda x: '{:,.0%}'.format(x))
AppUniq['Col5'] = AppUniq['Col5'].transform(lambda x: '{:,.0%}'.format(x))
WebUniq = LKM.loc[LKM['Col2'] == 'Foxweather.com Unique devices']
WebUniq['Col4'] = WebUniq['Col4'].transform(lambda x: '{:,.0%}'.format(x))
WebUniq['Col5'] = WebUniq['Col5'].transform(lambda x: '{:,.0%}'.format(x))
SEOUniq = LKM.loc[LKM['Col2'] == 'SEO Unique devices']
SEOUniq['Col4'] = SEOUniq['Col4'].transform(lambda x: '{:,.0%}'.format(x))
SEOUniq['Col5'] = SEOUniq['Col5'].transform(lambda x: '{:,.0%}'.format(x))

# # Digital
print(Auto)
TopArt = Auto[['Top Article']].iloc[0]
print(TopArt)
Dig = Auto[['Page Views']].iloc[0]
VideoStarts = Auto[['Video Starts']].iloc[0]
AvgTime = Auto[['Average Time']].iloc[0]

# App
GrossDown = Auto[['Gross Downloads']]
print(GrossDown)
AppPriorM = Auto[['Prior Month1']]
AppPriorM['Prior Month1'] = AppPriorM['Prior Month1'].transform(lambda x: '{:,.2%}'.format(x))
print(AppPriorM['Prior Month1'].iloc[0])
AppPriorY = Auto[['Prior Year1']]
AppPriorY['Prior Year1'] = AppPriorY['Prior Year1'].transform(lambda x: '{:,.2%}'.format(x))
print(AppPriorY['Prior Year1'].iloc[0])

print(AppMin['Col5'].iloc[2])

# Helper Functions
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

# Create Email 
# Note: win32client is not available on MacOS so the following code has not been tested

# ol = win32com.client.Dispatch('Outlook.Application')
# olmailitem = 0x0
# newmail = ol.CreateItem(olmailitem)

# newmail.Subject = 'Fox Weather - Monthly Dashboard' + str(formatted_date)
# newmail.HTMLBody = """
# <h4 style="font-weight: normal;">Hi all - </h4>
# <h4 style="font-weight: normal;">Attached is the Weekly Dashboard for FOX Weather Performance - (""" + str(formatted_date) + """) </h4>
# <p style ="margin-bottom: 0;"><u><b>OPERATIONAL HIGHLIGHTS</b></u></p>
# <ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
#     <li><em> INSERT MANUAL ANALYSIS HERE </em><ul>
#         <li style = "marging-top: 0;">""" + human_format(Total_View['Col3'].iloc[0]) + """ minutes of total view time and  """ + human_format(AMA['Col3'].iloc[0]) + """ Average Minute Audience on the FOX Weather stream driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>
#         </ul>        
#     <li><em> INSERT MANUAL ANALYSIS HERE </em><ul>
#         <li><em> MORE MANUAL ANALYSIS HERE </em></li>
#         </ul> 
#     <li>""" + human_format(WebMin['Col3'].iloc[1]) + """ page views on FOXWeather.com <em> INSERT MANUAL ANALYSIS HERE </em><ul> 
#         <li><em> MORE MANUAL ANALYSIS HERE </em></li>
#         <li><b><u>Top Article Headline: </b></u>""" + TopArt['Top Article'] + """ drove in """ + human_format(Dig.iloc[0]) + """ page views, """ + human_format(VideoStarts.iloc[0]) + """ video starts and """ + human_format(AvgTime.iloc[0]) + """ mins of average time spent driving by <em> INSERT MANUAL ANALYSIS HERE </em></li>
#         </ul> 
#     <li> <em> APPLE NEWS STATS HERE </em> <ul>
#         <li> <em>APPLE NEWS TOP ARTICLE HERE </em></li>
#         </ul>
#     <li> """ + human_format(AppMin['Col3'].iloc[0]) + """ minutes spent on the FOX Weather app down """ + up_down(AppMin['Col5'].iloc[0]) + """ """ + AppMin['Col5'].iloc[0] + """ compared to the prior month and """ + up_down(AppMin['Col6'].iloc[0]) + """ """ + AppMin['Col6'].iloc[0] + """ compared to the 5-month average driven by <em> INSERT MANUAL ANALYSIS HERE </em> <ul>
#         <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
#         </ul>
#     </li>
# </ul>
# <p  style ="margin-bottom: 0;"><u><b>FOX WEATHER STREAM ENGAGEMENT</b></u></p>
# <ol style="list-style-type: lower-alpha; margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;"">
#     <li><u>Unique Viewing Devices:</u> """ + human_format(Unique_Viewing['Col3'].iloc[0]) + """ unique viewing devices""" + up_down(Unique_Viewing['Col4'].iloc[0]) + """ """ + Unique_Viewing['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Unique_Viewing['Col5'].iloc[0]) + """ """ + Unique_Viewing['Col5'].iloc[0] + """ vs. last month &nbsp;</li> 
#     <li><u>Session Frequency:</u> """ + str(round(Session_Freq['Col3'].iloc[0],1)) + """ session frequency """ + up_down(Session_Freq['Col4'].iloc[0]) + """ """ + Session_Freq['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Session_Freq['Col5'].iloc[0]) + """ """ + Session_Freq['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
#     <li><u>Dwell Time (Mins):</u> """ + str(round(Dwell['Col3'].iloc[0],1)) + """ """ + up_down(Dwell['Col4'].iloc[0]) + """ """ + Dwell['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Dwell['Col5'].iloc[0]) + """ """ + Dwell['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
#     <li><u>Total View Time (Mins):</u> """ + human_format(Total_View['Col3'].iloc[0]) + """ """ + up_down(Total_View['Col4'].iloc[0]) + """ """ + Total_View['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Total_View['Col5'].iloc[0]) + """ """ + Total_View['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
# </ol>
# <p  style ="margin-bottom: 0;"><u><b>MINUTES SPENT</b></u></p>
# <ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
#     <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[7]) + """ of total minutes spent and """ + up_down(AppMin['Col4'].iloc[0]) + """ """ + AppMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[0]) + """ """ + AppMin['Col5'].iloc[0] + """ vs. last month</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[8]) + """ of total minutes spent and """ + up_down(WebMin['Col4'].iloc[0]) + """ """ + WebMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[0]) + """ """ + WebMin['Col5'].iloc[0] + """ vs. last month<ul>
#         <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[9]) + """ of total minutes spent and """ + up_down(SEOMin['Col4'].iloc[0]) + """ """ + SEOMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[0]) + """ """ + SEOMin['Col5'].iloc[0] + """ vs. last month</li>
#         </ul>
#     </li>
# </ul>
# <p style ="margin-bottom: 0;"><u><b>PAGE VIEWS</b></u></p>
# <ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
#     <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[11]) + """ of total page views and """ + up_down(AppMin['Col4'].iloc[1]) + """ """ + AppMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[1]) + """ """ +  AppMin['Col5'].iloc[1] + """ vs. last month</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[12]) + """ of total page views and """ + up_down(WebMin['Col4'].iloc[1]) + """ """ + WebMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[1]) + """ """ +  WebMin['Col5'].iloc[1] + """ vs. last month<ul>
#         <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[13]) + """ of total page views and """ + up_down(SEOMin['Col4'].iloc[1]) + """ """ + SEOMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[1]) + """ """ +  SEOMin['Col5'].iloc[1] + """ vs. last month</li>
#         </ul>
#     <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
# </ul>
# <p style ="margin-bottom: 0;"><u><b>UNIQUE DEVICES</b></u></p>
# <ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
#     <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[15]) + """ of total unique devices and """ + up_down(AppMin['Col4'].iloc[2]) + """ """ + AppMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[2]) + """ """ +  AppMin['Col5'].iloc[2] + """ vs. last month</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[16]) + """ of total unique devices and """ + up_down(WebMin['Col4'].iloc[2]) + """ """ + WebMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[2]) + """ """ +  WebMin['Col5'].iloc[2] + """ vs. last month<ul>
#         <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[17]) + """ of total unique devices and """ + up_down(SEOMin['Col4'].iloc[2]) + """ """ + SEOMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[2]) + """ """ +  SEOMin['Col5'].iloc[2] + """ vs. last month</li>
#         </ul>
#     <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
# </ul>
# <p style ="margin-bottom: 0;"><u><b>App Downloads</b></u></p>
# <ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
#  <li>""" + human_format(GrossDown['Gross Downloads'].iloc[0]) + """ gross app downloads in """ + str(formatted_date) + """ """ + up_down(AppPriorM['Prior Month1'].iloc[0]) + """ """ + AppPriorM['Prior Month1'].iloc[0] + """ and were """ + up_down(AppPriorY['Prior Year1'].iloc[0]) + """ """ + AppPriorY['Prior Year1'].iloc[0] + """ compared to the prior year driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>

# </ul>
#   <p style = "margin-bottom:0;">Regards,</p>
#   <p style="margin : 0; padding-top:0;">Sumukh Kamath</p>
# """

# newMail.To = "sumukh.kamath@fox.com"
# newMail.display()
# newMail.send()

# Test HTML Page
Function_Name = open("WEATHERTEST2.html","w")

Function_Name.write("""
<h4 style="font-weight: normal;">Hi all - </h4>
<h4 style="font-weight: normal;">Attached is the Monthly Dashboard for FOX Weather Performance - (""" + str(formatted_date) + """) </h4>
<p style ="margin-bottom: 0;"><u><b>OPERATIONAL HIGHLIGHTS</b></u></p>
<ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
    <li><em> INSERT MANUAL ANALYSIS HERE </em><ul>
        <li style = "marging-top: 0;">""" + human_format(Total_View['Col3'].iloc[0]) + """ minutes of total view time and  """ + human_format(AMA['Col3'].iloc[0]) + """ Average Minute Audience on the FOX Weather stream driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>
        </ul>        
    <li><em> INSERT MANUAL ANALYSIS HERE </em><ul>
        <li><em> MORE MANUAL ANALYSIS HERE </em></li>
        </ul> 
    <li>""" + human_format(WebMin['Col3'].iloc[1]) + """ page views on FOXWeather.com <em> INSERT MANUAL ANALYSIS HERE </em><ul> 
        <li><em> MORE MANUAL ANALYSIS HERE </em></li>
        <li><b><u>Top Article Headline: </u>""" + TopArt['Top Article'] + """</b> drove in """ + human_format(Dig.iloc[0]) + """ page views, """ + human_format(VideoStarts.iloc[0]) + """ video starts and """ + human_format(AvgTime.iloc[0]) + """ mins of average time spent driving by <em> INSERT MANUAL ANALYSIS HERE </em></li>
        </ul> 
    <li> <em> APPLE NEWS STATS HERE </em> <ul>
        <li> <em>APPLE NEWS TOP ARTICLE HERE </em></li>
        </ul>
    <li> """ + human_format(AppMin['Col3'].iloc[0]) + """ minutes spent on the FOX Weather app down """ + up_down(AppMin['Col5'].iloc[0]) + """ """ + AppMin['Col5'].iloc[0] + """ compared to the prior month and """ + up_down(AppMin['Col6'].iloc[0]) + """ """ + AppMin['Col6'].iloc[0] + """ compared to the 5-month average driven by <em> INSERT MANUAL ANALYSIS HERE </em> <ul>
        <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
        </ul>
    </li>
</ul>
<p  style ="margin-bottom: 0;"><u><b>FOX WEATHER STREAM ENGAGEMENT</b></u></p>
<ol style="list-style-type: lower-alpha; margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;"">
    <li><u>Unique Viewing Devices:</u> """ + human_format(Unique_Viewing['Col3'].iloc[0]) + """ unique viewing devices""" + up_down(Unique_Viewing['Col4'].iloc[0]) + """ """ + Unique_Viewing['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Unique_Viewing['Col5'].iloc[0]) + """ """ + Unique_Viewing['Col5'].iloc[0] + """ vs. last month &nbsp;</li> 
    <li><u>Session Frequency:</u> """ + str(round(Session_Freq['Col3'].iloc[0],1)) + """ session frequency """ + up_down(Session_Freq['Col4'].iloc[0]) + """ """ + Session_Freq['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Session_Freq['Col5'].iloc[0]) + """ """ + Session_Freq['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
    <li><u>Dwell Time (Mins):</u> """ + str(round(Dwell['Col3'].iloc[0],1)) + """ """ + up_down(Dwell['Col4'].iloc[0]) + """ """ + Dwell['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Dwell['Col5'].iloc[0]) + """ """ + Dwell['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
    <li><u>Total View Time (Mins):</u> """ + human_format(Total_View['Col3'].iloc[0]) + """ """ + up_down(Total_View['Col4'].iloc[0]) + """ """ + Total_View['Col4'].iloc[0] + """ vs. same month last year and """ + up_down(Total_View['Col5'].iloc[0]) + """ """ + Total_View['Col5'].iloc[0] + """ vs. last month &nbsp;</li>
</ol>
<p  style ="margin-bottom: 0;"><u><b>MINUTES SPENT</b></u></p>
<ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
    <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[7]) + """ of total minutes spent and """ + up_down(AppMin['Col4'].iloc[0]) + """ """ + AppMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[0]) + """ """ + AppMin['Col5'].iloc[0] + """ vs. last month</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[8]) + """ of total minutes spent and """ + up_down(WebMin['Col4'].iloc[0]) + """ """ + WebMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[0]) + """ """ + WebMin['Col5'].iloc[0] + """ vs. last month<ul>
        <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[0]) + """, """ + (PercentOfTotal.iloc[9]) + """ of total minutes spent and """ + up_down(SEOMin['Col4'].iloc[0]) + """ """ + SEOMin['Col4'].iloc[0] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[0]) + """ """ + SEOMin['Col5'].iloc[0] + """ vs. last month</li>
        </ul>
    </li>
</ul>
<p style ="margin-bottom: 0;"><u><b>PAGE VIEWS</b></u></p>
<ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
    <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[11]) + """ of total page views and """ + up_down(AppMin['Col4'].iloc[1]) + """ """ + AppMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[1]) + """ """ +  AppMin['Col5'].iloc[1] + """ vs. last month</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[12]) + """ of total page views and """ + up_down(WebMin['Col4'].iloc[1]) + """ """ + WebMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[1]) + """ """ +  WebMin['Col5'].iloc[1] + """ vs. last month<ul>
        <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[1]) + """, """ + str(PercentOfTotal.iloc[13]) + """ of total page views and """ + up_down(SEOMin['Col4'].iloc[1]) + """ """ + SEOMin['Col4'].iloc[1] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[1]) + """ """ +  SEOMin['Col5'].iloc[1] + """ vs. last month</li>
        </ul>
    <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
</ul>
<p style ="margin-bottom: 0;"><u><b>UNIQUE DEVICES</b></u></p>
<ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
    <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[15]) + """ of total unique devices and """ + up_down(AppMin['Col4'].iloc[2]) + """ """ + AppMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(AppMin['Col5'].iloc[2]) + """ """ +  AppMin['Col5'].iloc[2] + """ vs. last month</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[16]) + """ of total unique devices and """ + up_down(WebMin['Col4'].iloc[2]) + """ """ + WebMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(WebMin['Col5'].iloc[2]) + """ """ +  WebMin['Col5'].iloc[2] + """ vs. last month<ul>
        <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col3'].iloc[2]) + """, """ + str(PercentOfTotal.iloc[17]) + """ of total unique devices and """ + up_down(SEOMin['Col4'].iloc[2]) + """ """ + SEOMin['Col4'].iloc[2] + """ vs. the same month last year and """ + up_down(SEOMin['Col5'].iloc[2]) + """ """ +  SEOMin['Col5'].iloc[2] + """ vs. last month</li>
        </ul>
    <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
</ul>
<p style ="margin-bottom: 0;"><u><b>App Downloads</b></u></p>
<ul style="margin-top: 0; margin-bottom: 0; padding-top: 0; padding-bottom: 0;">
 <li>""" + human_format(GrossDown['Gross Downloads'].iloc[0]) + """ gross app downloads in """ + str(formatted_date) + """ """ + up_down(AppPriorM['Prior Month1'].iloc[0]) + """ """ + AppPriorM['Prior Month1'].iloc[0] + """ and were """ + up_down(AppPriorY['Prior Year1'].iloc[0]) + """ """ + AppPriorY['Prior Year1'].iloc[0] + """ compared to the prior year driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>

</ul>
  <p style = "margin-bottom:0;">Regards,</p>
  <p style="margin : 0; padding-top:0;">Sumukh Kamath</p>
""")
Function_Name.close()
