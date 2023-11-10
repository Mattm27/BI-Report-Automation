import pandas as pd
import numpy as np
import datetime as dt

pd.set_option('display.max_colwidth', None)

def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])

# Pull in sheets
Summary = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230723_Fox_Weather_Weekly_dashboard.xlsx', sheet_name='Summary - Weekly')
LKM = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230723_Fox_Weather_Weekly_dashboard.xlsx', sheet_name='LKM_data_wk')
Auto = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230723_Fox_Weather_Weekly_dashboard.xlsx', sheet_name='Email Auto')
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
Date = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20230723_Fox_Weather_Weekly_dashboard.xlsx', sheet_name='Summary - Weekly', header=None)
Date1= Date[37].iloc[0]

# Stream Engagement
Unique_Viewing = LKM.loc[LKM['Col1'] == 'Unique Viewers']
Unique_Viewing['Col4'] = Unique_Viewing['Col4'].transform(lambda x: '{:,.2%}'.format(x))
Unique_Viewing['Col7'] = Unique_Viewing['Col7'].transform(lambda x: '{:,.2%}'.format(x))
Unique_Viewing['Col8'] = Unique_Viewing['Col8'].transform(lambda x: '{:,.2%}'.format(x))
Session_Freq = LKM.loc[LKM['Col1'] == 'Session Frequency']
Session_Freq['Col4'] = Session_Freq['Col4'].transform(lambda x: '{:,.2%}'.format(x))
Session_Freq['Col7'] = Session_Freq['Col7'].transform(lambda x: '{:,.2%}'.format(x))
Session_Freq['Col8'] = Session_Freq['Col8'].transform(lambda x: '{:,.2%}'.format(x))
Dwell = LKM.loc[LKM['Col1'] == 'Dwell Time (Mins)']
Dwell['Col4'] = Dwell['Col4'].transform(lambda x: '{:,.2%}'.format(x))
Dwell['Col7'] = Dwell['Col7'].transform(lambda x: '{:,.2%}'.format(x))
Dwell['Col8'] = Dwell['Col8'].transform(lambda x: '{:,.2%}'.format(x))
Total_View = LKM.loc[LKM['Col1'] == 'Total View Time (Mins)']
Total_View['Col4'] = Total_View['Col4'].transform(lambda x: '{:,.2%}'.format(x))
Total_View['Col7'] = Total_View['Col7'].transform(lambda x: '{:,.2%}'.format(x))
Total_View['Col8'] = Total_View['Col8'].transform(lambda x: '{:,.2%}'.format(x))

# Minutes Spent
AppMin = LKM.loc[LKM['Col1'] == 'FOX Weather Time spent']
AppMin['Col4'] = AppMin['Col4'].transform(lambda x: '{:,.2%}'.format(x))
AppMin['Col7'] = AppMin['Col7'].transform(lambda x: '{:,.2%}'.format(x))
AppMin['Col8'] = AppMin['Col8'].transform(lambda x: '{:,.2%}'.format(x))
print(AppMin)
WebMin = LKM.loc[LKM['Col1'] == 'Foxweather.com Time Spent']
WebMin['Col4'] = WebMin['Col4'].transform(lambda x: '{:,.2%}'.format(x))
WebMin['Col7'] = WebMin['Col7'].transform(lambda x: '{:,.2%}'.format(x))
WebMin['Col8'] = WebMin['Col8'].transform(lambda x: '{:,.2%}'.format(x))
SEOMin = LKM.loc[LKM['Col1'] == 'SEO Time Spent']
SEOMin['Col4'] = SEOMin['Col4'].transform(lambda x: '{:,.2%}'.format(x))
SEOMin['Col7'] = SEOMin['Col7'].transform(lambda x: '{:,.2%}'.format(x))
SEOMin['Col8'] = SEOMin['Col8'].transform(lambda x: '{:,.2%}'.format(x))

# Page Views
AppPage = LKM.loc[LKM['Col1'] == 'FOX Weather App Page Views']
AppPage['Col4'] = AppPage['Col4'].transform(lambda x: '{:,.2%}'.format(x))
AppPage['Col7'] = AppPage['Col7'].transform(lambda x: '{:,.2%}'.format(x))
AppPage['Col8'] = AppPage['Col8'].transform(lambda x: '{:,.2%}'.format(x))
WebPage = LKM.loc[LKM['Col1'] == 'Foxweather.com Page views']
WebPage['Col4'] = WebPage['Col4'].transform(lambda x: '{:,.2%}'.format(x))
WebPage['Col7'] = WebPage['Col7'].transform(lambda x: '{:,.2%}'.format(x))
WebPage['Col8'] = WebPage['Col8'].transform(lambda x: '{:,.2%}'.format(x))
SEOPage = LKM.loc[LKM['Col1'] == 'SEO Page Views']
SEOPage['Col4'] = SEOPage['Col4'].transform(lambda x: '{:,.2%}'.format(x))
SEOPage['Col7'] = SEOPage['Col7'].transform(lambda x: '{:,.2%}'.format(x))
SEOPage['Col8'] = SEOPage['Col8'].transform(lambda x: '{:,.2%}'.format(x))

# Unique Devices
AppUniq = LKM.loc[LKM['Col1'] == 'FOX Weather App Unique devices']
AppUniq['Col4'] = AppUniq['Col4'].transform(lambda x: '{:,.2%}'.format(x))
AppUniq['Col7'] = AppUniq['Col7'].transform(lambda x: '{:,.2%}'.format(x))
AppUniq['Col8'] = AppUniq['Col8'].transform(lambda x: '{:,.2%}'.format(x))
WebUniq = LKM.loc[LKM['Col1'] == 'Foxweather.com Unique devices']
WebUniq['Col4'] = WebUniq['Col4'].transform(lambda x: '{:,.2%}'.format(x))
WebUniq['Col7'] = WebUniq['Col7'].transform(lambda x: '{:,.2%}'.format(x))
WebUniq['Col8'] = WebUniq['Col8'].transform(lambda x: '{:,.2%}'.format(x))
SEOUniq = LKM.loc[LKM['Col1'] == 'SEO Unique devices']
SEOUniq['Col4'] = SEOUniq['Col4'].transform(lambda x: '{:,.2%}'.format(x))
SEOUniq['Col7'] = SEOUniq['Col7'].transform(lambda x: '{:,.2%}'.format(x))
SEOUniq['Col8'] = SEOUniq['Col8'].transform(lambda x: '{:,.2%}'.format(x))

# Stream 
StreamTotal = Auto[['Total Minutes']]
AvgMin = Auto[['AvgMin']]
PriorY = Auto[['Prior Year']]
PriorY['Prior Year'] = PriorY['Prior Year'].transform(lambda x: '{:,.2%}'.format(x))
PriorW = Auto[['Prior Week']]
PriorW['Prior Week'] = PriorW['Prior Week'].transform(lambda x: '{:,.2%}'.format(x))

# Digital
TopArt = Auto[['Top Article']]
Dig = Auto[['Page Views']]
VideoStarts = Auto[['Video Starts']]
AvgTime = Auto[['Average Time']]

# App
GrossDown = Auto[['Gross Downloads']]
AppPriorW = Auto[['Prior Week1']]
AppPriorW['Prior Week1'] = AppPriorW['Prior Week1'].transform(lambda x: '{:,.2%}'.format(x))

# Helper Functions
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

# Create Email 
# ** Note: win32client is not available on MacOS so the following code has not been tested

# ol = win32com.client.Dispatch('Outlook.Application')
# olmailitem = 0x0
# newmail = ol.CreateItem(olmailitem)

# newmail.Subject = 'Fox Weather - Weekly Dashboard' + str(Date)
# newmail.HTMLBody = """<h4 style="font-weight: normal;">Hi all - </h4>
# <h4 style="font-weight: normal;">Attached is the Weekly Dashboard for FOX Weather Performance - (""" + str(Date1) + """) </h4>
# <p><u><b>OPERATIONAL HIGHLIGHTS</b></u></p>
# <ul>
#     <li> """ + human_format(StreamTotal['Total Minutes'].iloc[0]) + """ minutes of total watch time and """ + human_format(AvgMin['AvgMin'].iloc[0]) + """ Average Minute Audience on the FOX Weather stream, """ + up_down(PriorY['Prior Year'].iloc[0]) + """ """ + PriorY['Prior Year'].iloc[0] + """ compared to the same week last year driven by <em> INSERT MANUAL ANALYSIS HERE </em><ul>
#         <li> Engagement on the FOX Weather stream was """ + up_down(PriorW['Prior Week'].iloc[0]) + """ """ + PriorW['Prior Week'].iloc[0] + """ compared to the prior week driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>
#                 <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
#         </ul>        
#     <li> """ + human_format(WebPage['Col2'].iloc[0]) + """ """ + up_down(WebPage['Col7'].iloc[0]) + """ """ + WebPage['Col7'].iloc[0] + """ compared to the same week last year driven by <em> INSERT MANUAL ANALYSIS HERE </em><ul>
#         <li><b><u> Top Article Headline: </u>" """ + TopArt['Top Article'].iloc[0] + """ "</b> drove """ + human_format(Dig['Page Views'].iloc[0]) + """ page views, """ + human_format(VideoStarts['Video Starts'].iloc[0]) + """ video starts, and """ + str(round(AvgTime['Average Time'].iloc[0],1)) + """ of average time spent <em> INSERT MANUAL ANALYSIS </em> </li>
#         </ul>
#     <li> <em> APPLE NEWS STATS HERE </em> <ul>
#         <li> <em>APPLE NEWS TOP ARTICLE HERE </em></li>
#         </ul>
#     <li> """ + human_format(GrossDown['Gross Downloads'].iloc[0]) + """ gross app downloads last week """ + up_down(AppPriorW['Prior Week1'].iloc[0]) + """ """ + AppPriorW['Prior Week1'].iloc[0] + """ compared to the prior week driven by <em> INSERT MANUAL ANALYSIS HERE </em> <ul>
#         <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
#         </ul>
#     </li>
# </ul>
# <p><u><b>FOX WEATHER STREAM ENGAGEMENT</b></u></p>
# <ol style="list-style-type: lower-alpha;">
#     <li><u>Unique Viewing Devices:</u> """ + human_format(Unique_Viewing['Col2'].iloc[0]) + """ """ + up_down(Unique_Viewing['Col7'].iloc[0]) + """ """ + Unique_Viewing['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Unique_Viewing['Col8'].iloc[0]) + """ """ + Unique_Viewing['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li> 
#     <li><u>Session Frequency:</u> """ + str(round(Session_Freq['Col2'].iloc[0],1)) + """ """ + up_down(Session_Freq['Col7'].iloc[0]) + """ """ + Session_Freq['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Session_Freq['Col8'].iloc[0]) + """ """ + Session_Freq['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
#     <li><u>Dwell Time (Mins):</u> """ + str(round(Dwell['Col2'].iloc[0],1)) + """ """ + up_down(Dwell['Col7'].iloc[0]) + """ """ + Dwell['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Dwell['Col8'].iloc[0]) + """ """ + Dwell['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
#     <li><u>Total View Time (Mins):</u> """ + human_format(Total_View['Col2'].iloc[0]) + """ """ + up_down(Total_View['Col7'].iloc[0]) + """ """ + Total_View['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Total_View['Col8'].iloc[0]) + """ """ + Total_View['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
# </ol>
# <p><u><b>MINUTES SPENT</b></u></p>
# <ul>
#     <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col2'].iloc[0]) + """, """ + AppMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppMin['Col7'].iloc[0]) + """ """ + AppMin['Col7'].iloc[0] + """ vs. the same week last year</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col2'].iloc[0]) + """, """ + WebMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebMin['Col7'].iloc[0]) + """ """ + WebMin['Col7'].iloc[0] + """ vs. the same week last year<ul>
#         <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col2'].iloc[0]) + """, """ + SEOMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOMin['Col7'].iloc[0]) + """ """ + SEOMin['Col7'].iloc[0] + """ vs. the same week last year</li>
#         </ul>
#     </li>    
# </ul>
# <p><u><b>PAGE VIEWS</b></u></p>
# <ul>
#     <li><u>FOX Weather App:</u> """ + human_format(AppPage['Col2'].iloc[0]) + """, """ + AppPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppPage['Col7'].iloc[0]) + """ """ + AppPage['Col7'].iloc[0] + """ vs. the same week last year</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebPage['Col2'].iloc[0]) + """, """ + WebPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebPage['Col7'].iloc[0]) + """ """ + WebPage['Col7'].iloc[0] + """ vs. the same week last year<ul>
#         <li><u>SEO REFERRALS:</u> """ + human_format(SEOPage['Col2'].iloc[0]) + """, """ + SEOPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOPage['Col7'].iloc[0]) + """ """ + SEOPage['Col7'].iloc[0] + """ vs. the same week last year</li>
#         </ul>
#     <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
# </ul>
# <p><u><b>UNIQUE DEVICES</b></u></p>
# <ul>
#     <li><u>FOX Weather App:</u> """ + human_format(AppUniq['Col2'].iloc[0]) + """, """ + AppUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppUniq['Col7'].iloc[0]) + """ """ + AppUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
#     <li><u>FOXWeather.com:</u> """ + human_format(WebUniq['Col2'].iloc[0]) + """, """ + WebUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebUniq['Col7'].iloc[0]) + """ """ + WebUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
#     <li><u>SEO REFERRALS:</u> """ + human_format(SEOUniq['Col2'].iloc[0]) + """, """ + SEOUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOUniq['Col7'].iloc[0]) + """ """ + SEOUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
#     <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>        
# </ul>

# <p style = "margin-bottom:0;">Regards,</p>
# <p style="margin : 0; padding-top:0;">Sumukh Kamath</p>
# """

# newMail.To = "sumukh.kamath@fox.com"
# newMail.display()
# newMail.send()

# Test HTML Page
Function_Name = open("WEATHERTEST1.html","w")

Function_Name.write("""
<h4 style="font-weight: normal;">Hi all - </h4>
<h4 style="font-weight: normal;">Attached is the Weekly Dashboard for FOX Weather Performance - (""" + str(Date1) + """) </h4>
<p><u><b>OPERATIONAL HIGHLIGHTS</b></u></p>
<ul>
    <li> """ + human_format(StreamTotal['Total Minutes'].iloc[0]) + """ minutes of total watch time and """ + human_format(AvgMin['AvgMin'].iloc[0]) + """ Average Minute Audience on the FOX Weather stream, """ + up_down(PriorY['Prior Year'].iloc[0]) + """ """ + PriorY['Prior Year'].iloc[0] + """ compared to the same week last year driven by <em> INSERT MANUAL ANALYSIS HERE </em><ul>
        <li> Engagement on the FOX Weather stream was """ + up_down(PriorW['Prior Week'].iloc[0]) + """ """ + PriorW['Prior Week'].iloc[0] + """ compared to the prior week driven by <em> INSERT MANUAL ANALYSIS HERE </em></li>
                <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
        </ul>        
    <li> """ + human_format(WebPage['Col2'].iloc[0]) + """ """ + up_down(WebPage['Col7'].iloc[0]) + """ """ + WebPage['Col7'].iloc[0] + """ compared to the same week last year driven by <em> INSERT MANUAL ANALYSIS HERE </em><ul>
        <li><b><u> Top Article Headline: </u>" """ + TopArt['Top Article'].iloc[0] + """ "</b> drove """ + human_format(Dig['Page Views'].iloc[0]) + """ page views, """ + human_format(VideoStarts['Video Starts'].iloc[0]) + """ video starts, and """ + str(round(AvgTime['Average Time'].iloc[0],1)) + """ of average time spent <em> INSERT MANUAL ANALYSIS </em> </li>
        </ul>
    <li> <em> APPLE NEWS STATS HERE </em> <ul>
        <li> <em>APPLE NEWS TOP ARTICLE HERE </em></li>
        </ul>
    <li> """ + human_format(GrossDown['Gross Downloads'].iloc[0]) + """ gross app downloads last week """ + up_down(AppPriorW['Prior Week1'].iloc[0]) + """ """ + AppPriorW['Prior Week1'].iloc[0] + """ compared to the prior week driven by <em> INSERT MANUAL ANALYSIS HERE </em> <ul>
        <li> <em> INSERT MANUAL ANALYSIS HERE IF NEEDED </em> </li>
        </ul>
    </li>
</ul>
<p><u><b>FOX WEATHER STREAM ENGAGEMENT</b></u></p>
<ol style="list-style-type: lower-alpha;">
    <li><u>Unique Viewing Devices:</u> """ + human_format(Unique_Viewing['Col2'].iloc[0]) + """ """ + up_down(Unique_Viewing['Col7'].iloc[0]) + """ """ + Unique_Viewing['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Unique_Viewing['Col8'].iloc[0]) + """ """ + Unique_Viewing['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li> 
    <li><u>Session Frequency:</u> """ + str(round(Session_Freq['Col2'].iloc[0],1)) + """ """ + up_down(Session_Freq['Col7'].iloc[0]) + """ """ + Session_Freq['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Session_Freq['Col8'].iloc[0]) + """ """ + Session_Freq['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
    <li><u>Dwell Time (Mins):</u> """ + str(round(Dwell['Col2'].iloc[0],1)) + """ """ + up_down(Dwell['Col7'].iloc[0]) + """ """ + Dwell['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Dwell['Col8'].iloc[0]) + """ """ + Dwell['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
    <li><u>Total View Time (Mins):</u> """ + human_format(Total_View['Col2'].iloc[0]) + """ """ + up_down(Total_View['Col7'].iloc[0]) + """ """ + Total_View['Col7'].iloc[0] + """ vs. same week last year and """ + up_down(Total_View['Col8'].iloc[0]) + """ """ + Total_View['Col8'].iloc[0] + """ vs. 5-week average &nbsp;</li>
</ol>
<p><u><b>MINUTES SPENT</b></u></p>
<ul>
    <li><u>FOX Weather App:</u> """ + human_format(AppMin['Col2'].iloc[0]) + """, """ + AppMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppMin['Col7'].iloc[0]) + """ """ + AppMin['Col7'].iloc[0] + """ vs. the same week last year</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebMin['Col2'].iloc[0]) + """, """ + WebMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebMin['Col7'].iloc[0]) + """ """ + WebMin['Col7'].iloc[0] + """ vs. the same week last year<ul>
        <li><u>SEO REFERRALS:</u> """ + human_format(SEOMin['Col2'].iloc[0]) + """, """ + SEOMin['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOMin['Col7'].iloc[0]) + """ """ + SEOMin['Col7'].iloc[0] + """ vs. the same week last year</li>
        </ul>
    </li>    
</ul>
<p><u><b>PAGE VIEWS</b></u></p>
<ul>
    <li><u>FOX Weather App:</u> """ + human_format(AppPage['Col2'].iloc[0]) + """, """ + AppPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppPage['Col7'].iloc[0]) + """ """ + AppPage['Col7'].iloc[0] + """ vs. the same week last year</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebPage['Col2'].iloc[0]) + """, """ + WebPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebPage['Col7'].iloc[0]) + """ """ + WebPage['Col7'].iloc[0] + """ vs. the same week last year<ul>
        <li><u>SEO REFERRALS:</u> """ + human_format(SEOPage['Col2'].iloc[0]) + """, """ + SEOPage['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOPage['Col7'].iloc[0]) + """ """ + SEOPage['Col7'].iloc[0] + """ vs. the same week last year</li>
        </ul>
    <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>
</ul>
<p><u><b>UNIQUE DEVICES</b></u></p>
<ul>
    <li><u>FOX Weather App:</u> """ + human_format(AppUniq['Col2'].iloc[0]) + """, """ + AppUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(AppUniq['Col7'].iloc[0]) + """ """ + AppUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
    <li><u>FOXWeather.com:</u> """ + human_format(WebUniq['Col2'].iloc[0]) + """, """ + WebUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(WebUniq['Col7'].iloc[0]) + """ """ + WebUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
    <li><u>SEO REFERRALS:</u> """ + human_format(SEOUniq['Col2'].iloc[0]) + """, """ + SEOUniq['Col4'].iloc[0] + """ of total minutes spent and """ + up_down(SEOUniq['Col7'].iloc[0]) + """ """ + SEOUniq['Col7'].iloc[0] + """ vs. the same week last year</li>
    <li><u>Apple News:</u><em> INSERT APPLE NEWS STATS HERE </em> </li>        
</ul>

<p style = "margin-bottom:0;">Regards,</p>
<p style="margin : 0; padding-top:0;">Sumukh Kamath</p>
""")
Function_Name.close()







































