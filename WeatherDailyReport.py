import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import win32com.client

pd.set_option('display.max_colwidth', None)

def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])

def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

# # Pull in sheets
# Summary = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='Summary - Daily')
# LKM = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='LKM_data_day')
# # Auto = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='Email Auto')

# # Rename columns
# LKM = LKM.rename(columns={LKM.columns[0]: 'Col1'})
# LKM = LKM.rename(columns={LKM.columns[1]: 'Col2'})
# LKM = LKM.rename(columns={LKM.columns[2]: 'Col3'})
# LKM = LKM.rename(columns={LKM.columns[3]: 'Col4'})
# LKM = LKM.rename(columns={LKM.columns[4]: 'Col5'})
# LKM = LKM.rename(columns={LKM.columns[5]: 'Col6'})
# LKM = LKM.rename(columns={LKM.columns[6]: 'Col7'})
# LKM = LKM.rename(columns={LKM.columns[7]: 'Col8'})
# LKM = LKM.rename(columns={LKM.columns[8]: 'Col9'})

# Summary = Summary.rename(columns={Summary.columns[0]: 'Col1'})
# Summary = Summary.rename(columns={Summary.columns[1]: 'Col2'})
# Summary = Summary.rename(columns={Summary.columns[2]: 'Col3'})
# Summary = Summary.rename(columns={Summary.columns[3]: 'Col4'})
# Summary = Summary.rename(columns={Summary.columns[4]: 'Col5'})
# Summary = Summary.rename(columns={Summary.columns[5]: 'Col6'})
# Summary = Summary.rename(columns={Summary.columns[6]: 'Col7'})
# Summary = Summary.rename(columns={Summary.columns[7]: 'Col8'})
# Summary = Summary.rename(columns={Summary.columns[8]: 'Col9'})
# Summary = Summary.rename(columns={Summary.columns[9]: 'Col10'})
# Summary = Summary.rename(columns={Summary.columns[10]: 'Col11'})
# Summary = Summary.rename(columns={Summary.columns[11]: 'Col12'})
# Summary = Summary.rename(columns={Summary.columns[12]: 'Col13'})
# Summary = Summary.rename(columns={Summary.columns[13]: 'Col14'})
# Summary = Summary.rename(columns={Summary.columns[14]: 'Col15'})
# Summary = Summary.rename(columns={Summary.columns[15]: 'Col16'})
# Summary = Summary.rename(columns={Summary.columns[16]: 'Col17'})
# Summary = Summary.rename(columns={Summary.columns[17]: 'Col18'})
# Summary = Summary.rename(columns={Summary.columns[18]: 'Col19'})
# Summary = Summary.rename(columns={Summary.columns[19]: 'Col20'})

# Extract Date
Date = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='Summary - Daily', header=None)
CurrDate = Date[18].iloc[0]
print(CurrDate)

# Format date
stringdate = str(CurrDate)
date_obj = datetime.strptime(stringdate,"%Y-%m-%d %H:%M:%S")
formatted_date = date_obj.strftime("%m/%d")
print(formatted_date)

# Date 3 days ago
date_obj1 = datetime.strptime(formatted_date, "%m/%d")

# Calculate three days ago
three_days_ago = date_obj - timedelta(days=3)

# Format the result as "month/day"
formatted_date3 = three_days_ago.strftime("%m/%d")

print(formatted_date3)

# Date 2 days ago
date_obj2 = datetime.strptime(formatted_date, "%m/%d")

# Calculate three days ago
two_days_ago = date_obj - timedelta(days=2)

# Format the result as "month/day"
formatted_date2 = two_days_ago.strftime("%m/%d")

print(formatted_date2)

AvgMinAud = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='LKM_data_day', skiprows=16)
AvgMinAud3 = AvgMinAud['AMA'].iloc[0] # 3 days ago 
AvgMinAud2 = AvgMinAud['AMA'].iloc[3] # 2 days ago
AvgMinAud1 = AvgMinAud['AMA'].iloc[6] # yesterday
AvgMinAud3PercWeekDay = AvgMinAud['AMA'].iloc[2] # 3 days ago (Will be friday when used)
AvgMinAud3PercWeekDay = AvgMinAud3PercWeekDay * 100
AvgMinAud3PercWeekDay = '{:.2%}'.format(AvgMinAud3PercWeekDay)
AvgMinAud2PercWeekEnd = AvgMinAud['AMA'].iloc[5] # 2 days ago (Will be Saturday when used)
AvgMinAud2PercWeekEnd = AvgMinAud2PercWeekEnd * 100
AvgMinAud2PercWeekEnd = '{:.2%}'.format(AvgMinAud2PercWeekEnd)
AvgMinAud1PercWeekEnd = AvgMinAud['AMA'].iloc[8] # Yesterday comp to weekend average
AvgMinAud1PercWeekEnd = AvgMinAud1PercWeekEnd * 100
AvgMinAud1PercWeekEnd = '{:.2%}'.format(AvgMinAud1PercWeekEnd)
AvgMinAud1PercWeekDay = AvgMinAud['AMA'].iloc[9] # Yesterday comp to weekday average
AvgMinAud1PercWeekDay = AvgMinAud1PercWeekDay * 100
AvgMinAud1PercWeekDay = '{:.2%}'.format(AvgMinAud1PercWeekDay)

TotViewTim = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='LKM_data_day', skiprows=16)
TotViewTim3 = TotViewTim['Minutes Watched'].iloc[0] # 3 days ago 
TotViewTim2 = TotViewTim['Minutes Watched'].iloc[3] # 2 days ago
TotViewTim1 = TotViewTim['Minutes Watched'].iloc[6] # yesterday
TotViewTim3PercWeekDay = TotViewTim['Minutes Watched'].iloc[2] # 3 days ago (Will be friday when used)
TotViewTim3PercWeekDay = TotViewTim3PercWeekDay * 100
TotViewTim3PercWeekDay = '{:.2%}'.format(TotViewTim3PercWeekDay)
TotViewTim2PercWeekEnd = TotViewTim['Minutes Watched'].iloc[5] # 2 days ago (Will be Saturday when used)
TotViewTim2PercWeekEnd = TotViewTim2PercWeekEnd * 100
TotViewTim2PercWeekEnd = '{:.2%}'.format(TotViewTim2PercWeekEnd)
TotViewTim1PercWeekEnd = TotViewTim['Minutes Watched'].iloc[8] # Yesterday comp to weekend average
TotViewTim1PercWeekEnd = TotViewTim1PercWeekEnd * 100
TotViewTim1PercWeekEnd = '{:.2%}'.format(TotViewTim1PercWeekEnd)
TotViewTim1PercWeekDay = TotViewTim['Minutes Watched'].iloc[9] # Yesterday comp to weekday average
TotViewTim1PercWeekDay = TotViewTim1PercWeekDay * 100
TotViewTim1PercWeekDay = '{:.2%}'.format(TotViewTim1PercWeekDay)

UniqDev = pd.read_excel(r'/Users/mattmay/Desktop/Fox BI Summer 2023/20231009_Fox_Weather_Daily_dashboard.xlsx', sheet_name='LKM_data_day', skiprows=16)
UniqDev3 = UniqDev['Unique Viewing Devices'].iloc[0] # 3 days ago 
UniqDev2 = UniqDev['Unique Viewing Devices'].iloc[3] # 2 days ago
UniqDev1 = UniqDev['Unique Viewing Devices'].iloc[6] # yesterday
UniqDev3PercWeekDay = UniqDev['Unique Viewing Devices'].iloc[2] # 3 days ago (Will be friday when used)
UniqDev3PercWeekDay = UniqDev3PercWeekDay * 100
UniqDev3PercWeekDay = '{:.2%}'.format(UniqDev3PercWeekDay)
UniqDev2PercWeekEnd = UniqDev['Unique Viewing Devices'].iloc[5] # 2 days ago (Will be Saturday when used)
UniqDev2PercWeekEnd = UniqDev2PercWeekEnd * 100
UniqDev2PercWeekEnd = '{:.2%}'.format(UniqDev2PercWeekEnd)
UniqDev1PercWeekEnd = UniqDev['Unique Viewing Devices'].iloc[8] # Yesterday comp to weekend average
UniqDev1PercWeekEnd = UniqDev1PercWeekEnd * 100
UniqDev1PercWeekEnd = '{:.2%}'.format(UniqDev1PercWeekEnd)
UniqDev1PercWeekDay = UniqDev['Unique Viewing Devices'].iloc[9] # Yesterday comp to weekday average
UniqDev1PercWeekDay = UniqDev1PercWeekDay * 100
UniqDev1PercWeekDay = '{:.2%}'.format(UniqDev1PercWeekDay)

# Determine what day of the week it is (on monday send out weekend report)
if datetime.weekday(CurrDate) == 0:
    # print weekend format
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'FOX Weather - Daily Dashboard'
    newmail.HTMLBody = """
        <h4 style="font-weight: normal;">Hi all - </h4>
        <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Friday, Saturday, & Sunday. Let us know if you have any questions or feedback. </h4>
        <p><u><b>Key Performance Indicators</b></u></p>
        <p><b><mark>""" + formatted_date3 + """ Friday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud3) + """ """ + up_down(AvgMinAud3PercWeekDay) + """ """ + str(AvgMinAud3PercWeekDay) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim3) + """ """ + up_down(TotViewTim3PercWeekDay) + """ """ + str(TotViewTim3PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev3) + """ """ + up_down(UniqDev3PercWeekDay) + """ """ + str(UniqDev3PercWeekDay) + """ vs. previous year weekends </li>
        </ul>
        <p><b><mark>""" + formatted_date2 + """ Saturday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud2) + """ """ + up_down(AvgMinAud2PercWeekEnd) + """ """ + str(AvgMinAud2PercWeekEnd) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim2) + """ """ + up_down(TotViewTim2PercWeekEnd) + """ """ + str(TotViewTim2PercWeekEnd) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev2) + """ """ + up_down(UniqDev2PercWeekEnd) + """ """ + str(UniqDev2PercWeekEnd) + """ vs. previous year weekends </li>
        </ul>
        <p><b><mark>""" + formatted_date + """ Sunday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekEnd) + """ """ + str(AvgMinAud1PercWeekEnd) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekEnd) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekEnd) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
        </ul>"""
    newMail.To = "sumukh.kamath@fox.com"
    newMail.display()


elif datetime.weekday(CurrDate) == 1:
    # Print weekday format
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'FOX Weather - Daily Dashboard'
    newmail.HTMLBody = """
        <h4 style="font-weight: normal;">Hi all - </h4>
        <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Monday. Let us know if you have any questions or feedback. </h4>
        <p><u><b>Key Performance Indicators</b></u></p>
        <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
        </ul>"""
    newMail.To = "sumukh.kamath@fox.com"
    newMail.display()

elif datetime.weekday(CurrDate) == 2:
    # Print weekday format
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'FOX Weather - Daily Dashboard'
    newmail.HTMLBody = """
        <h4 style="font-weight: normal;">Hi all - </h4>
        <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Tuesday. Let us know if you have any questions or feedback. </h4>
        <p><u><b>Key Performance Indicators</b></u></p>
        <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
        </ul>
        """
    newMail.To = "sumukh.kamath@fox.com"
    newMail.display()

elif datetime.weekday(CurrDate) == 3:
    # Print weekday format
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'FOX Weather - Daily Dashboard'
    newmail.HTMLBody = """
        <h4 style="font-weight: normal;">Hi all - </h4>
        <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Wednesday. Let us know if you have any questions or feedback. </h4>
        <p><u><b>Key Performance Indicators</b></u></p>
        <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
        </ul>
        """
    newMail.To = "sumukh.kamath@fox.com"
    newMail.display()

elif datetime.weekday(CurrDate) == 4:
    # Print weekday format
    ol = win32com.client.Dispatch('Outlook.Application')
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'FOX Weather - Daily Dashboard'
    newmail.HTMLBody = """
        <h4 style="font-weight: normal;">Hi all - </h4>
        <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Thursday. Let us know if you have any questions or feedback. </h4>
        <p><u><b>Key Performance Indicators</b></u></p>
        <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
        <ul>
            <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
            <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
            <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
        </ul>"""
    newMail.To = "sumukh.kamath@fox.com"
    newMail.display()

# 
# 
# 
# 
# 
# 
# HTML TEST:

# Determine what day of the week it is (on monday send out weekend report)
# if datetime.weekday(CurrDate) == 0:
#     # Print weekend format
#     # Test HTML Page
#     Function_Name = open("WEATHERTESTDAILY.html","w")
#     Function_Name.write("""
#         <h4 style="font-weight: normal;">Hi all - </h4>
#         <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Friday, Saturday, & Sunday. Let us know if you have any questions or feedback. </h4>
#         <p><u><b>Key Performance Indicators</b></u></p>
#         <p><b><mark>""" + formatted_date3 + """ Friday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud3) + """ """ + up_down(AvgMinAud3PercWeekDay) + """ """ + str(AvgMinAud3PercWeekDay) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim3) + """ """ + up_down(TotViewTim3PercWeekDay) + """ """ + str(TotViewTim3PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev3) + """ """ + up_down(UniqDev3PercWeekDay) + """ """ + str(UniqDev3PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
#         <p><b><mark>""" + formatted_date2 + """ Saturday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud2) + """ """ + up_down(AvgMinAud2PercWeekEnd) + """ """ + str(AvgMinAud2PercWeekEnd) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim2) + """ """ + up_down(TotViewTim2PercWeekEnd) + """ """ + str(TotViewTim2PercWeekEnd) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev2) + """ """ + up_down(UniqDev2PercWeekEnd) + """ """ + str(UniqDev2PercWeekEnd) + """ vs. previous year weekends </li>
#         </ul>
#         <p><b><mark>""" + formatted_date + """ Sunday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekEnd) + """ """ + str(AvgMinAud1PercWeekEnd) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekEnd) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekEnd) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
# """)
# elif datetime.weekday(CurrDate) == 1:
#     # Print weekday format
#     # Test HTML Page
#     Function_Name = open("WEATHERTESTDAILY.html","w")
#     Function_Name.write("""
#         <h4 style="font-weight: normal;">Hi all - </h4>
#         <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Monday. Let us know if you have any questions or feedback. </h4>
#         <p><u><b>Key Performance Indicators</b></u></p>
#         <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
#         """)
# elif datetime.weekday(CurrDate) == 2:
#     # Print weekday format
#     # Test HTML Page
#     Function_Name = open("WEATHERTESTDAILY.html","w")
#     Function_Name.write("""
#         <h4 style="font-weight: normal;">Hi all - </h4>
#         <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Tuesday. Let us know if you have any questions or feedback. </h4>
#         <p><u><b>Key Performance Indicators</b></u></p>
#         <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
#         """)
# elif datetime.weekday(CurrDate) == 3:
#     # Print weekday format
#     # Test HTML Page
#     Function_Name = open("WEATHERTESTDAILY.html","w")
#     Function_Name.write("""
#         <h4 style="font-weight: normal;">Hi all - </h4>
#         <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Wednesday. Let us know if you have any questions or feedback. </h4>
#         <p><u><b>Key Performance Indicators</b></u></p>
#         <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
#         """)
# elif datetime.weekday(CurrDate) == 4:
#     # Print weekday format
#     # Test HTML Page
#     Function_Name = open("WEATHERTESTDAILY.html","w")
#     Function_Name.write("""
#         <h4 style="font-weight: normal;">Hi all - </h4>
#         <h4 style="font-weight: normal;">Please see below for FOX Weather updates for Thursday. Let us know if you have any questions or feedback. </h4>
#         <p><u><b>Key Performance Indicators</b></u></p>
#         <p><b><mark>""" + formatted_date + """ Monday </mark></b> <b>""" + """ """ + """includes</b> SamsungTV+, FOX Weather, Amazon News, DirecTV stream, Vizio, LG, Tubi, FOX Nation, FTS, Plex, Xumo, FOX News Digital and <b>excludes</b> Amazon FreeVee, YTTV, Fubo, Youtube.com, Roku Channel)</p>
#         <ul>
#             <li> Average Min Audience - """ + human_format(AvgMinAud1) + """ """ + up_down(AvgMinAud1PercWeekDay) + """ """ + str(AvgMinAud1PercWeekDay) + """ vs. previous year weekdays </li>
#             <li> Total View Time (Mins) - """ + human_format(TotViewTim1) + """ """ + up_down(TotViewTim1PercWeekDay) + """ """ + str(TotViewTim1PercWeekDay) + """ vs. previous year weekends </li>
#             <li> Unique Viewing Devices - """ + human_format(UniqDev1) + """ """ + up_down(UniqDev1PercWeekDay) + """ """ + str(UniqDev1PercWeekDay) + """ vs. previous year weekends </li>
#         </ul>
#         """)


