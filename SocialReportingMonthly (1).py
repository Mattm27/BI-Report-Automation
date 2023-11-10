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

def color_code_ahead(value):
            if value > 25:
                return """<b><p style="color:green;">"""
            else:
                return """<b><p style="color:gold;">"""
             
def color_code_fox_news(value):
         if value=="Fox News":
            return """<b><p style="color:blue;">"""
         else:
            return """</b><p style="color:black;">"""

# Date Formatting
date = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
date = date[['Date']].iloc[0:2]

#Total Interactions

overview  = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
Total_Interactions = overview
Total_Interactions = Total_Interactions[['Rank','Site', 'Total Interactions']]

Fox_Rank = Total_Interactions.loc[Total_Interactions['Site'] == 'Fox News']

MonthlyTotal = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal = MonthlyTotal[['Monthly','Yearly']]
MonthlyTotal['Monthly'] = MonthlyTotal['Monthly'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal['Yearly'] = MonthlyTotal['Yearly'].transform(lambda x: '{:,.0%}'.format(x))


print('\nFox News ranked #' + human_format(Fox_Rank['Rank'].iloc[0]) + ' against the competition set for ' + date['Date'].iloc[1]+ '-' + str((date['Date'].iloc[0])) + ', driving in ' + human_format(Fox_Rank['Total Interactions'].iloc[0]) + ' total social interactions (Facebook, Instagram, and Twitter combined)')
print('    - Overall social interactions were ' + MonthlyTotal['Yearly'].iloc[0] + ' versus prior year and ' + MonthlyTotal['Monthly'].iloc[0] + ' versus prior month')


# Facebook/Instagram/Twitter Interactions

Facebook = overview[['Rank1','Site.1', 'Facebook Interactions']]
Fox_Rank1 = Facebook.loc[Facebook['Site.1'] == 'Fox News']

Instagram = overview[['Rank2', 'Site.2', 'Instagram Interactions']]
Fox_Rank2 = Instagram.loc[Instagram['Site.2']=='Fox News']

Twitter = overview[['Rank3', 'Site.3', 'Twitter Interactions']]
Fox_Rank3 = Twitter.loc[Twitter['Site.3']=='Fox News']

MonthlyTotal1 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal1 = MonthlyTotal1[['Monthly1','Yearly1']]
MonthlyTotal1['Monthly1'] = MonthlyTotal1['Monthly1'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal1['Yearly1'] = MonthlyTotal1['Yearly1'].transform(lambda x: '{:,.0%}'.format(x))

MonthlyTotal2 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal2 = MonthlyTotal2[['Monthly2','Yearly2']]
MonthlyTotal2['Monthly2'] = MonthlyTotal2['Monthly2'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal2['Yearly2'] = MonthlyTotal2['Yearly2'].transform(lambda x: '{:,.0%}'.format(x))

MonthlyTotal3 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal3 = MonthlyTotal3[['Monthly3','Yearly3']]
MonthlyTotal3['Monthly3'] = MonthlyTotal3['Monthly3'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal3['Yearly3'] = MonthlyTotal3['Yearly3'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox News ranked #' + human_format(Fox_Rank1['Rank1'].iloc[0]) + ' in social engagement on Facebook, #' + human_format(Fox_Rank2['Rank2'].iloc[0]) + ' on Instagram, and #' + human_format(Fox_Rank3['Rank3'].iloc[0]) + ' on twitter against the competitive set')
print('    - Fox News drove in ' + human_format(Fox_Rank1['Facebook Interactions'].iloc[0]) + ', '+ MonthlyTotal1['Yearly1'].iloc[0] + ' versus prior year and ' + MonthlyTotal1['Yearly1'].iloc[0] + ' versus prior month')
print('    - Fox News drove in ' + human_format(Fox_Rank2['Instagram Interactions'].iloc[0]) + ', '+ MonthlyTotal2['Yearly2'].iloc[0] + ' versus prior year and ' + MonthlyTotal2['Yearly2'].iloc[0] + ' versus prior month')
print('    - Fox News drove in ' + human_format(Fox_Rank3['Twitter Interactions'].iloc[0]) + ', '+ MonthlyTotal3['Yearly3'].iloc[0] + ' versus prior year and ' + MonthlyTotal3['Yearly3'].iloc[0] + ' versus prior month')

# Youtube views

Youtube = overview[['Rank4', 'Site.4', 'Youtube Views']]
Fox_Rank4 = Youtube.loc[Youtube['Site.4']== 'Fox News']

MonthlyTotal4 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal4 = MonthlyTotal4[['Monthly4','Yearly4']]
MonthlyTotal4['Monthly4'] = MonthlyTotal4['Monthly4'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal4['Yearly4'] = MonthlyTotal4['Yearly4'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox News ranked #' + human_format(Fox_Rank4['Rank4'].iloc[0]) + ' in youtube views amongst the competitive set')
print('    - Overall youtube views were ' + MonthlyTotal4['Yearly4'].iloc[0] + ' versus prior year and ' + MonthlyTotal4['Yearly4'].iloc[0] + ' versus prior month')


# Business Total

BizTotal = overview[['Rank5','Site.5', 'Business Total Interactions']]
Fox_Rank5 = BizTotal.loc[BizTotal['Site.5'] == 'Fox Business']


MonthlyTotal5 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal5 = MonthlyTotal5[['Monthly5','Yearly5']]
MonthlyTotal5['Monthly5'] = MonthlyTotal5['Monthly5'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal5['Yearly5'] = MonthlyTotal5['Yearly5'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox Business ranked #' + human_format(Fox_Rank5['Rank5'].iloc[0]) + ' agaimst the competitive set for ' + date['Date'].iloc[1]+ '-' + str((date['Date'].iloc[0])) + ' driving in ' +  human_format(Fox_Rank5['Business Total Interactions'].iloc[0]) + ' total social interactions (Facebook, Twitter, and Instagram combined)')
print('    - Overall social interactions were ' + MonthlyTotal5['Yearly5'].iloc[0] + ' versus prior year and ' + MonthlyTotal5['Monthly5'].iloc[0] + ' versus prior month')
# Business Facebook 

BizFacebook = overview[['Rank6','Site.6', 'Business Facebook Interactions']]
Fox_Rank6 = BizFacebook.loc[BizFacebook['Site.6'] == 'Fox Business']


MonthlyTotal6 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal6 = MonthlyTotal6[['Monthly6','Yearly6']]
MonthlyTotal6['Monthly6'] = MonthlyTotal6['Monthly6'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal6['Yearly6'] = MonthlyTotal6['Yearly6'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox Business ranked #' + human_format(Fox_Rank6['Rank6'].iloc[0]) + ' in social engagement on Facebook')
print('    - Overall drove in ' + human_format(Fox_Rank6['Business Facebook Interactions'].iloc[0]) + ' interactions and were ' + MonthlyTotal6['Yearly6'].iloc[0] + ' versus prior year and ' + MonthlyTotal6['Monthly6'].iloc[0] + ' versus prior month')

# Business Instagram

BizInstagram = overview[['Rank7','Site.7', 'Business Instagram Interactions']]
Fox_Rank7 = BizInstagram.loc[BizInstagram['Site.7'] == 'Fox Business']

MonthlyTotal7 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal7['Monthly7'] = MonthlyTotal7['Monthly7'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal7['Yearly7'] = MonthlyTotal7['Yearly7'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox Business ranked #' + human_format(Fox_Rank7['Rank7'].iloc[0]) + ' in social engagement on Instagram')
print('    - Overall drove in ' + human_format(Fox_Rank7['Business Instagram Interactions'].iloc[0]) + ' interactions and were ' + MonthlyTotal7['Yearly7'].iloc[0] + ' versus prior year and ' + MonthlyTotal7['Monthly7'].iloc[0] + ' versus prior month')

# Business Twitter

BizTwitter = overview[['Rank8','Site.8', 'Business Twitter Interactions']]
Fox_Rank8 = BizTwitter.loc[BizTwitter['Site.8'] == 'Fox Business']

MonthlyTotal8 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal8 = MonthlyTotal8[['Monthly8','Yearly8']]
MonthlyTotal8['Monthly8'] = MonthlyTotal8['Monthly8'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal8['Yearly8'] = MonthlyTotal8['Yearly8'].transform(lambda x: '{:,.0%}'.format(x))

print('\nFox Business ranked #' + human_format(Fox_Rank8['Rank8'].iloc[0]) + ' in social engagement on Twitter')
print('    - Overall drove in ' + human_format(Fox_Rank8['Business Twitter Interactions'].iloc[0]) + ' interactions and were ' + MonthlyTotal8['Yearly8'].iloc[0] + ' versus prior year and ' + MonthlyTotal8['Monthly8'].iloc[0] + ' versus prior month')
      
# Business Youtube

BizYoutube = overview[['Rank9', 'Site.9', 'Business Youtube Views']]
Fox_Rank9 = BizYoutube.loc[BizYoutube['Site.9'] == 'Fox Business']

MonthlyTotal9 = pd.read_excel(r'/Users/VayosKa\Desktop\Social Reporting\FNC-FBN Historical Trends_Social Media Highlights 20230801.xlsm', sheet_name='Email Auto')
MonthlyTotal9 = MonthlyTotal9[['Monthly9','Yearly9']]
MonthlyTotal9['Monthly9'] = MonthlyTotal9['Monthly9'].transform(lambda x: '{:,.0%}'.format(x))
MonthlyTotal9['Yearly9'] = MonthlyTotal9['Yearly9'].transform(lambda x: '{:,.0%}'.format(x))

# Helper functions
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')

# Function to count consecutive months Fox has been rank #1 in any category (Specify by changing .txt file and which table you pull teh rank from)
# def number_one_total(textfile,num):
#     if num == "1":
#         f = open(textfile,"r")
#         val = int(f.read())
#         val = val + 1
#         f.close()
#         f1 = open(textfile,"w")
#         f1.write(str(val))
#         f1.close()
#         return(val)
#     else:
#         val = 0
#         f = open(textfile,"w")
#         f.write("0")
#         f.write("0")
#         f.close()
#         return(val)

# Create Email 
# ** Note: win32client is not available on MacOS so the following code has not been tested
import win32com.client
ol = win32com.client.Dispatch('Outlook.Application')
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

newmail.Subject = 'Monthly Social Update - ' + date['Date'].iloc[1]+ ' ' + str((date['Date'].iloc[0]))
newmail.HTMLBody = """ 
<h4 style="font-weight: normal;">Hi all - </h4>
<h4 style="font-weight: normal;">For the month of """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """, below are the social highlights for both News and Business. </h4>
<p><strong><u><span style="background-color: rgb(247, 218, 100);"> """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ Social Highlights </span></u></strong></p>
<p><strong><u>News Social Highlights - """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ (Source Emplifi):</u></strong></p>
<ul>
    <li>Fox News <strong>ranked #""" + human_format(Fox_Rank['Rank'].iloc[0]) + """ against the competitive set for """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ driving in """ + human_format(Fox_Rank['Total Interactions'].iloc[0]) + """ total interactions</strong> (Facebook, Instagram, Twitter combined)<ul>  
            <li>Overall social interactions were """ + up_down(MonthlyTotal['Yearly'].iloc[0]) + """ """ + MonthlyTotal['Yearly'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal['Monthly'].iloc[0]) + """ """ + MonthlyTotal['Monthly'].iloc[0] + """ vs. prior month.</li>
        </ul>
    </li>
    <li>Fox News ranked <strong> #""" + human_format(Fox_Rank1['Rank1'].iloc[0]) + """ in social engagement on facebook</strong>, <strong>#""" + human_format(Fox_Rank2['Rank2'].iloc[0]) + """ on Instagram</strong>, and <strong> #""" + human_format(Fox_Rank3['Rank3'].iloc[0]) + """ on twitter</strong> amongst the competitive set<ul>
            <li>Fox News drove in """ + human_format(Fox_Rank1['Facebook Interactions'].iloc[0]) + """ Facebook interactions, """ + up_down(MonthlyTotal1['Yearly1'].iloc[0]) + """ """ + MonthlyTotal1['Yearly1'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal1['Monthly1'].iloc[0])+ """ """ + MonthlyTotal1['Monthly1'].iloc[0] + """ vs. prior month</li>
            <li>Fox News drove in """ + human_format(Fox_Rank2['Instagram Interactions'].iloc[0]) + """ Instagram interactions, """+ up_down(MonthlyTotal2['Yearly2'].iloc[0]) + """ """ + MonthlyTotal2['Yearly2'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal2['Monthly2'].iloc[0])+ """ """ + MonthlyTotal2['Monthly2'].iloc[0] + """ vs. prior month</li>
            <li>Fox News drove in """ + human_format(Fox_Rank3['Twitter Interactions'].iloc[0]) + """ Twitter interactions, """ + up_down(MonthlyTotal3['Yearly3'].iloc[0]) + """ """ + MonthlyTotal3['Yearly3'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal3['Monthly3'].iloc[0])+ """ """ + MonthlyTotal3['Monthly3'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
    <li>Fox News ranked  <strong> #""" + human_format(Fox_Rank4['Rank4'].iloc[0]) + """ in youtube views amongst the News competitive set </strong> <ul>
            <li>Fox News drove in """ + human_format(Fox_Rank4['Youtube Views'].iloc[0]) + """ YouTube Views, """ + up_down(MonthlyTotal4['Yearly4'].iloc[0]) + """ """ + MonthlyTotal4['Yearly4'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal4['Monthly4'].iloc[0]) + """ """ + MonthlyTotal4['Monthly4'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
</ul>
<p><strong><u>Business Social Highlights - """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ (Source Emplifi):</u></strong></p>
<ul>
    <li>Fox Business ranked <strong> #""" + human_format(Fox_Rank5['Rank5'].iloc[0]) + """ against the competitive set for """ + date['Date'].iloc[1] + """ """ + str((date['Date'].iloc[0])) + """, driving in """ + human_format(Fox_Rank5['Business Total Interactions'].iloc[0]) + """ total social interactions </strong>(Facebook, Instagram, Twitter combined)<ul>
            <li>Overall social interactions were """ + up_down(MonthlyTotal5['Yearly5'].iloc[0]) + """ """ + MonthlyTotal5['Yearly5'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal5['Monthly5'].iloc[0]) + """ """ + MonthlyTotal5['Monthly5'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
    <li>Fox Business ranked <strong> #""" + human_format(Fox_Rank6['Rank6'].iloc[0]) + """ in social engagement on Facebook</strong>
        <ul>
            <li>Overall drove in """ + human_format(Fox_Rank6['Business Facebook Interactions'].iloc[0]) + """ interactions and were """ + up_down(MonthlyTotal6['Yearly6'].iloc[0]) + """ """ + MonthlyTotal6['Yearly6'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal6['Monthly6'].iloc[0]) + """ """ + MonthlyTotal6['Monthly6'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
    <li>Fox Business ranked<strong>&nbsp;</strong><strong> #""" + human_format(Fox_Rank7['Rank7'].iloc[0]) + """ in social engagement on Instagram</strong>
        <ul>
            <li>Overall drove in """ + human_format(Fox_Rank7['Business Instagram Interactions'].iloc[0]) + """ interactions and were """ + up_down(MonthlyTotal7['Yearly7'].iloc[0]) + """ """ + MonthlyTotal7['Yearly7'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal7['Monthly7'].iloc[0]) + """ """ + MonthlyTotal7['Monthly7'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
    <li>Fox Business ranked <strong>#""" + human_format(Fox_Rank9['Rank9'].iloc[0]) + """ in Youtube views</strong> against competitors, driving in """ + human_format(Fox_Rank9['Business Youtube Views'].iloc[0]) + """ video views (Source: Shareablee) &nbsp;<ul>
            <li>On Youtube, Fox Business video views were """ + up_down(MonthlyTotal9['Yearly9'].iloc[0]) + """ """ + MonthlyTotal9['Yearly9'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal9['Monthly9'].iloc[0]) + """ """ + MonthlyTotal9['Monthly9'].iloc[0] + """ vs. prior month</li>
        </ul>
    </li>
</ul>
"""
#<p><u><strong>News Competitive Set Top 5:</strong></u></p>
#<table style="width: 100%; border-collapse: collapse; border: 1px solid rgb(0, 0, 0);">
#    <tbody>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Rank</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Total Int.</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Facebook Int.</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Instagram Int.</strong></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);"><strong>Twitter Int.</strong></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);"><strong>Youtube Views</strong></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">1.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[0] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[0] + """<br></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[0] + """<br></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[0] + """<br></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[0] + """<br></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%;">2.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[1] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[1] + """<br></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[1] + """<br></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[1] + """<br></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[1] + """<br></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">3.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[2] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[2] + """<br></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[2] + """<br></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[2] + """<br></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[2] + """<br></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">4.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[3] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[3] + """<br></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[3] + """<br></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[3] + """<br></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[3] + """<br></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">5.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[4] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[4] + """<br></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[4] + """<br></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[4] + """<br></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[4] + """<br></td>
#        </tr>
#    </tbody>
#</table>
#<p><u><strong>Business News Competitive Set Rankings:</strong></u></p>
#<table style="width: 100%; border-collapse: collapse; border: 1px solid rgb(0, 0, 0);">
#    <tbody>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Rank</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Total Int.</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Facebook Int.</strong></td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Instagram Int.</strong></td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);"><strong>Twitter Int.</strong></td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);"><strong>Youtube Views</strong></td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">1.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[0] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[0] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[0] + """</td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[0] + """</td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[0] + """</td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">2.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[1] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[1] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[1] + """</td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[1] + """</td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[1] + """</td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">3.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[2] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[2] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[2] + """</td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[2] + """</td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[2] + """</td>
#        </tr>
#        <tr>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">4.</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[3] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[3] + """</td>
#            <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[3] + """</td>
#            <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[3] + """</td>
#            <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[3] + """</td>
#        </tr>
#    </tbody>
#</table>
"""<p><br></p>
"""

newmail.To = "kayla.vayos@fox.com"
newmail.display()


#newMail.send()

# Test HTML Page
# Function_Name = open("GFG-1.html","w")

# Function_Name.write(""" 
# <h4 style="font-weight: normal;">Hi all - </h4>
# <h4 style="font-weight: normal;">For the month of """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """, below are the social highlights for both News and Business. </h4>
# <p><strong><u><span style="background-color: rgb(247, 218, 100);"> """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ Social Highlights </span></u></strong></p>
# <p><strong><u>News Social Highlights - """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ (Source Emplifi):</u></strong></p>
# <ul>
#     <li>Fox News <strong>ranked #""" + human_format(Fox_Rank['Rank'].iloc[0]) + """ against the competitive set for """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ driving in """ + human_format(Fox_Rank['Total Interactions'].iloc[0]) + """ total interactions</strong> (Facebook, Instagram, Twitter combined)<ul>
#             <li>Overall social interactions were """ + up_down(MonthlyTotal['Yearly'].iloc[0]) + """ """ + MonthlyTotal['Yearly'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal['Monthly'].iloc[0]) + """ """ + MonthlyTotal['Monthly'].iloc[0] + """ vs. prior month.</li>
#         </ul>
#     </li>
#     <li>Fox News ranked <strong> #""" + human_format(Fox_Rank1['Rank1'].iloc[0]) + """ in social engagement on facebook</strong>, <strong>#""" + human_format(Fox_Rank2['Rank2'].iloc[0]) + """ on Instagram</strong>, and <strong> #""" + human_format(Fox_Rank3['Rank3'].iloc[0]) + """ on twitter</strong> amongst the competitive set<ul>
#             <li>Fox News drove in """ + human_format(Fox_Rank1['Facebook Interactions'].iloc[0]) + """ Facebook interactions, """ + up_down(MonthlyTotal1['Yearly1'].iloc[0]) + """ """ + MonthlyTotal1['Yearly1'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal1['Monthly1'].iloc[0])+ """ """ + MonthlyTotal1['Monthly1'].iloc[0] + """ vs. prior month</li>
#             <li>Fox News drove in """ + human_format(Fox_Rank2['Instagram Interactions'].iloc[0]) + """ Instagram interactions, """+ up_down(MonthlyTotal2['Yearly2'].iloc[0]) + """ """ + MonthlyTotal2['Yearly2'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal2['Monthly2'].iloc[0])+ """ """ + MonthlyTotal2['Monthly2'].iloc[0] + """ vs. prior month</li>
#             <li>Fox News drove in """ + human_format(Fox_Rank3['Twitter Interactions'].iloc[0]) + """ Twitter interactions, """ + up_down(MonthlyTotal3['Yearly3'].iloc[0]) + """ """ + MonthlyTotal3['Yearly3'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal3['Monthly3'].iloc[0])+ """ """ + MonthlyTotal3['Monthly3'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
#     <li>Fox News ranked  <strong> #""" + human_format(Fox_Rank4['Rank4'].iloc[0]) + """ in youtube views amongst the News competitive set </strong> <ul>
#             <li>Overall youtube views were """ + up_down(MonthlyTotal4['Yearly4'].iloc[0]) + """ """ + MonthlyTotal4['Yearly4'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal4['Monthly4'].iloc[0]) + """ """ + MonthlyTotal4['Monthly4'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
# </ul>
# <p><strong><u>Business Social Highlights - """ + date['Date'].iloc[1]+ """ """ + str((date['Date'].iloc[0])) + """ (Source Emplifi):</u></strong></p>
# <ul>
#     <li>Fox Business ranked <strong> #""" + human_format(Fox_Rank5['Rank5'].iloc[0]) + """ against the competitive set for """ + date['Date'].iloc[1] + """ """ + str((date['Date'].iloc[0])) + """, driving in """ + human_format(Fox_Rank5['Business Total Interactions'].iloc[0]) + """ total social interactions </strong>(Facebook, Instagram, Twitter combined)<ul>
#             <li>Overall social interactions were """ + up_down(MonthlyTotal5['Yearly5'].iloc[0]) + """ """ + MonthlyTotal5['Yearly5'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal5['Monthly5'].iloc[0]) + """ """ + MonthlyTotal5['Monthly5'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
#     <li>Fox Business ranked <strong> #""" + human_format(Fox_Rank6['Rank6'].iloc[0]) + """ in social engagement on Facebook</strong>
#         <ul>
#             <li>Overall drove in """ + human_format(Fox_Rank6['Business Facebook Interactions'].iloc[0]) + """ interactions and were """ + up_down(MonthlyTotal6['Yearly6'].iloc[0]) + """ """ + MonthlyTotal6['Yearly6'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal6['Monthly6'].iloc[0]) + """ """ + MonthlyTotal6['Monthly6'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
#     <li>Fox Business ranked<strong>&nbsp;</strong><strong> #""" + human_format(Fox_Rank7['Rank7'].iloc[0]) + """ in social engagement on Instagram</strong>
#         <ul>
#             <li>Overall drove in """ + human_format(Fox_Rank7['Business Instagram Interactions'].iloc[0]) + """ interactions and were """ + up_down(MonthlyTotal7['Yearly7'].iloc[0]) + """ """ + MonthlyTotal7['Yearly7'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal7['Monthly7'].iloc[0]) + """ """ + MonthlyTotal7['Monthly7'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
#     <li>Fox Business ranked <strong> #""" + human_format(Fox_Rank8['Rank8'].iloc[0]) + """ in social engagement on Twitter</strong>
#         <ul>
#             <li>Overall drove in """ + human_format(Fox_Rank8['Business Twitter Interactions'].iloc[0]) + """ interactions and were """ + up_down(MonthlyTotal8['Yearly8'].iloc[0]) + """ """ + MonthlyTotal8['Yearly8'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal8['Monthly8'].iloc[0]) + """ """ + MonthlyTotal8['Monthly8'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
#     <li>Fox Business ranked <strong>#""" + human_format(Fox_Rank9['Rank9'].iloc[0]) + """ in Youtube views</strong> against competitors, driving in """ + human_format(Fox_Rank9['Business Youtube Views'].iloc[0]) + """ video views (Source: Shareablee) &nbsp;<ul>
#             <li>On Youtube, Fox Business video views were """ + up_down(MonthlyTotal9['Yearly9'].iloc[0]) + """ """ + MonthlyTotal9['Yearly9'].iloc[0] + """ vs. prior year and """ + up_down(MonthlyTotal9['Monthly9'].iloc[0]) + """ """ + MonthlyTotal9['Monthly9'].iloc[0] + """ vs. prior month</li>
#         </ul>
#     </li>
# </ul>
# <p><u><strong>News Competitive Set Top 5:</strong></u></p>
# <table style="width: 100%; border-collapse: collapse; border: 1px solid rgb(0, 0, 0);">
#     <tbody>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Rank</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Total Int.</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Facebook Int.</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Instagram Int.</strong></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);"><strong>Twitter Int.</strong></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);"><strong>Youtube Views</strong></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">1.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[0] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[0] + """<br></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[0] + """<br></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[0] + """<br></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[0] + """<br></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%;">2.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[1] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[1] + """<br></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[1] + """<br></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[1] + """<br></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[1] + """<br></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">3.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[2] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[2] + """<br></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[2] + """<br></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[2] + """<br></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[2] + """<br></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">4.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[3] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[3] + """<br></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[3] + """<br></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[3] + """<br></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[3] + """<br></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">5.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Total_Interactions['Site'].iloc[4] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Facebook['Site.1'].iloc[4] + """<br></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + Instagram['Site.2'].iloc[4] + """<br></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + Twitter['Site.3'].iloc[4] + """<br></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + Youtube['Site.4'].iloc[4] + """<br></td>
#         </tr>
#     </tbody>
# </table>
# <p><u><strong>Business News Competitive Set Rankings:</strong></u></p>
# <table style="width: 100%; border-collapse: collapse; border: 1px solid rgb(0, 0, 0);">
#     <tbody>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Rank</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Total Int.</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Facebook Int.</strong></td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);"><strong>Instagram Int.</strong></td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);"><strong>Twitter Int.</strong></td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);"><strong>Youtube Views</strong></td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">1.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[0] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[0] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[0] + """</td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[0] + """</td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[0] + """</td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">2.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[1] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[1] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[1] + """</td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[1] + """</td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[1] + """</td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">3.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[2] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[2] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[2] + """</td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[2] + """</td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[2] + """</td>
#         </tr>
#         <tr>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">4.</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizTotal['Site.5'].iloc[3] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizFacebook['Site.6'].iloc[3] + """</td>
#             <td style="width: 16.6667%; border: 1px solid rgb(0, 0, 0);">""" + BizInstagram['Site.7'].iloc[3] + """</td>
#             <td style="width: 15.268%; border: 1px solid rgb(0, 0, 0);">""" + BizTwitter['Site.8'].iloc[3] + """</td>
#             <td style="width: 18.1239%; border: 1px solid rgb(0, 0, 0);">""" + BizYoutube['Site.9'].iloc[3] + """</td>
#         </tr>
#     </tbody>
# </table>
# <p><br></p>
# """)

# Function_Name.close()
