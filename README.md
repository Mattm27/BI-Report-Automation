## Pandas Package Functions
In this section, functions originated from the pandas library are defined and syntax is provided.

### Importing Pandas Package:
```bash
Import pandas as pd
```

### Loading in Excel Sheets:
```bash
DatasetName = pd.read_excel(Pathname, sheet_name=sheet, skiprows=number)
```
- Can also specify specific sheets to work within or number of rows to skip for subsetting 

### Indexing to Specific Row in Column
```bash
DatasetName[ColumnName].iloc[index]
```

### Converting Observations in Column to Percentage Format:
```bash
DatasetName[ColumnName] = DatasetName[ColumnName].transform(lambda x: '{:,.2%}'.format(x))
```

### Renaming Columns:
```bash
DatatsetName = DatasetName.rename(columns={DatasetName.columns[column#]: NewName})
```

### Locating and Subsetting Data Frame Based on Column Value:
```bash
NewDatasetName= DatasetName.loc[DatasetName[ColumnName] == string parameter]
```

### Selecting Observations That Satisfy Given Parameter:
```bash
DatasetName = DatasetName[DatesetName[ColumnName] > Parameter]
```
- Logical operator can be changed (<,=,!=,>=) and parameter can also be boolean expression (True, False) or a string

### Selecting Greatest Observation in Given Column:
```bash
DatasetName = DatasetName.loc[DatasetName[ColumnName].idxmax()]
```


## Datetime Package Functions
Here are definitions and syntax for functions originating from the datetime library

### Importing Datetime Package:
```bash
Import datetime as datetime
```

### Stripping Time Values From a Date
```bash
date_obj = datetime.strptime(datename,"%Y-%m-%d %H:%M:%S")
```

### Formatt Date to Only Include Certain Elements
```bash
formatted_date = date_obj.strftime("%m/%d")
```
- Can specify between %m, %d, %Y, %b, %B and others when formatting dates depending on desired outcome


## HTML tags
When composing emails, here is how the HTML tags used operate

### Create a header
```<h4> </h4>```
- Can alsp specify h3,h2,h1 to achieve larger headers

### Creates a Paragraph
```<p> </p>```

### Underline Text
```<u> </u>```

### Bold Text
```<b> </b>```

### Initialize List
```<ul> </ul>```

### Create Individual Bullet Points
```<li> </li>```


## Custom Functions:
Below are definitions and syntax for all custom functions created throughout the scripts.

### Podcast Functions
```bash
def up_down(str):
    index = str.find('-')
    if index == -1:
        return('up')
    else:
        return('down')
```
- up_down is used to print the string either “up” or “down” depending on if the percentage change is either positive or negative
```bash
def sim(str1,str2,num2):
    if str1 == str2:
        return(""" and in Unique Listeners with """ + human_format(num2) + """ during""")
    else:
        return(""", while <u>""" + str2 + """</u> set a new historic record for Unique Listeners with """ + human_format(num2) + """ during""")
```
- sim is used to to determine if the given podcast reached a new historical high in unique listeners and outputs the proper HTML code
```bash
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
```
- FBNHelp determines is the Fox Business Hourly Update increased/decreased in unique downloads and unique listeners and outputs the proper HTML code
```bash
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
```
- QuarterEndHelp1 has similar functionality to the FBNHelp but inputs are total unique downloads and unique listeners from current and past quarter
```bash
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
```
- WhichQ is used to determine the current quarter for the report and converts it to a string to be printed in HTML code
```bash
def QuarterEnd(date):
    index = date.find('March')
    index1 = date.find('June')
    index2 = date.find('September')
    index3 = date.find('December')
    if index == -1 and index1 == -1 and index2 == -1 and index3 == -1:
        return("""<h4 style="font-weight: normal;">For the month of """ + str(Date['Date'].iloc[0]) + """, below are key takeaways from Fox News and Fox Business podcast performance. Please let me know if you have any questions. </h4>
                  <p style = "margin-bottom:0;">Thanks!</p>
                  <p style = "margin :0; padding-top:0;">Kayla</p>""")
    else:
        return("""<h4 style="font-weight: normal;">Attached are key takeaways and Quarter-End highlights for month end (""" + date + """) and quarter-end (""" + whichQ(date) + """). Please let me know if you have any questions.</h4>
                  <p style = "margin-bottom:0;">Thanks!</p>
                  <p style = "margin :0; padding-top:0;">Kayla</p>""")
```
- QuarterEnd and QuarterEnd1 are used to format email with proper HTML code if it is the end of a quarter
- Make sure to update any lines that contain FYxx Qx to contain proper year and quarter (i.e line 83, line 228)

### Social Functions
```bash
def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'K', 'M', 'B', 'T'][magnitude])
```
- human_format creates a format to convert the numerics of the report to specify if it is in the Thousands, millions etc.
```bash
def check_up_down(str1,str2):
    if str1 == str2:
        return("and ")
    else:
        return("but ")
```
- check_up_down prints "and" or "but" in email depending on values of parameters
- Need to manually sort Total, Facebook, Instagram, and Twitter Interactions in Excel sheet

### Comscore Functions
```bash
def difference(prev, curr):
    return(abs(curr - prev))
```
- difference is used to determine how many spots an Outkick.com ranking has changed
```bash

def rise_drop(prev, curr):
    if prev > curr:
        return("""moved down """ + str(difference(prev,curr)) + """ spots  to #""" + str(curr))
    elif prev == curr:
        return("""remained in the """ + str(curr) + """spot""")
    else:
        return("""moved up """ + str(difference(prev,curr)) + """ spots  to #""" + str(curr))
```
- rise_drop returns proper html code depending on whether or not a ranking has increased or decreased








