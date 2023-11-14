## Pandas Package Functions
In this section, functions originated from the pandas library are defined and syntax is provided.

### Importing Pandas Package:
```Import pandas as pd```

### Loading in Excel Sheets:
```DatasetName = pd.read_excel(Pathname, sheet_name=sheet, skiprows=number)```
- Can also specify specific sheets to work within or number of rows to skip for subsetting 

### Indexing to Specific Row in Column
```DatasetName[ColumnName].iloc[index]```

### Converting Observations in Column to Percentage Format:
```DatasetName[ColumnName] = DatasetName[ColumnName].transform(lambda x: '{:,.2%}'.format(x))```

### Renaming Columns:
```DatatsetName = DatasetName.rename(columns={DatasetName.columns[column#]: NewName})```

### Locating and Subsetting Data Frame Based on Column Value:
```NewDatasetName= DatasetName.loc[DatasetName[ColumnName] == string parameter]```

### Selecting Observations That Satisfy Given Parameter:
```DatasetName = DatasetName[DatesetName[ColumnName] > Parameter]```
- Logical operator can be changed (<,=,!=,>=) and parameter can also be boolean expression (True, False) or a string

### Selecting Greatest Observation in Given Column:
```DatasetName = DatasetName.loc[DatasetName[ColumnName].idxmax()]```


## Datetime Package Functions
Here are definitions and syntax for functions originating from the datetime library

### Importing Datetime Package:
```Import datetime as datetime```

### Stripping Time Values From a Date
```date_obj = datetime.strptime(datename,"%Y-%m-%d %H:%M:%S")```

### Formatt Date to Only Include Certain Elements
```formatted_date = date_obj.strftime("%m/%d")```
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
- up_down is used to print the string either “up” or “down” depending on if the percentage change is either positive or negative
```
- sim is used to to determine if the given podcast reached a new historical high in unique listeners and outputs the proper HTML code
- FBNHelp determines is the Fox Business Hourly Update increased/decreased in unique downloads and unique listeners and outputs the proper HTML code
- QuarterEndHelp1 has similar functionality to the FBNHelp but inputs are total unique downloads and unique listeners from current and past quarter
- WhichQ is used to determine the current quarter for the report and converts it to a string to be printed in HTML code
- QuarterEnd and QuarterEnd1 are used to format email with proper HTML code if it is the end of a quarter

### Social Functions
- number_one_total is used to keep track of how many months Fox has been ranked #1 in a specific category by updating a external .txt file each time the script is run, either +1 or resetting to 0
- human_format creates a format to convert the numerics of the report to specify if it is in the Thousands, millions etc.

### Comscore Functions
- difference is used to determine how many spots an Outkick.com ranking has changed
- rise_drop returns proper html code depending on whether or not a ranking has increased or decreased








