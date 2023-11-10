# Documentation
The following contains function definitions and syntax that were utilized in the .py files above to automate email construction.

## Pandas Package Functions
In this section, functions originated from the pandas library are defined and syntax is provided.

### Importing Pandas Package:
- At the top of the sheet packages must be imported in order to utilize their function libraries
```Import pandas as pd```

### Loading in Excel Sheets:
- Used to load in excel files into python. Can also specify specific sheets to work within or number of rows to skip for subsetting 
```DatasetName = pd.read_excel(Pathname, sheet_name=sheet, skiprows=number)```

### Indexing to Specific Row in Column
```DatasetName[ColumnName].iloc[index]```

### Converting Observations in Column to Percentage Format:
```DatasetName[ColumnName] = DatasetName[ColumnName].transform(lambda x: '{:,.2%}'.format(x))```

### Renaming Columns:
```DatatsetName = DatasetName.rename(columns={DatasetName.columns[column#]: NewName})```

### Locating and Subsetting Data Frame Based on Column Value:
```NewDatasetName= DatasetName.loc[DatasetName[ColumnName] == string parameter]```

### Selecting Observations That Satisfy Given Parameter:
- Logical operator can be changed (<,=,!=,>=) and parameter can also be boolean expression (True, False) or a string
```DatasetName = DatasetName[DatesetName[ColumnName] > Parameter]```

### Selecting Greatest Observation in Given Column:
```DatasetName = DatasetName.loc[DatasetName[ColumnName].idxmax()]```

#





