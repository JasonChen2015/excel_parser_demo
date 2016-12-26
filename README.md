**This java demo is aimed at parsing the excel files(include xls & xlsx format) then write the content to text file.**

## Background

Sometimes I want to export the data from excel files to database(DB), but actually DB can't read files with xls or xlsx format. Thus I want to get a text file which DB can accepted first.

For easily reading the data, I want the output data should be seperated by '|'. For instance, "Number|YourName|YourGender|YourAge|". And it may skip some rows(e.g. header) I don't want or ...


## Call Command

```bash
java -cp Class_Path J_Cvt Input_File Ouput_File Output_Column Sheet_No Skip_Row [Prefix_String]
```
J_Cvt.java is used to get all content of a sheet in excel file

```
Input_File:    The input excel file.
Output_File:   The ouput text file.
Output_Column: The column of sheet of the input excel file in order, format: a1 or aa1 (accept both lower case and upper case).
Sheet_No:      The sheet number of input excel file, begins at 1.
Skip_Row:      Skip rows of sheet of the input excel file, begins at 1.
Prefix_String: For individual use, the prefix string of each ouput line, can be ignored.
```

```bash
java -cp Class_Path J_Value Input_File Output_Cell Sheet_No
```
J_Cell.java is used to get one cell  content of a sheet in excel file

```
Input_File:    The input excel file.
Output_Cell:   The cell position of sheet of the input excel file, format: a1 or aa1 (accept both lower case and upper case).
Sheet_No:      The sheet number of input excel file, begins at 1.
```

## Needed Jar File

Apache POI (currently use POI 3.15)

* poi-3.15.jar
* poi-ooxml-3.15.jar
* poi-ooxml-schemas-3.15.jar"
* xmlbeans-2.6.0.jar"

Apache Xerces Project (currently use Xerces2 Java 2.11.0)

* xercesImpl.jar
* xml-apis.jar"


## Features

- Parse xls and xlsx(large memory) file.
- Ouput number, string, date, formula, boolean. But notice that the date is in "yyyy-MM-dd" format.
- Exit at once after getting the needed data.

