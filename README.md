# Jinja Excel Template Engine

This package can handle some Excel reporting requirements.

It contains:

1. Dynamic sheets based on the data
2. Images are inside the Excel file,
3. Multiple level of grouping in the data (this library does not handle it, but you can have your own data pipeline to prepare the data.),
4. Printing header setting,
5. Basic charts


This library will **not** be suitable for you if you are after a completely non-technical maintainable Excel template. In fact, there is none of them, if your Excel template is complicated enough, especially dynamic sheets.


This package will suitable for you for the following scenarios:
1. if you have XML knowledge and know how to read openpyxl documents, and
1. if you have Jinja2 knowledge



# Concept
Excel templating is always complicated and normal Excel file does not support them. We have to use an intermediate file named ExcelXML.

Data + Jinja Template ---(Jinja Engine)---> ExcelXML ---(Excel Generator)---> Excel

With Data and Jinja Template, the Jinja Engine can generate the ExcelXML, then the Excel Generator can generate the Excel file. An ExcelXML is an XML representing a single Excel file.

The ExcelXML has a specific format, so the Excel Generator can read it and draw the Excel.

Here is an example of the ExcelXML.

```xml
<Excel>
    <sheet sheetName="myFirstSheet" print_area="A1:L10" print_header="1:3">
        <row> <!-- first row -->
            <cell value="A1" />
            <cell value="B1" />
            <cell value="C1" />
        </row>
        <row /> <!-- second row, empty row -->
        <row> <!-- third row -->
            <cell value="A3" />
            <cell value="B3" />
            <cell value="C3" />
        </row>
    </sheet>
    <sheet sheetName="mySecondSheet" print_area="A1:L10" print_header="1:3">
        <row> <!-- first row -->
            <cell value="A1" />
            <cell value="B1" />
            <cell value="C1" />
        </row>
        <row /> <!-- second row, empty row -->
        <row> <!-- third row -->
            <cell value="A3" />
            <cell value="B3" />
            <cell value="C3" />
        </row>
    </sheet>
</Excel>
```

The Jinja template is to construct the ExcelXML with your own data pipeline.

Why did I chose XML not JSON? JSON is too messy to representing meta data and attributes, so I chose XML.


# Sample
Please check the `tests/countries_test.py`. It shows the raw data and group countires by region and display countries in multiple sheets. It also contains some basic diagrams.

If you are interested in more complicated example, you can refer to `tests/test.xml` that is the test XML I am using to test my output and it finally hit 99% code coverage.


# TODO list
1. Filtering and grouping
1. Auto column width (the current implement is not good enough)
1. Calculated cell
1. Auto number, date and text detection
1. Spend more time on openpyxl document (yes unfortunately, I haven't got enough time)
1. Pivot table?? (not sure the use case about it, put it in low priority)


# About me
I am from .NET, Java and NodeJS background. This is my first project in Python. When I started this project, I wanted to go to NodeJS, but ExcelJS is not powerful enough comparing to openpyxl. That is why I chose Python. If I am not good at the Python standard, please let me know by email hkalex@gmail.com. Thanks.
