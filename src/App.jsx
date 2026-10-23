import React, { useEffect, useMemo, useState } from "react";
import { BookOpen, CheckCircle2, Code2, Database, FileSpreadsheet, Gauge, GraduationCap, LayoutDashboard, RotateCcw, Search, Table2, Terminal, Trophy, X } from "lucide-react";

const LEVELS = ["Beginner", "Intermediate", "Professional"];
const modules = {
  "Excel": { icon: FileSpreadsheet, lessons: { "Beginner": ["Workbook Fundamentals", "Data Entry and Cell References", "Essential Math Formulas", "Text and Date Functions", "Formatting for Readability", "Sorting and Filtering", "Excel Tables", "Introduction to Charts", "Conditional Formatting", "Beginner Sales Tracker Project"], "Intermediate": ["Logical Functions", "XLOOKUP and INDEX-MATCH", "SUMIFS and COUNTIFS", "Data Validation and Named Ranges", "PivotTables", "PivotCharts and Slicers", "Cleaning Data with Power Query", "Dynamic Array Formulas", "What-If Analysis and Goal Seek", "Interactive Operations Dashboard"], "Professional": ["Advanced Formula Architecture with LET", "Power Pivot Data Models", "DAX Measures in Excel", "Advanced Power Query M", "VBA and Macro Automation", "Workbook Performance Optimization", "Financial Modeling", "Forecasting and Scenario Design", "Governance and Error Controls", "Executive Analytics Solution"] } },
  "SQL": { icon: Database, lessons: { "Beginner": ["Relational Database Foundations", "SELECT and FROM", "Filtering with WHERE", "Sorting and Limiting Results", "Aliases and Calculated Columns", "Aggregate Functions", "GROUP BY and HAVING", "Introduction to JOINs", "NULL and CASE", "Beginner Sales Database Project"], "Intermediate": ["Multi-Table JOINs", "Subqueries", "Common Table Expressions", "Window Functions", "String Functions", "Date and Time Analysis", "Set Operators", "INSERT UPDATE and DELETE", "Views and Reusable Queries", "Customer Analytics Project"], "Professional": ["Query Execution and Optimization", "Index Design", "Advanced Window Patterns", "Recursive CTEs", "Transactions and Concurrency", "Stored Procedures", "Dimensional Data Modeling", "Security and Permissions", "Data Warehousing Patterns", "Production Analytics Pipeline"] } },
  "Google Sheets": { icon: Table2, lessons: { "Beginner": ["Sheets Interface and Files", "Cells Ranges and References", "Core Formulas", "Text and Date Functions", "Formatting and Themes", "Sort and Filter", "Data Validation", "Charts", "Sharing and Collaboration", "Beginner Team Tracker Project"], "Intermediate": ["IF and Logical Functions", "Lookup Functions", "QUERY Function", "FILTER and Array Formulas", "Pivot Tables", "Named Ranges", "Import Functions", "Advanced Charts", "Protected Workflows", "Collaborative Dashboard"], "Professional": ["Advanced QUERY Patterns", "Scalable Array Formulas", "Apps Script Fundamentals", "Custom Functions", "Triggers and Automation", "API Workflow Concepts", "Performance Design", "Security and Governance", "Connected Data Workflows", "Automated Reporting System"] } },
  "Power BI": { icon: Gauge, lessons: { "Beginner": ["Business Intelligence Foundations", "Power BI Desktop Tour", "Connecting to Data", "Power Query Basics", "Core Visuals", "Filters and Slicers", "Table Relationships", "First DAX Measures", "Report Design Principles", "Beginner Sales Report Project"], "Intermediate": ["Star Schema Modeling", "DAX Evaluation Context", "Time Intelligence", "Advanced Power Query", "Drillthrough and Tooltips", "Bookmarks and Navigation", "Row-Level Security", "Performance Analyzer", "Power BI Service", "Operations Dashboard Project"], "Professional": ["Enterprise Semantic Models", "Advanced DAX Patterns", "Calculation Groups Concepts", "Incremental Refresh", "Composite Model Strategy", "Deployment Pipelines", "Governance and Lineage", "Model Optimization", "Executive Data Storytelling", "Enterprise BI Solution"] } },
  "Tableau": { icon: LayoutDashboard, lessons: { "Beginner": ["Visual Analytics Foundations", "Connect to Data", "Dimensions and Measures", "Core Charts", "Filters and Sorting", "Groups and Sets", "Calculated Fields", "Maps", "Dashboard Basics", "Beginner Sales Story Project"], "Intermediate": ["Table Calculations", "Level of Detail Expressions", "Parameters", "Relationships and Data Blending", "Dashboard Actions", "Advanced Mapping", "Date Analysis", "Dashboard Usability", "Publishing and Permissions", "Interactive KPI Dashboard"], "Professional": ["Advanced LOD Patterns", "Table Calculation Architecture", "Performance Recording", "Data Source Strategy", "Dynamic Dashboard Systems", "Statistical Analytics", "Enterprise Governance", "Embedded Analytics Concepts", "Executive Storytelling", "Enterprise Tableau Solution"] } },
  "Python": { icon: Terminal, lessons: { "Beginner": ["Python Setup and Syntax", "Variables and Data Types", "Conditions", "Loops", "Functions", "Lists and Tuples", "Dictionaries and Sets", "Files and Exceptions", "Introduction to pandas", "Beginner Sales Analyzer Project"], "Intermediate": ["Object-Oriented Python", "Comprehensions and Generators", "Modules and Virtual Environments", "NumPy Foundations", "pandas Transformation", "Data Cleaning", "Data Visualization", "Regular Expressions", "Testing and Debugging", "Analytics Pipeline Project"], "Professional": ["Software Architecture", "Advanced Python Patterns", "Type Hints and Validation", "Performance and Profiling", "Concurrent Workflows", "Database Integration", "API Design Concepts", "Production Testing", "Packaging and Deployment", "Production Data Application"] } }
};

const lessonExamples = {
  Excel: [
    ["Create sheets named Raw_Data, Calculations, and Dashboard.", ["Sheet", "Purpose", "Status"], [["Raw_Data", "Store source records", "Ready"], ["Calculations", "Formulas and analysis", "Ready"], ["Dashboard", "Charts and KPIs", "Ready"]]],
    ["Enter =A2*B2 in C2 and =$E$1*C2 in F2.", ["Units", "Price", "Relative Result", "Fixed Rate", "Final Result"], [["10", "25", "250", "12%", "30"], ["8", "40", "320", "12%", "38.40"]]],
    ["=SUM(D2:D11)\n=AVERAGE(D2:D11)\n=MIN(D2:D11)\n=MAX(D2:D11)", ["Function", "Formula", "Result"], [["SUM", "=SUM(D2:D11)", "1,850"], ["AVERAGE", "=AVERAGE(D2:D11)", "185"], ["MIN", "=MIN(D2:D11)", "95"], ["MAX", "=MAX(D2:D11)", "320"]]],
    ["=TEXTJOIN(\" \",TRUE,A2,B2)\n=TEXT(C2,\"mmmm yyyy\")", ["First Name", "Last Name", "Date", "Output"], [["Ana", "Reyes", "2026-01-15", "Ana Reyes | January 2026"]]],
    ["Format Revenue as Currency and Growth as Percentage.", ["Field", "Raw Value", "Displayed Format"], [["Revenue", "125000.5", "$125,000.50"], ["Growth", "0.127", "12.7%"]]],
    ["Filter Region to East and sort Revenue largest to smallest.", ["Rank", "Region", "Revenue"], [["1", "East", "$18,400"], ["2", "East", "$14,250"], ["3", "East", "$9,800"]]],
    ["Convert A1:D100 to SalesData and use =SUM(SalesData[Revenue]).", ["Table", "Rows", "Expansion", "Total"], [["SalesData", "101", "Automatic", "$248,500"]]],
    ["Insert a line chart using Month and Sales.", ["Month", "Sales", "Trend"], [["Jan", "$42,000", "Baseline"], ["Feb", "$47,500", "+13.1%"], ["Mar", "$51,200", "+7.8%"]]],
    ["Highlight D2:D100 red when values are below 0.", ["Product", "Profit", "Rule", "Result"], [["A", "$2,400", "Profit < 0", "No highlight"], ["B", "-$650", "Profit < 0", "Red fill"]]],
    ["Build a table with Date, Product, Region, Units, Price, and Revenue.", ["Product", "Units", "Price", "Revenue"], [["Laptop", "4", "$850", "$3,400"], ["Monitor", "7", "$220", "$1,540"]]]
  ],
  SQL: [
    ["CREATE TABLE customers (customer_id INT PRIMARY KEY, name VARCHAR(100));", ["Table", "Primary Key", "Constraint"], [["customers", "customer_id", "Unique, not null"]]],
    ["SELECT order_id, customer_id, sales FROM orders;", ["order_id", "customer_id", "sales"], [["1001", "C101", "1250"], ["1002", "C102", "980"]]],
    ["SELECT * FROM orders WHERE sales > 1000;", ["order_id", "sales", "Filter Test"], [["1001", "1250", "Included"], ["1003", "1540", "Included"]]],
    ["SELECT * FROM orders ORDER BY sales DESC LIMIT 10;", ["rank", "order_id", "sales"], [["1", "1042", "9850"], ["2", "1018", "8730"]]],
    ["SELECT quantity * unit_price AS revenue FROM order_items;", ["quantity", "unit_price", "revenue"], [["4", "250", "1000"], ["3", "425", "1275"]]],
    ["SELECT COUNT(*) AS orders, SUM(sales) AS total_sales FROM orders;", ["orders", "total_sales", "average_sale"], [["125", "248500", "1988"]]],
    ["SELECT region, SUM(sales) total_sales FROM orders GROUP BY region HAVING SUM(sales)>100000;", ["region", "total_sales", "HAVING"], [["East", "142000", "Included"], ["West", "119500", "Included"]]],
    ["SELECT o.order_id,c.name FROM orders o JOIN customers c ON o.customer_id=c.customer_id;", ["order_id", "customer_name", "match"], [["1001", "Ana Trading", "JOIN match"], ["1002", "Bright Retail", "JOIN match"]]],
    ["SELECT CASE WHEN shipped_date IS NULL THEN 'Open' ELSE 'Shipped' END AS status FROM orders;", ["order_id", "shipped_date", "status"], [["1001", "2026-01-18", "Shipped"], ["1002", "NULL", "Open"]]],
    ["SELECT region,COUNT(*) orders,SUM(sales) sales FROM orders GROUP BY region;", ["region", "orders", "sales"], [["East", "48", "142000"], ["West", "39", "119500"]]]
  ],
  Python: [
    ["print(\"Hello, data!\")", ["Code", "Console Output", "Status"], [["print('Hello, data!')", "Hello, data!", "Success"]]],
    ["units = 12\nprice = 19.5\nrevenue = units * price\nprint(revenue)", ["Variable", "Type", "Value"], [["units", "int", "12"], ["price", "float", "19.5"], ["revenue", "float", "234.0"]]],
    ["sales = 1200\nprint(\"High\" if sales >= 1000 else \"Standard\")", ["sales", "Condition", "Output"], [["1200", "sales >= 1000", "High"], ["750", "sales >= 1000", "Standard"]]],
    ["for value in [10, 20, 30]:\n    print(value * 2)", ["Iteration", "Input", "Output"], [["1", "10", "20"], ["2", "20", "40"], ["3", "30", "60"]]],
    ["def margin(sales, cost):\n    return (sales-cost)/sales\n\nprint(margin(1000,700))", ["Function Call", "Arguments", "Return"], [["margin(1000,700)", "1000, 700", "0.30"]]],
    ["sales = [100, 150, 200]\nprint(sum(sales))", ["Collection", "Expression", "Result"], [["sales", "sum(sales)", "450"]]],
    ["prices = {'A': 10, 'B': 15}\nprint(prices['B'])", ["Structure", "Input", "Result"], [["dictionary", "prices['B']", "15"]]],
    ["try:\n    open('sales.csv')\nexcept FileNotFoundError:\n    print('File missing')", ["Scenario", "Operation", "Output"], [["File missing", "except FileNotFoundError", "File missing"]]],
    ["import pandas as pd\ndf = pd.DataFrame({'region':['East','West','East'],'sales':[1200,900,1500]})\nprint(df.groupby('region')['sales'].sum())", ["region", "records", "total_sales"], [["East", "2", "2,700"], ["West", "1", "900"]]],
    ["df = pd.read_csv('sales.csv').drop_duplicates()\nsummary = df.groupby('region')['sales'].sum()\nsummary.to_csv('summary.csv')", ["Pipeline Step", "Rows In", "Rows Out", "Artifact"], [["Read CSV", "1,005", "1,005", "DataFrame"], ["Remove duplicates", "1,005", "1,000", "Clean DataFrame"], ["Group region", "1,000", "4", "Summary CSV"]]]
  ]
};

function genericExample(module, index, title) {
  if (module === "Google Sheets") {
    const samples = ["Create a spreadsheet named Team Tracker.", "=A2*$F$1", "=SUM(D2:D6)\n=AVERAGE(D2:D6)", "=TEXTJOIN(\" \",TRUE,A2,B2)\n=TEXT(C2,\"mmmm yyyy\")", "Apply navy headers and currency formatting.", "Sort by Sales descending and filter Region.", "Create a dropdown containing North, South, East, and West.", "Insert a column chart using Month and Sales.", "Share the file with Editor and Viewer permissions.", "Build a task tracker with Owner, Due Date, and Status."];
    return [samples[index], ["Item", "Configuration", "Result"], [[title, "Applied to sample data", "Verified"], ["Quality check", "Manual comparison", "Passed"]]];
  }
  if (module === "Power BI") {
    const samples = ["Define the business question and KPI before building visuals.", "Review Report, Data, and Model views.", "Home > Get data > Excel > Load Sales.xlsx", "Remove blanks, remove duplicates, and set data types in Power Query.", "Create a line chart with Month and Total Sales.", "Add Region and Date slicers, then test interactions.", "Relate Date[Date] to Sales[Date] as one-to-many.", "Total Sales = SUM(Sales[Sales])\nSales YTD = TOTALYTD([Total Sales], 'Date'[Date])", "Apply navy and teal theme, grid alignment, and decision-focused titles.", "Build KPI cards, trend, regional bar chart, and slicers."];
    return [samples[index], ["Power BI Item", "Configuration", "Result"], [[title, "Model or report configured", "Interactive"], ["Filter test", "Region and Date", "Passed"]]];
  }
  if (module === "Tableau") {
    const samples = ["Define the analytical question: Which region drives sales?", "Connect to Orders.xlsx using an Extract.", "Set Region as a Dimension and Sales as a Measure.", "Columns: MONTH(Order Date); Rows: SUM(Sales); Marks: Line", "Drag Category to Filters and select Technology.", "Create a Top 10 Customers set by Sales.", "IF SUM([Sales]) = 0 THEN 0 ELSE SUM([Profit])/SUM([Sales]) END", "Double-click State, then place Sales on Color.", "Place KPI, trend, and regional worksheets on a dashboard.", "Create story points for What changed, Where, and Why."];
    return [samples[index], ["Tableau Item", "Configuration", "Result"], [[title, "Worksheet or dashboard configured", "Verified"], ["Interaction test", "Filter action", "Passed"]]];
  }
  return [`Apply ${title} to the sample dataset.`, ["Lesson", "Action", "Result"], [[title, "Completed", "Verified"]]];
}

function getLessonContent(module, level, index, title) {
  const levelGuidance = {
    Beginner: {
      objective: "Build a correct foundation and understand what each action changes.",
      steps: [
        `Open ${module} and create a small practice workspace for ${title}.`,
        "Enter the sample data exactly as shown and inspect the field names and values.",
        "Apply the formula, query, code, or configuration in the worked example.",
        "Run or refresh the solution and compare every output column with the expected result.",
        "Change one input value and confirm that the output updates in the expected direction."
      ]
    },
    Intermediate: {
      objective: "Apply the technique to a realistic business requirement and handle common data issues.",
      steps: [
        `Translate the business requirement for ${title} into clear technical rules.`,
        "Prepare a multi-table or multi-step sample containing blanks, duplicates, or exceptions.",
        "Build the solution in small, testable stages using the worked example as the starting point.",
        "Validate totals, row counts, filter behavior, and edge cases with a second method.",
        "Package the result as a reusable report, query, function, or documented workflow."
      ]
    },
    Professional: {
      objective: "Design a production-ready solution with performance, security, governance, and maintainability.",
      steps: [
        `Define scale, refresh, security, ownership, and service expectations for ${title}.`,
        "Choose an architecture that separates source data, business logic, presentation, and controls.",
        "Implement the professional pattern shown below with explicit error handling and documentation.",
        "Measure performance and reconcile the result against a trusted control total.",
        "Document deployment, monitoring, recovery, change control, and support responsibilities."
      ]
    }
  };

  const samplePatterns = {
    Excel: {
      Beginner: [
        "Create sheets named Raw_Data, Calculations, and Dashboard.",
        "In C2 enter =A2*B2. In F2 enter =$E$1*C2, then copy both formulas down.",
        "=SUM(D2:D11)\n=AVERAGE(D2:D11)\n=MIN(D2:D11)\n=MAX(D2:D11)",
        "=TEXTJOIN(\" \",TRUE,A2,B2)\n=TEXT(C2,\"mmmm yyyy\")",
        "Format Revenue as Currency, Growth as Percentage, and Order Date as a date.",
        "Filter Region to East, then sort Revenue from Largest to Smallest.",
        "Convert A1:D100 to a table named SalesData.\n=SUM(SalesData[Revenue])",
        "Select Month and Sales, then insert a Line chart titled Monthly Sales Trend.",
        "Apply a red-fill rule to Profit values below 0.",
        "Create Revenue with =[@Units]*[@Price], then add a total row and a Region filter."
      ],
      Intermediate: [
        "=IF(AND(D2>=100000,E2>=0.9),\"On Track\",\"Review\")",
        "=XLOOKUP(A2,Products[ProductID],Products[Price],\"Not found\")",
        "=SUMIFS(Sales[Revenue],Sales[Region],H2,Sales[Date],\">=\"&H3,Sales[Date],\"<=\"&H4)",
        "Create a named range RegionList and apply List validation to Sales[Region].",
        "PivotTable: Region in Rows, Product in Columns, Revenue in Values as Sum.",
        "Insert a PivotChart and connect Region and Product slicers.",
        "Power Query: Trim Customer, replace null Quantity with 0, type Date, remove duplicate OrderID.",
        "=SORT(UNIQUE(FILTER(Sales[Product],Sales[Region]=H2)))",
        "Use Goal Seek: set Profit to 500000 by changing Unit Price.",
        "Combine KPI formulas, PivotTables, charts, a timeline, and slicers on one dashboard."
      ],
      Professional: [
        "=LET(r,Sales[Revenue],reg,Sales[Region],SUM(FILTER(r,reg=H2)))",
        "Create one-to-many relationships from Calendar and Product to the Sales fact table.",
        "Total Sales := SUM(Sales[Revenue])\nSales LY := CALCULATE([Total Sales],SAMEPERIODLASTYEAR(Calendar[Date]))",
        "Table.NestedJoin(Sales,{\"CustomerID\"},Customers,{\"CustomerID\"},\"Customer\",JoinKind.LeftOuter)",
        "Sub RefreshReport()\n  ThisWorkbook.RefreshAll\n  Worksheets(\"Dashboard\").Range(\"B2\").Value = Now\nEnd Sub",
        "Replace volatile full-column formulas with bounded table references and record recalculation time.",
        "Free Cash Flow = EBIT*(1-Tax Rate)+D&A-Capex-Change in NWC",
        "Create Base, Upside, and Downside assumption blocks and select one using CHOOSE/MATCH.",
        "=IFERROR(XLOOKUP(A2,Products[ID],Products[Price]),NA())",
        "Deliver a governed model, KPI dictionary, reconciliation sheet, executive dashboard, and refresh log."
      ]
    },
    SQL: {
      Beginner: [
        "CREATE TABLE customers (customer_id INT PRIMARY KEY, name VARCHAR(100));",
        "SELECT order_id, customer_id, sales FROM orders;",
        "SELECT * FROM orders WHERE sales > 1000;",
        "SELECT * FROM orders ORDER BY sales DESC LIMIT 10;",
        "SELECT quantity * unit_price AS revenue FROM order_items;",
        "SELECT COUNT(*) AS orders, SUM(sales) AS total_sales FROM orders;",
        "SELECT region, SUM(sales) total_sales FROM orders GROUP BY region HAVING SUM(sales)>100000;",
        "SELECT o.order_id,c.name FROM orders o JOIN customers c ON o.customer_id=c.customer_id;",
        "SELECT order_id,CASE WHEN shipped_date IS NULL THEN 'Open' ELSE 'Shipped' END status FROM orders;",
        "SELECT region,COUNT(*) orders,SUM(sales) sales FROM orders GROUP BY region ORDER BY sales DESC;"
      ],
      Intermediate: [
        "SELECT o.order_id,c.name,p.product_name FROM orders o JOIN customers c ON o.customer_id=c.customer_id JOIN products p ON o.product_id=p.product_id;",
        "SELECT * FROM orders WHERE sales > (SELECT AVG(sales) FROM orders);",
        "WITH regional AS (SELECT region,SUM(sales) sales FROM orders GROUP BY region) SELECT * FROM regional WHERE sales>100000;",
        "SELECT region,order_date,sales,SUM(sales) OVER(PARTITION BY region ORDER BY order_date) running_sales FROM orders;",
        "SELECT UPPER(TRIM(customer_name)) clean_name FROM customers;",
        "SELECT DATE_TRUNC('month',order_date) month,SUM(sales) FROM orders GROUP BY 1;",
        "SELECT customer_id FROM web_orders UNION SELECT customer_id FROM store_orders;",
        "UPDATE orders SET status='Closed' WHERE shipped_date IS NOT NULL;",
        "CREATE VIEW monthly_sales AS SELECT DATE_TRUNC('month',order_date) month,SUM(sales) sales FROM orders GROUP BY 1;",
        "WITH x AS (SELECT customer_id,SUM(sales) sales FROM orders GROUP BY customer_id) SELECT *,NTILE(4) OVER(ORDER BY sales DESC) segment FROM x;"
      ],
      Professional: [
        "EXPLAIN ANALYZE SELECT * FROM orders WHERE customer_id=42;",
        "CREATE INDEX idx_orders_customer_date ON orders(customer_id,order_date);",
        "SELECT *,LAG(sales) OVER(PARTITION BY region ORDER BY month) prior_sales FROM monthly_sales;",
        "WITH RECURSIVE org AS (SELECT employee_id,manager_id,1 depth FROM employees WHERE manager_id IS NULL UNION ALL SELECT e.employee_id,e.manager_id,o.depth+1 FROM employees e JOIN org o ON e.manager_id=o.employee_id) SELECT * FROM org;",
        "BEGIN; UPDATE accounts SET balance=balance-100 WHERE id=1; UPDATE accounts SET balance=balance+100 WHERE id=2; COMMIT;",
        "CREATE PROCEDURE close_period(IN p_date DATE) BEGIN UPDATE periods SET status='Closed' WHERE end_date=p_date; END;",
        "CREATE TABLE fact_sales (...); CREATE TABLE dim_customer (...); CREATE TABLE dim_date (...);",
        "GRANT SELECT ON analytics.monthly_sales TO analyst_role;",
        "INSERT INTO fact_sales SELECT ... FROM staging_sales s LEFT JOIN dim_customer c ON s.customer_key=c.customer_key;",
        "BEGIN; CALL validate_stage(:load_id); CALL load_dimensions(:load_id); CALL load_facts(:load_id); COMMIT;"
      ]
    },
    "Google Sheets": {
      Beginner: ["Create tabs named Raw Data, Analysis, and Dashboard.","=A2*$F$1","=SUM(D2:D100)","=TEXTJOIN(\" \",TRUE,A2,B2)&\" | \"&TEXT(C2,\"mmm yyyy\")","Apply currency, percentage, and date formats to matching fields.","Create a filter and sort Sales Z to A.","Apply a dropdown from Settings!A2:A10.","Insert a line chart using Month and Sales.","Share with an Editor and a Viewer, then review version history.","Build a shared tracker with Owner, Due Date, Status, and Notes."],
      Intermediate: ["=IF(AND(D2>=100000,E2>=0.9),\"On Track\",\"Review\")","=XLOOKUP(A2,Products!A:A,Products!D:D,\"Not found\")","=QUERY(A1:D,\"select B,sum(D) group by B label sum(D) 'Sales'\",1)","=SORT(FILTER(A2:D,D2:D>1000),4,FALSE)","Create a pivot table with Region as Rows and SUM of Sales as Values.","Name Settings!B2 StartDate and use =FILTER(A:D,A:A>=StartDate).","=IMPORTRANGE(\"spreadsheet_url\",\"Sales!A1:D100\")","Create a combo chart with Sales columns and Margin % line.","Protect formula columns and allow edits only in input cells.","Combine QUERY summaries, charts, slicers, and KPI cells."],
      Professional: ["=QUERY({Sales!A:D},\"select Col2,sum(Col4) where Col1 is not null group by Col2\",1)","=MAP(A2:A,LAMBDA(id,IF(id=\"\",,XLOOKUP(id,Products!A:A,Products!D:D))))","function onOpen(){SpreadsheetApp.getUi().createMenu('Tools').addItem('Refresh','refresh').addToUi();}","function NETSALES(gross,discount){return gross*(1-discount);}","Create a weekday 7 AM trigger for refreshReport.","Use UrlFetchApp, parse JSON, and write values with setValues in one batch.","Replace repeated full-column formulas with bounded ARRAYFORMULA logic.","Protect source tabs, use Google Groups, and audit sharing monthly.","Connect intake, validation, reporting, and archive sheets with controlled IDs.","Validate submissions, refresh summaries, create a PDF, and log completion with Apps Script."]
    },
    "Power BI": {
      Beginner: ["Define the question: Which regions miss monthly target?","Review Report, Data, and Model views.","Get Data > Excel > select Sales > Transform Data.","Remove blanks, set Date type, and rename Amount to Sales.","Line chart: Month on Axis and Total Sales in Values.","Add a Region slicer and a Sales > 0 visual filter.","Relate Sales[ProductID] many-to-one to Product[ProductID].","Total Sales = SUM(Sales[Sales])","Align visuals to a grid and apply a navy-and-teal theme.","Build KPI cards, monthly trend, region bars, product table, and slicers."],
      Intermediate: ["Model FactSales with DimDate, DimProduct, DimCustomer, and DimRegion.","Sales All Products = CALCULATE([Total Sales],REMOVEFILTERS(Product))","Sales YTD = TOTALYTD([Total Sales],'Date'[Date])","Combine folder files, invoke a transform function, and load only required columns.","Create a Customer drillthrough page and a report-page tooltip.","Create Overview and Detail bookmarks connected to navigation buttons.","[Region] = LOOKUPVALUE(Security[Region],Security[UPN],USERPRINCIPALNAME())","Record the page in Performance Analyzer and inspect DAX query duration.","Publish to a workspace, configure credentials, schedule refresh, and create an app.","Build service level, volume, AHT, staffing, trend, and interval drillthrough views."],
      Professional: ["Create certified shared dimensions and thin reports connected to one semantic model.","Rolling 12M = CALCULATE([Total Sales],DATESINPERIOD('Date'[Date],MAX('Date'[Date]),-12,MONTH))","Create a Time calculation group with Current, MTD, QTD, YTD, and Prior Year.","Keep 5 years and refresh only the last 10 days with incremental refresh.","Use DirectQuery for current transactions and Import for historical aggregates.","Promote content through Development, Test, and Production pipelines.","Endorse the model, document lineage, assign owners, and monitor refresh failures.","Remove unused columns, reduce cardinality, simplify relationships, and benchmark DAX.","Show KPI status, target variance, trend, drivers, exceptions, and recommended action.","Deliver a secured app, deployment process, monitoring plan, and support guide."]
    },
    Tableau: {
      Beginner: ["Ask: How do sales and profit vary by region and month?","Connect to Orders.xlsx and verify field types.","Drag Region to Rows and Sales to Text.","Drag Month to Columns and Sales to Rows; choose Line.","Filter Category to Technology and sort Region by SUM(Sales).","Group low-volume products as Other and create a Top 10 Customer set.","IF SUM([Sales])=0 THEN 0 ELSE SUM([Profit])/SUM([Sales]) END","Assign State a geographic role and place Sales on Color.","Place KPI, trend, and region sheets into a tiled dashboard.","Build story points for performance, drivers, and exceptions."],
      Intermediate: ["RUNNING_SUM(SUM([Sales]))","{ FIXED [Customer ID] : SUM([Sales]) }","Create a Top N parameter and filter with INDEX() <= [Top N].","Relate Orders to Returns by Order ID and preserve each table grain.","Add a filter action from Region chart to Product detail table.","Create distance bands and combine a filled map with proportional symbols.","DATETRUNC('month',[Order Date]) and LOOKUP(SUM([Sales]),-12)","Use consistent spacing, descriptive titles, limited colors, and visible filter state.","Publish to a project, assign permissions, and test as Viewer.","Combine KPI tiles, parameter-driven Top N, action filters, and detail-on-demand."],
      Professional: ["{ FIXED [Customer ID] : MIN([First Order Date]) }","WINDOW_SUM(SUM([Sales]),-2,0)","Record performance and inspect query, layout, and rendering events.","Use extracts for history and governed published sources for shared dimensions.","Use parameters and show/hide containers to switch metrics and layouts.","MODEL_QUANTILE(0.95,SUM([Sales]),COUNT([Orders]))","Certify shared sources, document calculations, assign owners, and review permissions.","Embed a trusted workbook with user identity and permission enforcement.","Sequence context, performance, drivers, exceptions, and recommended action.","Deliver governed sources, optimized workbooks, secure projects, monitoring, and publishing guide."]
    },
    Python: {
      Beginner: ["print('Hello, data!')","units=12\nprice=19.5\nrevenue=units*price\nprint(revenue)","sales=1200\nprint('High' if sales>=1000 else 'Standard')","for n in [10,20,30]:\n    print(n*2)","def margin(revenue,cost):\n    return (revenue-cost)/revenue\nprint(margin(1000,700))","sales=[100,150,200]\nprint(sum(sales),sales[0])","prices={'A':10,'B':15}\nregions=set(['East','East','West'])\nprint(prices['B'],regions)","try:\n    with open('sales.txt') as f: print(f.read())\nexcept FileNotFoundError:\n    print('File missing')","import pandas as pd\ndf=pd.DataFrame({'region':['East','West'],'sales':[100,150]})\nprint(df['sales'].sum())","df=pd.read_csv('sales.csv')\nprint(df.groupby('region')['sales'].sum())"],
      Intermediate: ["class Sale:\n    def __init__(self,qty,price): self.qty,self.price=qty,price\n    def revenue(self): return self.qty*self.price","squares=[n*n for n in range(5)]\ngen=(n*n for n in range(5))","from pathlib import Path\nprint(Path.cwd())","import numpy as np\na=np.array([1,2,3])\nprint(a.mean())","summary=df.groupby('region',as_index=False).agg(sales=('sales','sum')).sort_values('sales',ascending=False)","df['date']=pd.to_datetime(df['date'],errors='coerce')\ndf=df.drop_duplicates('order_id').fillna({'sales':0})","import matplotlib.pyplot as plt\ndf.groupby('month')['sales'].sum().plot()\nplt.show()","import re\nprint(re.findall(r'[A-Z]{2}-\\d{4}','IDs: AB-1234'))","def add_tax(x): return round(x*1.12,2)\nassert add_tax(100)==112.0","df=pd.read_csv('sales.csv').drop_duplicates()\ndf.to_csv('clean_sales.csv',index=False)"],
      Professional: ["Separate ingestion, validation, transformation, and reporting into modules with explicit interfaces.","from functools import lru_cache\n@lru_cache\ndef rate(code): return expensive_lookup(code)","def total(values: list[float]) -> float:\n    return sum(values)","import cProfile\ncProfile.run('build_report()')","from concurrent.futures import ThreadPoolExecutor\nwith ThreadPoolExecutor() as ex: results=list(ex.map(fetch,urls))","from sqlalchemy import text\nwith engine.begin() as c: rows=c.execute(text('SELECT * FROM sales WHERE date>=:d'),{'d':start})","@app.get('/sales/{region}')\ndef sales(region:str): return service.by_region(region)","def test_total():\n    assert total([10,20])==30","python -m build\npip install dist/analytics_app-1.0.0-py3-none-any.whl","Package a validated data service, database layer, tests, logging, configuration, and deployment definition."]
    }
  };

  const sample = samplePatterns[module][level][index];
  const outputs = {
    Beginner: [["Lesson", "Applied Skill", "Result"], [[title, "Foundation workflow", "Correct and verified"], ["Manual check", "One sample record", "Passed"]]],
    Intermediate: [["Business Requirement", "Implementation", "Validation"], [[title, "Reusable applied solution", "Control total matched"], ["Edge cases", "Blanks and duplicates tested", "Passed"]]],
    Professional: [["Production Control", "Implementation", "Status"], [[title, "Governed scalable solution", "Ready for review"], ["Performance and security", "Measured and documented", "Passed"]]]
  };
  const [columns, rows] = outputs[level];
  return {
    explanation: `${title} is a ${level.toLowerCase()} ${module} capability. ${levelGuidance[level].objective} This lesson focuses on a different skill and deliverable from the other levels, so the example, process, and outcome are specific to ${title}.`,
    steps: levelGuidance[level].steps,
    sample,
    columns,
    rows,
    result: `${title} produces a ${level.toLowerCase()}-appropriate ${module} deliverable that has been validated against the expected outcome.`
  };
}

export default function App() {
  const [selected, setSelected] = useState(null);
  const [showDashboard, setShowDashboard] = useState(false);
  const [moduleFilter, setModuleFilter] = useState("All");
  const [levelFilter, setLevelFilter] = useState("All");
  const [query, setQuery] = useState("");
  const [completed, setCompleted] = useState(() => { try { return JSON.parse(localStorage.getItem("dataskills-progress") || "[]"); } catch { return []; } });
  useEffect(() => localStorage.setItem("dataskills-progress", JSON.stringify(completed)), [completed]);

  const lessonId = lesson => `${lesson.module}|${lesson.level}|${lesson.index}`;
  const toggleComplete = lesson => setCompleted(current => current.includes(lessonId(lesson)) ? current.filter(id => id !== lessonId(lesson)) : [...current, lessonId(lesson)]);
  const lessons = useMemo(() => Object.entries(modules).flatMap(([module, data]) => LEVELS.flatMap(level => data.lessons[level].map((title, index) => ({module, level, title, index})))), []);
  const filtered = lessons.filter(item => (moduleFilter === "All" || item.module === moduleFilter) && (levelFilter === "All" || item.level === levelFilter) && `${item.module} ${item.level} ${item.title}`.toLowerCase().includes(query.toLowerCase()));
  const totalLessons = lessons.length;
  const completedCount = completed.length;
  const overallPercent = totalLessons ? Math.round(completedCount / totalLessons * 100) : 0;
  const moduleProgress = Object.keys(modules).map(name => { const total = 30; const done = completed.filter(id => id.startsWith(`${name}|`)).length; return {name, total, done, percent: Math.round(done/total*100)}; });
  const completedLevels = Object.keys(modules).flatMap(name => LEVELS.map(level => completed.filter(id => id.startsWith(`${name}|${level}|`)).length)).filter(done => done === 10).length;
  const levelProgress = LEVELS.map(level => { const total = 60; const done = completed.filter(id => id.split("|")[1] === level).length; return {level, total, done, percent: Math.round(done/total*100)}; });
  const strongestModule = [...moduleProgress].sort((a,b)=>b.percent-a.percent)[0] || {name:"No module yet",percent:0};
  const nextMilestone = overallPercent < 25 ? 25 : overallPercent < 50 ? 50 : overallPercent < 75 ? 75 : 100;
  const levelStyle = {Beginner:"bg-emerald-50 text-emerald-700 ring-emerald-200", Intermediate:"bg-amber-50 text-amber-700 ring-amber-200", Professional:"bg-violet-50 text-violet-700 ring-violet-200"};

  if (showDashboard) return <main className="min-h-screen bg-[#F4F7FA] text-slate-900">
    <header className="sticky top-0 z-40 border-b bg-white/95 shadow-sm backdrop-blur-xl"><div className="mx-auto flex max-w-7xl items-center justify-between px-4 py-4 sm:px-6"><button onClick={()=>setShowDashboard(false)} className="flex items-center gap-3"><span className="grid h-10 w-10 place-items-center rounded-xl bg-gradient-to-br from-teal-400 to-teal-600 text-white"><GraduationCap size={22}/></span><span className="text-left"><span className="block font-black text-[#0A2342]">DataSkills Academy</span><span className="block text-xs text-slate-500">Progress Dashboard</span></span></button><button onClick={()=>setShowDashboard(false)} className="rounded-xl bg-[#0A2342] px-5 py-2.5 text-sm font-bold text-white">Back to courses</button></div></header>
    <section className="bg-gradient-to-br from-[#071A2F] via-[#0A2B4D] to-[#07535D] text-white"><div className="mx-auto max-w-7xl px-4 py-14 sm:px-6"><p className="text-xs font-bold uppercase tracking-[.2em] text-teal-300">Study dashboard</p><h1 className="mt-3 text-4xl font-black sm:text-5xl">Your learning progress</h1><p className="mt-4 text-slate-300">Track completed lessons and progress across every learning module.</p></div></section>
    <section className="mx-auto max-w-7xl px-4 py-10 sm:px-6">
      <div className="grid gap-5 lg:grid-cols-[1.2fr_.8fr]"><div className="rounded-3xl bg-gradient-to-br from-[#0A2342] to-[#07535D] p-8 text-white shadow-xl"><div className="flex items-start justify-between"><div><p className="font-semibold text-teal-200">Overall completion</p><p className="mt-2 text-6xl font-black">{overallPercent}%</p><p className="mt-3 text-slate-300">{completedCount} of {totalLessons} lessons completed</p></div><Trophy className="text-teal-300" size={42}/></div><div className="mt-8 h-4 overflow-hidden rounded-full bg-white/10"><div className="h-full rounded-full bg-gradient-to-r from-teal-300 to-emerald-400" style={{width:`${overallPercent}%`}}/></div></div><div className="grid grid-cols-2 gap-4"><div className="rounded-3xl border bg-white p-6"><CheckCircle2 className="text-emerald-500"/><p className="mt-4 text-4xl font-black">{completedCount}</p><p className="text-sm text-slate-500">Completed</p></div><div className="rounded-3xl border bg-white p-6"><Trophy className="text-amber-500"/><p className="mt-4 text-4xl font-black">{completedLevels}</p><p className="text-sm text-slate-500">Levels finished</p></div><div className="col-span-2 rounded-3xl border bg-white p-6"><p className="text-sm text-slate-500">Lessons remaining</p><p className="mt-2 text-4xl font-black">{totalLessons-completedCount}</p></div></div></div>
      <div className="mt-8 grid gap-5 lg:grid-cols-3"><div className="rounded-3xl border bg-white p-7"><p className="text-xs font-bold uppercase tracking-wider text-teal-600">Completion chart</p><h2 className="mt-2 text-xl font-black">Overall progress</h2><div className="mt-6 flex justify-center"><div className="grid h-48 w-48 place-items-center rounded-full" style={{background:`conic-gradient(#14b8a6 ${overallPercent*3.6}deg,#e2e8f0 0deg)`}}><div className="grid h-36 w-36 place-items-center rounded-full bg-white"><div className="text-center"><p className="text-4xl font-black">{overallPercent}%</p><p className="text-xs uppercase text-slate-400">Complete</p></div></div></div></div><div className="mt-6 flex justify-between rounded-2xl bg-slate-50 p-4 text-sm"><span>Next milestone</span><b className="text-teal-700">{nextMilestone}%</b></div></div><div className="rounded-3xl border bg-white p-7 lg:col-span-2"><p className="text-xs font-bold uppercase tracking-wider text-teal-600">Level comparison</p><h2 className="mt-2 text-xl font-black">Progress by skill level</h2><div className="mt-7 space-y-6">{levelProgress.map((item,i)=>{const colors=["bg-emerald-500","bg-amber-500","bg-violet-500"];return <div key={item.level}><div className="mb-2 flex justify-between"><div><b>{item.level}</b><p className="text-xs text-slate-500">{item.done} of {item.total} lessons</p></div><b>{item.percent}%</b></div><div className="h-3 rounded-full bg-slate-100"><div className={`h-full rounded-full ${colors[i]}`} style={{width:`${item.percent}%`}}/></div></div>})}</div><div className="mt-7 grid grid-cols-2 gap-3"><div className="rounded-2xl bg-teal-50 p-4"><p className="text-xs font-bold uppercase text-teal-700">Leading module</p><p className="mt-2 text-lg font-black">{strongestModule.name}</p><p className="text-sm text-slate-500">{strongestModule.percent}% complete</p></div><div className="rounded-2xl bg-slate-50 p-4"><p className="text-xs font-bold uppercase text-slate-500">Coverage</p><p className="mt-2 text-lg font-black">{moduleProgress.filter(x=>x.done>0).length} of 6</p><p className="text-sm text-slate-500">modules started</p></div></div></div></div>
      <div className="mt-10 flex justify-between"><div><h2 className="text-2xl font-black">Progress by module</h2><p className="text-sm text-slate-500">Each module contains 30 lessons.</p></div><button onClick={()=>setCompleted([])} disabled={!completedCount} className="flex items-center gap-2 rounded-xl border bg-white px-4 py-2 text-sm font-bold disabled:opacity-40"><RotateCcw size={16}/>Reset</button></div><div className="mt-6 grid gap-5 md:grid-cols-2 xl:grid-cols-3">{moduleProgress.map(item=>{const Icon=modules[item.name].icon;return <div key={item.name} className="rounded-3xl border bg-white p-6"><div className="flex justify-between"><span className="grid h-11 w-11 place-items-center rounded-xl bg-teal-50 text-teal-700"><Icon size={21}/></span><b className="text-xl text-teal-700">{item.percent}%</b></div><h3 className="mt-5 text-xl font-black">{item.name}</h3><p className="text-sm text-slate-500">{item.done} of {item.total} lessons</p><div className="mt-5 h-2.5 rounded-full bg-slate-100"><div className="h-full rounded-full bg-teal-500" style={{width:`${item.percent}%`}}/></div></div>})}</div>
    </section>
  </main>;

  if (selected) { const lesson=getLessonContent(selected.module,selected.level,selected.index,selected.title); const titles=modules[selected.module].lessons[selected.level]; return <main className="min-h-screen bg-[#F4F7FA]">
    <header className="sticky top-0 z-40 border-b bg-white/95 shadow-sm backdrop-blur"><div className="mx-auto flex max-w-7xl items-center justify-between px-4 py-3"><button onClick={()=>setSelected(null)} className="flex items-center gap-3"><span className="grid h-10 w-10 place-items-center rounded-xl bg-teal-500 text-white"><GraduationCap size={22}/></span><span className="text-left"><b className="block text-[#0A2342]">DataSkills Academy</b><span className="text-xs text-slate-500">Learn. Practice. Build.</span></span></button><button onClick={()=>setSelected(null)} className="flex items-center gap-2 rounded-xl border px-4 py-2"><X size={16}/>Close lesson</button></div></header>
    <div className="mx-auto grid max-w-7xl gap-8 px-4 py-8 lg:grid-cols-[260px_1fr]"><aside className="hidden lg:block"><div className="sticky top-24 overflow-hidden rounded-3xl bg-[#0A2342] p-3 text-white"><div className="p-3"><p className="text-xs uppercase tracking-wider text-teal-300">{selected.module}</p><h2 className="mt-2 text-xl font-black">{selected.level} path</h2></div>{titles.map((title,index)=><button key={title} onClick={()=>setSelected({...selected,title,index})} className={`mb-1 w-full rounded-xl px-3 py-2 text-left text-sm ${index===selected.index?"bg-teal-500":"hover:bg-white/10"}`}>{index+1}. {title}</button>)}</div></aside><article className="min-w-0"><div className="rounded-3xl bg-gradient-to-br from-[#0A2342] to-[#07535D] p-8 text-white"><div className="flex gap-2 text-xs"><span className="rounded-full bg-white/10 px-3 py-1">{selected.module}</span><span className="rounded-full bg-teal-400/20 px-3 py-1">{selected.level}</span></div><h1 className="mt-6 text-4xl font-black">{selected.title}</h1><p className="mt-4 text-slate-200">Learn the concept, follow the process, study a practical example, and compare your work with the expected outcome.</p></div>
    <section className="mt-6 rounded-3xl border bg-white p-7"><h2 className="flex items-center gap-3 text-2xl font-black"><BookOpen className="text-teal-600"/>Concept explanation</h2><p className="mt-5 leading-8 text-slate-600">{lesson.explanation}</p></section>
    <section className="mt-6 rounded-3xl border bg-white p-7"><h2 className="text-2xl font-black">Step-by-step process</h2><ol className="mt-6 space-y-4">{lesson.steps.map((step,i)=><li key={step} className="flex gap-4 rounded-2xl bg-slate-50 p-4"><span className="grid h-9 w-9 shrink-0 place-items-center rounded-xl bg-[#0A2342] text-white">{i+1}</span><p className="pt-1.5">{step}</p></li>)}</ol></section>
    <section className="mt-6 overflow-hidden rounded-3xl bg-[#071A2F] text-white"><div className="border-b border-white/10 p-5"><h2 className="flex items-center gap-2 font-bold"><Code2 className="text-teal-300"/>Example code or action</h2></div><pre className="overflow-x-auto whitespace-pre-wrap p-7 text-sm leading-7 text-teal-100"><code>{lesson.sample}</code></pre><div className="border-t border-white/10 bg-white/5 p-6"><p className="text-xs uppercase text-teal-300">What this produces</p><p className="mt-2 text-slate-200">{lesson.result}</p></div></section>
    <section className="mt-6 overflow-hidden rounded-3xl border bg-white"><div className="border-b bg-slate-50 p-6"><h2 className="text-2xl font-black">Example outcome</h2></div><div className="overflow-x-auto"><table className="w-full text-left text-sm"><thead className="bg-[#0A2342] text-white"><tr>{lesson.columns.map(x=><th key={x} className="px-6 py-4">{x}</th>)}</tr></thead><tbody>{lesson.rows.map((row,i)=><tr key={i} className="border-b even:bg-slate-50">{row.map((x,j)=><td key={j} className="px-6 py-4">{x}</td>)}</tr>)}</tbody></table></div></section>
    <button onClick={()=>toggleComplete(selected)} className={`mt-6 flex w-full items-center justify-center gap-2 rounded-2xl px-6 py-4 font-black text-white ${completed.includes(lessonId(selected))?"bg-emerald-600":"bg-teal-600"}`}><CheckCircle2 size={20}/>{completed.includes(lessonId(selected))?"Lesson completed - click to undo":"Mark lesson as complete"}</button>
    <div className="mt-8 flex justify-between"><button disabled={selected.index===0} onClick={()=>{const i=selected.index-1;setSelected({...selected,index:i,title:titles[i]});window.scrollTo(0,0)}} className="rounded-xl border bg-white px-5 py-3 disabled:opacity-40">Previous lesson</button><button disabled={selected.index===9} onClick={()=>{const i=selected.index+1;setSelected({...selected,index:i,title:titles[i]});window.scrollTo(0,0)}} className="rounded-xl bg-[#0A2342] px-5 py-3 text-white disabled:opacity-40">Next lesson</button></div></article></div></main> }

  return <main className="min-h-screen bg-[#F4F7FA]">
    <header className="sticky top-0 z-40 border-b border-white/10 bg-[#071A2F]/95 text-white backdrop-blur"><div className="mx-auto flex max-w-7xl items-center justify-between px-4 py-4"><div className="flex items-center gap-3"><span className="grid h-10 w-10 place-items-center rounded-xl bg-teal-500"><GraduationCap size={22}/></span><div><b>DataSkills Academy</b><p className="text-xs text-slate-400">Manual by <span className="text-teal-300">Jayson Lofranco</span></p></div></div><div className="flex gap-3"><button onClick={()=>{setShowDashboard(true);window.scrollTo(0,0)}} className="flex items-center gap-2 rounded-xl bg-teal-400 px-4 py-2.5 font-black text-[#06233C]"><LayoutDashboard size={17}/>Dashboard</button><span className="hidden rounded-full border border-white/10 px-3 py-2 text-xs lg:block">180 structured lessons</span></div></div></header>
    <section className="relative overflow-hidden bg-gradient-to-br from-[#071A2F] via-[#0A2B4D] to-[#07535D] text-white"><div className="absolute -right-20 -top-20 h-80 w-80 rounded-full bg-teal-400/20 blur-3xl"/><div className="relative mx-auto grid max-w-7xl gap-12 px-4 py-20 lg:grid-cols-[1.05fr_.95fr]"><div><span className="rounded-full bg-teal-400/15 px-4 py-2 text-xs font-bold uppercase tracking-wider text-teal-200">Learn from beginner to professional</span><h1 className="mt-6 text-5xl font-black leading-tight">Build skills that turn<br/><span className="text-teal-300">data into decisions.</span></h1><p className="mt-6 max-w-2xl text-lg leading-8 text-slate-300">A practical academy for Excel, SQL, Google Sheets, Power BI, Tableau, and Python. Every lesson includes guided steps, examples, and visible outcomes.</p><button onClick={()=>document.getElementById('catalog')?.scrollIntoView()} className="mt-8 rounded-xl bg-teal-400 px-6 py-3.5 font-bold text-[#06233C]">Explore the curriculum</button></div><div className="grid grid-cols-2 gap-3">{Object.entries(modules).map(([name,data])=>{const Icon=data.icon;return <button key={name} onClick={()=>{setModuleFilter(name);setTimeout(()=>document.getElementById('catalog')?.scrollIntoView(),20)}} className="rounded-2xl border border-white/10 bg-white/[.07] p-5 text-left hover:bg-white/[.12]"><span className="grid h-11 w-11 place-items-center rounded-xl bg-teal-400/15 text-teal-300"><Icon size={22}/></span><b className="mt-4 block">{name}</b><p className="text-xs text-slate-400">30 lessons · 3 levels</p></button>})}</div></div></section>
    <section id="catalog" className="mx-auto max-w-7xl px-4 py-14"><div className="flex justify-between"><div><p className="text-xs font-bold uppercase tracking-wider text-teal-600">Full curriculum</p><h2 className="mt-2 text-4xl font-black">Choose your learning path</h2><p className="mt-2 text-slate-500">Search or filter all lessons.</p></div><p className="text-sm text-slate-500">Showing <b>{filtered.length}</b> lessons</p></div><div className="mt-8 grid gap-3 rounded-2xl border bg-white p-3 md:grid-cols-[1fr_auto_auto]"><div className="relative"><Search className="absolute left-4 top-3.5 text-slate-400" size={20}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search lessons..." className="w-full rounded-xl border bg-slate-50 py-3 pl-12 pr-4"/></div><select value={moduleFilter} onChange={e=>setModuleFilter(e.target.value)} className="rounded-xl border px-4"><option>All</option>{Object.keys(modules).map(x=><option key={x}>{x}</option>)}</select><select value={levelFilter} onChange={e=>setLevelFilter(e.target.value)} className="rounded-xl border px-4"><option>All</option>{LEVELS.map(x=><option key={x}>{x}</option>)}</select></div><div className="mt-7 grid gap-5 md:grid-cols-2 xl:grid-cols-3">{filtered.map(lesson=>{const Icon=modules[lesson.module].icon;return <button key={`${lesson.module}-${lesson.level}-${lesson.index}`} onClick={()=>{setSelected(lesson);window.scrollTo(0,0)}} className="group rounded-3xl border bg-white p-6 text-left shadow-sm hover:-translate-y-1 hover:shadow-xl"><div className="flex justify-between"><span className="grid h-12 w-12 place-items-center rounded-2xl bg-slate-100"><Icon size={23}/></span><span className="text-xs text-slate-300">{String(lesson.index+1).padStart(2,'0')}</span></div><div className="mt-5 flex gap-2"><span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs">{lesson.module}</span><span className={`rounded-full px-2.5 py-1 text-xs ring-1 ${levelStyle[lesson.level]}`}>{lesson.level}</span></div><h3 className="mt-4 text-xl font-black group-hover:text-teal-700">{lesson.title}</h3><p className="mt-2 text-sm text-slate-500">Concept, guided process, example, and outcome.</p><p className="mt-5 text-sm font-bold text-teal-700">Open lesson →</p></button>})}</div></section>
    <footer className="border-t bg-white"><div className="mx-auto flex max-w-7xl justify-between px-4 py-8 text-sm text-slate-500"><b className="text-[#0A2342]">DataSkills Academy</b><p>Practical learning from foundations to professional workflows.</p></div></footer>
  </main>;
}
