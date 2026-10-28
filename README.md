# ExcelToJson - Simple Beginner Project

This is a very simple C# .NET 10 console application.

It does:

Excel file
→ Read rows
→ Convert rows to C# Dictionary objects
→ Convert objects to JSON
→ Save JSON file

## Run

```bash
dotnet restore
dotnet run
```

When asked for the Excel path, enter for example:

```text
/Users/amans/Desktop/my .Net Journey/products.xlsx
```

The JSON file will be created next to the Excel file.

## Excel format

The first row must contain column names:

| Id | Name | Price | Category |
|---|---|---|---|
| 1 | Laptop | 75000 | Electronics |
| 2 | Mouse | 1200 | Accessories |
| 3 | Keyboard | 2500 | Accessories |
