# MergeExcel

MergeExcel is a small programe that allow you can merge Multiple Excel files into one using Microsoft Interop

the basic approach :

* Initialize Excel Application to interact with Excel.

* Open each workbook using the Workbooks.Open() method.

* Copy Data from its sheets and paste it into the target workbook.

* Close the source workbooks After merging.
  
![1](https://github.com/user-attachments/assets/09ad5b1b-a21c-4d70-ab76-e6ebf50bf9e4)

i have a small issue :

![984192](https://github.com/user-attachments/assets/5f29cad1-37fb-4ddb-9891-e79f10c6ce9a)


This message showing when the Excel data is copied to the clipboard is likely a result of Excel's internal notification system that pops up when a clipboard operation is performed (e.g., when you copy data).

To suppress these notifications or avoid the clipboard message :

> Application.DisplayAlerts = false; 

This will suppress any Excel alert dialogs, including the one that may show when copying to the clipboard. :tada:


[Microsoft Interop](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia)
