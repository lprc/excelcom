# excelcom
Easy to use excel modification library using JNA and COM. Works on Windows only.
Covers only really basic operations such as reading and writing content and coloring.

excelcom requires Java Version 1.6 or higher and the COM service (normally included with 
an office installation).

## How to import
Add the following dependency to your `pom.xml`:

    <dependencies>
        <dependency>
            <groupId>com.github.lprc</groupId>
            <artifactId>excelcom</artifactId>
            <version>0.0.5</version>
        </dependency>
    </dependencies>


## How to use

     import excelcom.api.*;
     import excelcom.util.Util;
     
     // note that conn.quit() MUST be called later for uninitializing COM correctly.
     // otherwise opened excel proecesses will remain in task manager.
     try {
         // connect to a new excel instance and don't show any dialogs
         ExcelConnection conn = ExcelConnection.connect();
         conn.setDisplayAlerts(false);
         
        
         // open a workbook
         Workbook wb = conn.openWorkbook(new File("test.xlsx"));
         
        
         // open a worksheet
         Worksheet ws = wb.getWorksheet("Tabelle1");
         
        
         // write some content, mutliple cell range and unary cell range
         ws.setContent("A4:B5", new Object[][]{ {123, 456.5}, {"test", "äöüß"} });
         ws.setContent("A6", 432.4f);
         
        
         // read content
         Util.printMatrix(ws.getContent("A4:B6"));
         System.out.println(ws.getUnaryContent("A4"));
         
        
         // colorize some cells
         ws.setFillColor("A4", ExcelColor.LIGHT_GREEN);
         ws.setFontColor("A5", ExcelColor.RED);
         
        
         // attach some comments (works for one cell only)
         ws.setComment("A6", "test comment");
         
        
         // save and close workbook
         wb.close(true);
     } finally {
         // quit excel instance and uninitialize COM
         conn.quit();
     }

## Known problems
Since COM doesn't provide exact failure descriptions and calling the
 same COM function can have multiple return types, there a some tradeoffs:

- Setting a comment to a cell works for one cell at a time only. 
Calling `Worksheet#setComment` with a multiple cell range throws an
`IllegalArgumentException`.

- Getting a color from multiple cell range which has different colors of the same type
(e.g. different fill colors) results in a `NullPointerException`.

Improvements to the code of any kind are very welcome!

