package controllers;

import excel.processing.ExcelExportManager;

/**
 * Created by Heach on 19.05.2014.
 * Test class to load excel import export
 */
public class MainWindowController {
    public static void main (String args[])
    {
        ExcelExportManager exportManager = new ExcelExportManager();

        exportManager.createTableHeader();
        exportManager.writeWorkbookToFile("d:/tempExcelPOI/ExportedExcel.xlsx");
    }
}
