package excel.processing;

import excel.templates.StudyPlanCell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Created by Heach on 20.05.2014.
 * Class, that managing excell file,
 * creating sheet, cell and filling them with data;
 */
public class ExcelExportManager {
    private SXSSFWorkbook workbook;
    private ExcelSheet excelSheet;

    public ExcelExportManager()
    {
        initializeWorkbook("Лист 1");
    }

    public void initializeWorkbook (String sheetName)
    {
        workbook =  new SXSSFWorkbook(SXSSFWorkbook.DEFAULT_WINDOW_SIZE);

        Sheet sheet = workbook.createSheet(sheetName);
        CellStyle cellStyle = workbook.createCellStyle();

        excelSheet = new ExcelSheet(sheet);
    }

    public void createTableHeader ()
    {
        String[] rowTitles = {"№п/п", "Назва дисципліни", "Розподіл по семестрам", "Кредити ECTS", "Години", "Розподіл по курсам та семестрам", "Розподіл по курсам та семестрам"};
        int[] cellIndexes = {0,1,19,25,26,33,37};
        int[][] cellRanges = {{0,7}, {17,7}, {5,2}, {0,7}, {6, 0}, {3, 0}, {3, 0}};


        ArrayList<StudyPlanCell> studyPlanCells = new ArrayList<StudyPlanCell>();
        CellStyle cellStyle = workbook.createCellStyle();

        for (String cellValue : rowTitles)
        {
            int cellIndex = studyPlanCells.size();
            StudyPlanCell studyPlanCell = new StudyPlanCell(cellStyle);
            studyPlanCell.setIndexPath(0, cellIndexes[cellIndex]);
            studyPlanCell.setCellRange(cellRanges[cellIndex]);
            studyPlanCell.setCellValue(rowTitles[cellIndex]);

            studyPlanCells.add(studyPlanCell);
        }

        excelSheet.createRowWithCells(0, studyPlanCells);
    }

    public void writeWorkbookToFile (String fileName)
    {
        File f = new File(fileName);

        if(!f.exists()) {
            //If directories are not available then create it
            File parentDirectory = f.getParentFile();
            if (null != parentDirectory) {
                parentDirectory.mkdirs();
            }
            try {
                f.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(f,false);
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        workbook.dispose();
    }
}
