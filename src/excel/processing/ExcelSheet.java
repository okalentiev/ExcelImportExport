package excel.processing;

/**
 * Created by Heach on 25.05.2014.
 * Class excel sheet. Used to create sheet and manage data.
 */

import excel.templates.StudyPlanCell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

public class ExcelSheet {
    private Sheet currentSheet;

    public ExcelSheet(Sheet currentSheet)
    {
        this.currentSheet = currentSheet;
    }

    public void createRowWithCells(int rowIndex, ArrayList<StudyPlanCell> studyPlanCells)
    {
        Row currentRow = currentSheet.createRow(rowIndex);
        for (StudyPlanCell cell: studyPlanCells) {

            Cell currentCell = currentRow.createCell(cell.getIndexPath().getCell());

            currentCell.setCellValue(cell.getCellValue());
            currentCell.setCellStyle(cell.getCellStyle());

            createMergedRegion(new CellRangeAddress(cell.getIndexPath().getRow(),
                    cell.getIndexPath().getRow() + cell.getCellRange().getHeight(),
                    cell.getIndexPath().getCell(),
                    cell.getIndexPath().getCell() + cell.getCellRange().getWidth()), cell.getCellStyle());
        }
    }

    private void cleanBeforeMergeOnValidCells(CellRangeAddress region, CellStyle cellStyle )
    {
        Row tempRow = currentSheet.getRow(region.getFirstRow());
        Cell cell = tempRow.getCell(region.getFirstColumn());

        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            tempRow = currentSheet.getRow(i);

            for (int j = region.getFirstColumn(); j <= region.getLastColumn()  ; j++) {
                tempRow.getCell(j).setCellStyle(cellStyle);
            }
        }
    }

    public void createMergedRegion(CellRangeAddress cellRangeAddress, CellStyle cellStyle)
    {
//        cleanBeforeMergeOnValidCells(cellRangeAddress, cellStyle);
        currentSheet.addMergedRegion(cellRangeAddress);
    }

    public Sheet getCurrentSheet ()
    {
        return this.currentSheet;
    }
}
