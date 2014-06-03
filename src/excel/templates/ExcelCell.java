package excel.templates;

import excel.model.ExcelCellRange;
import excel.model.ExcelIndexPath;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Created by Heach on 23.05.2014.
 */

public class ExcelCell {
    private String cellValue;
    private CellStyle cellStyle;
    private ExcelIndexPath indexPath;
    private ExcelCellRange cellRange;

    public void setCellValue(String cellValue)
    {
        this.cellValue = cellValue;
    }

    public String getCellValue ()
    {
        return this.cellValue;
    }

    public void setCellStyle(CellStyle cellStyle)
    {
        this.cellStyle = cellStyle;
    }

    public CellStyle getCellStyle()
    {
        return this.cellStyle;
    }

    public void setIndexPath (ExcelIndexPath indexPath)
    {
        this.indexPath = indexPath;
    }

    public void setIndexPath (int row, int cell)
    {
        this.indexPath = new ExcelIndexPath(row, cell);
    }

    public ExcelIndexPath getIndexPath()
    {
        return this.indexPath;
    }
    
    public void setCellRange (ExcelCellRange cellRange)
    {
        this.cellRange = cellRange;
    }

    public void setCellRange(int width, int height)
    {
        setCellRange(new ExcelCellRange(width, height));
    }

    public void setCellRange(int[] range)
    {
        if (range.length > 0)
            setCellRange(new ExcelCellRange(range[0], range[1]));
    }

    public ExcelCellRange getCellRange ()
    {
        return this.cellRange;
    }
}
