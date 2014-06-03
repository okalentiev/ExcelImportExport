package excel.templates;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Created by Heach on 23.05.2014.
 */
public class StudyPlanCell extends ExcelCell {

    public StudyPlanCell(CellStyle cellStyle)
    {
        cellStyle.setBorderBottom(CellStyle.BORDER_MEDIUM);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(CellStyle.BORDER_MEDIUM);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(CellStyle.BORDER_MEDIUM);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

        setCellStyle(cellStyle);
    }

}
