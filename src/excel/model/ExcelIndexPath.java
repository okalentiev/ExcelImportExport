package excel.model;

/**
 * Created by Heach on 25.04.2014.
 */

public class ExcelIndexPath {
    private int row;
    private int cell;

    public ExcelIndexPath(int row, int cell)
    {
        this.row = row;
        this.cell = cell;
    }

    public int getRow()
    {
        return this.row;
    }

    public int getCell()
    {
        return this.cell;
    }
}
