package excel.model;

/**
 * Created by Heach on 25.05.2014.
 * Note: width anf height it is number of cells
 */
public class ExcelCellRange {
    private int width;
    private int height;

    public ExcelCellRange(int width, int height)
    {
        this.width = width;
        this.height = height;
    }

    public int getWidth()
    {
        return this.width;
    }

    public int getHeight()
    {
        return this.height;
    }

    public void setWidth(int width)
    {
        this.width = width;
    }

    public void setHeight(int height)
    {
        this.height = height;
    }
}
