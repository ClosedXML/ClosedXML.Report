namespace ClosedXML.Report.Excel;

struct CellPosition
{
    public CellPosition(int row, int column)
    {
        Row = row;
        Column = column;
    }

    public int Row { get; set; }

    public int Column { get; set; }
}