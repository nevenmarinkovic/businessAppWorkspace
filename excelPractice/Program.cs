// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using System.IO;

FileInfo fileInfo = new FileInfo(@"C:\\Users\\nemar\\OneDrive\\Documents\\Me\\Time_Tracking_Template.xlsx");
        using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
        {
            // Get the first worksheet
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
            //Console.WriteLine(worksheet.Name);
            for(int i = 1; i < 5; i++)
            {
                for(int j = 1; j < 13; j++)
                {
                    object cellValue = worksheet.Cells[j, i].Value;
                    Console.WriteLine(cellValue);
                }
                Console.WriteLine();
            }
           
        }


/*
class Program
{
    static void Main(string[] args)
    {
        // Load the Excel file
        FileInfo fileInfo = new FileInfo(@"C:\\Users\\nemar\\OneDrive\\Documents\\Me\\Time_Tracking_Template.xlsx");
        using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
        {
            // Get the first worksheet
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
            Console.WriteLine(worksheet.Name);
            
            // Get the dimension of the worksheet
            ExcelCellAddress startCell = worksheet.Dimension.Start;
            ExcelCellAddress endCell = worksheet.Dimension.End;

            // Loop through all cells in the worksheet
            for (int row = startCell.Row; row <= endCell.Row; row++)
            {
                for (int col = startCell.Column; col <= endCell.Column; col++)
                {
                    // Get the cell value
                    object cellValue = worksheet.Cells[row, col].Value;

                    // Do something with the cell value (e.g., print it)
                    Console.Write(cellValue + "\t");
                }
                Console.WriteLine(); // Move to the next row
            }
        }
    }
}
*/