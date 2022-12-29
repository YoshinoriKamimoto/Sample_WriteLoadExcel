using ClosedXML.Excel;

internal class Program
{
    private static void Main(string[] args)
    {
        // エクセル操作用のインスタンスを生成
        XLWorkbook book = new XLWorkbook(@"C:\Users\kamimoto\Desktop\work\tmp\Sample_WriteExcel\Book1.xlsx");

        // シート数を取得
        int sheetCnt = book.Worksheets.Count();
        Console.WriteLine($"シート数：{sheetCnt}");

        // シート数分処理をループ
        for (int i = 1; i <= sheetCnt; i++)
        {
            Console.WriteLine($"シート{i}");
            // 対象のシートを指定
            IXLWorksheet sheet = book.Worksheet(i);

            // 行数を取得
            int rowsCnt = sheet.RowsUsed().Count();
            Console.WriteLine($"行数：{rowsCnt}");

            // 行数分ループ
            for (int j = 1; j <= rowsCnt; j++)
            {
                // セルの値を取得
                string str = sheet.Cell(j, 1).Value.ToString();
                Console.WriteLine($"シート{j},セル({j},1)：{str}");
            }

            // セルに値を代入
            sheet.Cell(rowsCnt + 1, 1).SetValue("次へ");
            
        }

        // 変更を保存
        book.Save();
    }
}