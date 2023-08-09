using OfficeOpenXml;

namespace ExelManager;

public static class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        ExcelPackage package = new(new FileInfo(@"C:\Users\SUPER-LAPTOP\Downloads\9.xlsx"));

        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        Dictionary<int, int> counts = new();

        int coolRows = 0;

        for (int r = 1; r < worksheet.Dimension.Rows + 1; r++)
        {
            for (int c = 1; c < worksheet.Dimension.Columns + 1; c++)
            {
                int key = int.Parse(worksheet.Cells[r, c].Text);

                if (counts.ContainsKey(key))
                    counts[key]++;
                else
                    counts.Add(key, 1);
            }
            if (MainCondition(counts))
                coolRows++;

            counts.Clear();
        }

        Console.WriteLine(coolRows);
    }

    private static bool SecondaryCondition(Dictionary<int, int> counts)
    {
        int sum = 0;
        int threesome = 0;

        foreach (var kv in counts)
        {
            if (kv.Value == 1)
                sum += kv.Key;
            else
                threesome = kv.Key;
        }

        return sum / 4 <= threesome;
    }

    private static bool MainCondition(Dictionary<int, int> counts)
    {
        int counts3s = 0;
        foreach (var val in counts.Values)
        {

            if (val is 2 or > 3)
                return false;
            else if (val == 3)
            {
                if (counts3s != 0)
                    return false;
                counts3s++;
            }
        }
        if (counts3s == 0)
            return false;

        return SecondaryCondition(counts);
    }
}