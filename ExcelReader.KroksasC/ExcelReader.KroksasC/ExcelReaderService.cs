using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Reflection;

namespace ExcelReader.KroksasC
{
    public static class ExcelReaderService
    {
        public static IEnumerable<T> ConvertToObjects<T>(this ExcelTable table) where T : new()
        {
            ExcelCellAddress start = table.Address.Start;
            ExcelCellAddress end = table.Address.End;
            List<ExcelRange> cells = new();

            for (int r = start.Row; r <= end.Row; r++)
                for (int c = start.Column; c <= end.Column; c++)
                    cells.Add(table.WorkSheet.Cells[r, c]);

            List<IGrouping<int, ExcelRange>> allRows = cells
                .GroupBy(cell => cell.Start.Row)
                .OrderBy(cell => cell.Key)
                .ToList();

            IEnumerable<PropertyInfo> typeProperties = typeof(T).GetProperties();

            IGrouping<int, ExcelRangeBase> header = allRows.First();

            Dictionary<PropertyInfo, int> columns = new();

            foreach (ExcelRangeBase col in header)
            {
                string propName = col.GetValue<string>();
                PropertyInfo propInfo = typeProperties.FirstOrDefault(x => x.Name.Equals(propName));
                if (propInfo != null)
                {
                    columns.Add(propInfo, col.Start.Column);
                }
            }

            IEnumerable<IGrouping<int, ExcelRangeBase>> rows = allRows.Skip(1);

            List<T> objects = new();

            foreach (IGrouping<int, ExcelRangeBase> row in rows)
            {
                T obj = new();
                foreach (KeyValuePair<PropertyInfo, int> colInfo in columns)
                {
                    ExcelRangeBase col = row.First(x => x.Start.Column == colInfo.Value);

                    if (col.Value == null)
                        continue;

                    object value = Convert.ChangeType(col.Value, Nullable.GetUnderlyingType(colInfo.Key.PropertyType) ?? colInfo.Key.PropertyType);
                    colInfo.Key.SetValue(obj, value);
                }
                objects.Add(obj);
            }

            return objects;
        }
        internal static IEnumerable<Person> GetPersons()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(new FileInfo("C:\\Users\\Admins\\source\\repos\\ExcelReader.KroksasC\\ExcelReader.KroksasC\\Learning_Data.xlsx"));
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            ExcelTable table = worksheet.Tables.First();
            Console.WriteLine("Reading excel file...\n");
            Thread.Sleep(1000);
            var persons = table.ConvertToObjects<Person>();
            return persons;
        }
    }
}
