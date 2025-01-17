namespace ExcelReader.KroksasC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DatabaseManagment.DeleteDatabase();
            DatabaseManagment.InsertExcelData();
            ExcelDataDisplayModel.ShowPersonsData();
        }
    }
}
