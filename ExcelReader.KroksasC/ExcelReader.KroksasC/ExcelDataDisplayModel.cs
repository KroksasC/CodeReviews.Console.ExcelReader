using Spectre.Console;

namespace ExcelReader.KroksasC
{
    internal class ExcelDataDisplayModel
    {
        public static void ShowPersonsData()
        {
            Table table = new Table();

            var persons = DatabaseManagment.GetPersons();
            Thread.Sleep(1000);

            Console.WriteLine("Creating data table...\n");

            if (persons == null) 
            {
                Console.WriteLine("No person was found in file");
                Console.WriteLine("App shutting down...");
                Thread.Sleep(1000);
                Environment.Exit(0);
            }
            else
            {
                table.AddColumns("Id", "Name", "Age", "Score", "Category", "TimeStamp");
                foreach (var person in persons)
                {
                    table.AddRow(person.Id, person.Name, person.Age.ToString(), person.Score.ToString(), person.Category, person.TimeStamp.Value.ToString("yyyy-MM-dd"));
                }
                AnsiConsole.Write(table);
                Console.WriteLine("App shutting down...");
                Thread.Sleep(1000);
                Environment.Exit(0);
            }
        }
    }
}
