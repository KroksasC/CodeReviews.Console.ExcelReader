namespace ExcelReader.KroksasC
{
    internal class DatabaseManagment
    {
        public static void DeleteDatabase()
        {
            using (var context = new ExcelContext())
            {
                Console.WriteLine("Deleting earlier database...\n");
                context.Database.EnsureDeleted();
            }    
        }
        public static void InsertExcelData()
        {
            using (var context = new ExcelContext())
            {
                Console.WriteLine("Creating new database...\n");
                Thread.Sleep(1000);
                context.Database.EnsureCreated();

                var persons = ExcelReaderService.GetPersons();

                foreach (Person person in persons) 
                {
                    context.Add(person);
                }
                Console.WriteLine("Inserting data to database...\n");
                Thread.Sleep(1000);
                context.SaveChanges();
            }
        }
        public static List<Person> GetPersons() 
        {
            using (var context = new ExcelContext()) 
            {
                Console.WriteLine("Extracting data from database...\n");
                Thread.Sleep(1000);
                return context.Persons.ToList();
            }
        }
    }
}
