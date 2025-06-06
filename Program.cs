using System;

namespace RetailDatabaseManager
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "LR5-var1.xls";
            var manager = new DatabaseManager();

            try
            {
                manager.LoadDatabase(filePath);
                manager.ShowMainMenu();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Критическая ошибка: {ex.Message}");
                manager.Logger.Log($"Критическая ошибка: {ex.Message}");
            }
        }
    }
}
