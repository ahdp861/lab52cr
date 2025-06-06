using System;
using System.Collections.Generic;

namespace ConsoleApp
{
  class Program
  {
    static void Main(string[] args)
    {
      string filePath = "LR5-var1.xlsx";
      var databaseManager = new DatabaseManager();
      databaseManager.LoadDatabase(filePath);

      while (true)
      {
        Console.WriteLine("1. Просмотреть базу данных");
        Console.WriteLine("2. Удалить элемент (по ключу)");
        Console.WriteLine("3. Редактировать элемент (по ключу)");
        Console.WriteLine("4. Добавить элемент");
        Console.WriteLine("5. Выполнить запросы");
        Console.WriteLine("6. Выйти");
        Console.Write("Выберите действие: ");

        string choice = Console.ReadLine();
        databaseManager.Logger.Log($"Пользователь выбрал действие: {choice}");

        switch (choice)
        {
          case "1":
            databaseManager.ViewDatabase();
            break;
          case "2":
            Console.Write("Введите ключ для удаления: ");
            string deleteKey = Console.ReadLine();
            databaseManager.DeleteElement(deleteKey);
            break;
          case "3":
            Console.Write("Введите ключ для редактирования: ");
            string editKey = Console.ReadLine();
            databaseManager.EditElement(editKey);
            break;
          case "4":
            databaseManager.AddElement();
            break;
          case "5":
            databaseManager.ExecuteQueries();
            break;
          case "6":
            databaseManager.Logger.Log("Программа завершена.");
            return;
          default:
            Console.WriteLine("Неверный выбор. Попробуйте снова.");
            databaseManager.Logger.Log("Ошибка: Неверный выбор действия.");
            break;
        }
      }
    }
  }
}
