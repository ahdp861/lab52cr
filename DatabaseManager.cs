using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ConsoleApp
{
  public class DatabaseManager
  {
    private List<Dictionary<string, string>> database;
    public Logger Logger { get; }

    public DatabaseManager()
    {
      database = new List<Dictionary<string, string>>();
      Logger = new Logger("log.txt");
    }

    public void LoadDatabase(string filePath)
    {
      try
      {
        var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
          var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
          for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
          {
            string key = worksheet.Cells[1, col].Value?.ToString() ?? string.Empty;
            string value = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
            record[key] = value;
          }
          database.Add(record);
        }

        Logger.Log($"База данных успешно загружена из файла: {filePath}");
        Console.WriteLine("База данных успешно загружена.");
      }
      catch (Exception ex)
      {
        Logger.Log($"Ошибка при загрузке базы данных: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    public void ViewDatabase()
    {
      Logger.Log("Просмотр базы данных");
      Console.WriteLine("\nСодержимое базы данных:");

      if (database.Count == 0)
      {
        Console.WriteLine("База данных пуста.");
        return;
      }

      var keys = database[0].Keys.ToList();


      Console.WriteLine(string.Join("\t", keys));


      foreach (var record in database)
      {
        Console.WriteLine(string.Join("\t", keys.Select(k => record.TryGetValue(k, out var v) ? v : "N/A")));
      }
    }

    public void DeleteElement()
    {
      try
      {
        Console.Write("Введите идентификатор элемента для удаления: ");
        string id = Console.ReadLine();

        int initialCount = database.Count;
        database.RemoveAll(record => record.TryGetValue("Идентификатор", out var value) && value == id);

        if (database.Count < initialCount)
        {
          Logger.Log($"Удален элемент с идентификатором: {id}");
          Console.WriteLine("Элемент успешно удален.");
        }
        else
        {
          Logger.Log($"Попытка удаления элемента с идентификатором {id} - элемент не найден");
          Console.WriteLine("Элемент с указанным идентификатором не найден.");
        }
      }
      catch (Exception ex)
      {
        Logger.Log($"Ошибка при удалении элемента: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    public void EditElement()
    {
      try
      {
        Console.Write("Введите идентификатор элемента для редактирования: ");
        string id = Console.ReadLine();

        var record = database.FirstOrDefault(r => r.TryGetValue("Идентификатор", out var value) && value == id);

        if (record == null)
        {
          Logger.Log($"Попытка редактирования элемента с идентификатором {id} - элемент не найден");
          Console.WriteLine("Элемент с указанным идентификатором не найден.");
          return;
        }

        Logger.Log($"Начало редактирования элемента с идентификатором: {id}");

        var keys = record.Keys.ToList();

        foreach (var key in keys)
        {
          Console.Write($"{key} (текущее значение: {record[key]}): ");
          string newValue = Console.ReadLine();

          if (!string.IsNullOrEmpty(newValue) && newValue != record[key])
          {
            Logger.Log($"Изменено поле {key} с {record[key]} на {newValue}");
            record[key] = newValue;
          }
        }

        Logger.Log($"Завершено редактирование элемента с идентификатором: {id}");
        Console.WriteLine("Элемент успешно отредактирован.");
      }
      catch (Exception ex)
      {
        Logger.Log($"Ошибка при редактировании элемента: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    public void AddElement()
    {
      try
      {
        Logger.Log("Начало добавления нового элемента");

        var newRecord = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        Console.Write("Идентификатор: ");
        newRecord["Идентификатор"] = Console.ReadLine();

        Console.Write("Магазин: ");
        newRecord["Магазин"] = Console.ReadLine();

        Console.Write("Округ: ");
        newRecord["Округ"] = Console.ReadLine();

        Console.Write("Адрес: ");
        newRecord["Адрес"] = Console.ReadLine();

        Console.Write("Артикул: ");
        newRecord["Артикул"] = Console.ReadLine();

        Console.Write("Название: ");
        newRecord["Название"] = Console.ReadLine();

        Console.Write("Количество упаковок: ");
        newRecord["Количество упаковок"] = Console.ReadLine();

        Console.Write("Наличие карты покупателя: ");
        newRecord["Наличие карты покупателя"] = Console.ReadLine();

        Console.Write("ID категории: ");
        newRecord["ID категории"] = Console.ReadLine();

        Console.Write("Категория: ");
        newRecord["Категория"] = Console.ReadLine();

        Console.Write("Единица измерения: ");
        newRecord["Единица измерения"] = Console.ReadLine();

        Console.Write("Количество в упаковке: ");
        newRecord["Количество в упаковке"] = Console.ReadLine();

        Console.Write("Цена за упаковку: ");
        newRecord["Цена за упаковку"] = Console.ReadLine();

        database.Add(newRecord);
        Logger.Log($"Добавлен новый элемент с идентификатором: {newRecord["Идентификатор"]}");
        Console.WriteLine("Новый элемент успешно добавлен.");
      }
      catch (Exception ex)
      {
        Logger.Log($"Ошибка при добавлении элемента: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    public void ExecuteQueries()
    {
      try
      {
        Logger.Log("Выполнение запросов к базе данных");

        // Запрос 1: Общая стоимость детских товаров из категории «Радиоуправляемые игрушки 12+»
        var toysQuery = database
          .Where(record => record.TryGetValue("Категория", out var category) &&
                          category == "Радиоуправляемые игрушки 12+")
          .Sum(record => int.TryParse(record.TryGetValue("Количество упаковок", out var qty) ? qty : "0", out var qtyValue) &&
                      int.TryParse(record.TryGetValue("Цена за упаковку", out var price) ? price : "0", out var priceValue)
                      ? qtyValue * priceValue : 0);

        Console.WriteLine($"\nОбщая стоимость детских товаров из категории «Радиоуправляемые игрушки 12+»: {toysQuery}");
        Logger.Log($"Выполнен запрос: Общая стоимость детских товаров из категории «Радиоуправляемые игрушки 12+»: {toysQuery}");

        // Запрос 2: Средняя цена за упаковку по всем товарам
        var avgPriceQuery = database
          .Where(record => int.TryParse(record.TryGetValue("Цена за упаковку", out var price) ? price : "0", out _))
          .Average(record => int.Parse(record["Цена за упаковку"]));

        Console.WriteLine($"Средняя цена за упаковку по всем товарам: {avgPriceQuery:F2}");
        Logger.Log($"Выполнен запрос: Средняя цена за упаковку по всем товарам: {avgPriceQuery:F2}");

        // Запрос 3: Количество товаров в каждом магазине
        var storesQuery = database
          .GroupBy(record => record.TryGetValue("Магазин", out var store) ? store : "Неизвестно")
          .Select(group => new { Store = group.Key, Count = group.Count() });

        Console.WriteLine("\nКоличество товаров в каждом магазине:");
        foreach (var item in storesQuery)
        {
          Console.WriteLine($"{item.Store}: {item.Count}");
        }
        Logger.Log("Выполнен запрос: Количество товаров в каждом магазине");
      }
      catch (Exception ex)
      {
        Logger.Log($"Ошибка при выполнении запросов: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }
  }

  public class Logger
  {
    private readonly string _logFilePath;

    public Logger(string logFilePath)
    {
      _logFilePath = logFilePath;
      Log("Начало работы программы");
    }

    public void Log(string message)
    {
      string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
      try
      {
        File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
      }
      catch
      {
        Console.WriteLine($"Лог: {logMessage}");
      }
    }
  }
}
