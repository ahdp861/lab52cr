using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace RetailDatabaseManager {
  /// <summary>
  /// Основной класс программы, содержащий точку входа.
  /// </summary>
  class Program {
    /// <summary>
    /// Точка входа в программу. Создает экземпляр DatabaseManager и запускает основной цикл программы.
    /// </summary>
    /// <param name="args">Аргументы командной строки (не используются).</param>
    static void Main(string[] args) {
      string filePath = "retail_database.xlsx";
      var manager = new DatabaseManager();

      try {
        manager.LoadDatabase(filePath);
        manager.ShowMainMenu();
      } catch (Exception ex) {
        Console.WriteLine($"Критическая ошибка: {ex.Message}");
        manager.Logger.Log($"Критическая ошибка: {ex.Message}");
      }
    }
  }

  /// <summary>
  /// Основной класс для управления базой данных, загруженной из Excel-файла.
  /// Обеспечивает функциональность для просмотра, редактирования и анализа данных.
  /// </summary>
  public class DatabaseManager {
    private Dictionary<string, List<Dictionary<string, string>>> database;

    /// <summary>
    /// Получает экземпляр логгера для записи операций.
    /// </summary>
    public Logger Logger { get; }

    /// <summary>
    /// Инициализирует новый экземпляр класса DatabaseManager.
    /// </summary>
    public DatabaseManager() {
      database = new Dictionary<string, List<Dictionary<string, string>>>();
      Logger = new Logger("operations.log");
    }

    /// <summary>
    /// Загружает данные из Excel-файла в память.
    /// </summary>
    /// <param name="filePath">Путь к Excel-файлу.</param>
    /// <exception cref="System.Exception">Возникает при ошибке загрузки данных.</exception>
    public void LoadDatabase(string filePath) {
      Logger.Log("Начало загрузки базы данных");

      try {
        var package = new ExcelPackage(new FileInfo(filePath));

        foreach (var worksheet in package.Workbook.Worksheets) {
          var sheetData = new List<Dictionary<string, string>>();
          var sheetName = worksheet.Name;

          if (worksheet.Dimension == null) continue;

          var headers = new List<string>();
          for (int col = 1; col <= worksheet.Dimension.End.Column; col++) {
            headers.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
          }

          for (int row = 2; row <= worksheet.Dimension.End.Row; row++) {
            var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++) {
              string key = headers[col - 1];
              string value = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
              record[key] = value;
            }
            sheetData.Add(record);
          }

          database[sheetName] = sheetData;
          Logger.Log($"Загружен лист: {sheetName}, записей: {sheetData.Count}");
        }

        Console.WriteLine("База данных успешно загружена.");
        Logger.Log("База данных успешно загружена");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при загрузке базы данных: {ex.Message}");
        throw new Exception($"Ошибка при загрузке базы данных: {ex.Message}");
      }
    }

    /// <summary>
    /// Отображает главное меню программы и обрабатывает выбор пользователя.
    /// </summary>
    public void ShowMainMenu() {
      while (true) {
        Console.WriteLine("\nГлавное меню:");
        Console.WriteLine("1. Просмотр данных");
        Console.WriteLine("2. Редактирование данных");
        Console.WriteLine("3. Выполнение запросов");
        Console.WriteLine("4. Выход");
        Console.Write("Выберите действие: ");

        string choice = Console.ReadLine();
        Logger.Log($"Выбран пункт меню: {choice}");

        switch (choice) {
          case "1":
            ShowViewMenu();
            break;
          case "2":
            ShowEditMenu();
            break;
          case "3":
            ExecuteQueries();
            break;
          case "4":
            Logger.Log("Завершение работы программы");
            return;
          default:
            Console.WriteLine("Ошибка: Неверный выбор. Попробуйте снова.");
            break;
        }
      }
    }

    /// <summary>
    /// Отображает меню просмотра данных и обрабатывает выбор пользователя.
    /// </summary>
    private void ShowViewMenu() {
      while (true) {
        Console.WriteLine("\nМеню просмотра:");
        Console.WriteLine("1. Просмотр движения товаров");
        Console.WriteLine("2. Просмотр магазинов");
        Console.WriteLine("3. Просмотр товаров");
        Console.WriteLine("4. Просмотр категорий");
        Console.WriteLine("5. Назад");
        Console.Write("Выберите действие: ");

        string choice = Console.ReadLine();

        switch (choice) {
          case "1":
            ViewData("Движение товаров");
            break;
          case "2":
            ViewData("Магазин");
            break;
          case "3":
            ViewData("Товар");
            break;
          case "4":
            ViewData("Категория");
            break;
          case "5":
            return;
          default:
            Console.WriteLine("Ошибка: Неверный выбор. Попробуйте снова.");
            break;
        }
      }
    }

    /// <summary>
    /// Отображает меню редактирования данных и обрабатывает выбор пользователя.
    /// </summary>
    private void ShowEditMenu() {
      while (true) {
        Console.WriteLine("\nМеню редактирования:");
        Console.WriteLine("1. Добавить запись");
        Console.WriteLine("2. Редактировать запись");
        Console.WriteLine("3. Удалить запись");
        Console.WriteLine("4. Назад");
        Console.Write("Выберите действие: ");

        string choice = Console.ReadLine();

        switch (choice) {
          case "1":
            AddRecordMenu();
            break;
          case "2":
            EditRecordMenu();
            break;
          case "3":
            DeleteRecordMenu();
            break;
          case "4":
            return;
          default:
            Console.WriteLine("Ошибка: Неверный выбор. Попробуйте снова.");
            break;
        }
      }
    }

    /// <summary>
    /// Отображает данные из указанного листа.
    /// </summary>
    /// <param name="sheetName">Название листа для отображения.</param>
    private void ViewData(string sheetName) {
      if (!database.ContainsKey(sheetName)) {
        Console.WriteLine($"Ошибка: Лист '{sheetName}' не найден в базе данных.");
        Logger.Log($"Попытка просмотра несуществующего листа: {sheetName}");
        return;
      }

      var data = database[sheetName];
      Console.WriteLine($"\nДанные из листа '{sheetName}' (записей: {data.Count}):");

      if (data.Count == 0) {
        Console.WriteLine("Нет данных для отображения.");
        return;
      }

      var headers = data[0].Keys.ToList();
      Console.WriteLine(string.Join("\t", headers));

      foreach (var record in data) {
        Console.WriteLine(string.Join("\t", headers.Select(h => record.TryGetValue(h, out var v) ? v : "N/A")));
      }

      Logger.Log($"Просмотрены данные из листа: {sheetName}");
    }

    /// <summary>
    /// Обрабатывает процесс добавления новой записи в выбранный лист.
    /// </summary>
    private void AddRecordMenu() {
      Console.WriteLine("\nВыберите лист для добавления записи:");
      Console.WriteLine("1. Движение товаров");
      Console.WriteLine("2. Магазин");
      Console.WriteLine("3. Товар");
      Console.WriteLine("4. Категория");
      Console.Write("Выберите лист: ");

      string choice = Console.ReadLine();
      string sheetName = choice switch {
        "1" => "Движение товаров",
        "2" => "Магазин",
        "3" => "Товар",
        "4" => "Категория",
        _ => ""
      };

      if (string.IsNullOrEmpty(sheetName)) {
        Console.WriteLine("Ошибка: Неверный выбор листа.");
        return;
      }

      if (!database.ContainsKey(sheetName)) {
        Console.WriteLine($"Ошибка: Лист '{sheetName}' не найден в базе данных.");
        return;
      }

      try {
        var newRecord = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sampleRecord = database[sheetName].FirstOrDefault();

        if (sampleRecord == null) {
          Console.WriteLine("Ошибка: Нет образца записи для этого листа.");
          return;
        }

        Console.WriteLine($"\nДобавление новой записи в лист '{sheetName}':");

        foreach (var field in sampleRecord.Keys) {
          Console.Write($"{field}: ");
          newRecord[field] = Console.ReadLine();
        }

        database[sheetName].Add(newRecord);
        Logger.Log($"Добавлена новая запись в лист {sheetName}");
        Console.WriteLine("Запись успешно добавлена.");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при добавлении записи: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Обрабатывает процесс редактирования существующей записи в выбранном листе.
    /// </summary>
    private void EditRecordMenu() {
      Console.WriteLine("\nВыберите лист для редактирования записи:");
      Console.WriteLine("1. Движение товаров");
      Console.WriteLine("2. Магазин");
      Console.WriteLine("3. Товар");
      Console.WriteLine("4. Категория");
      Console.Write("Выберите лист: ");

      string choice = Console.ReadLine();
      string sheetName = choice switch {
        "1" => "Движение товаров",
        "2" => "Магазин",
        "3" => "Товар",
        "4" => "Категория",
        _ => ""
      };

      if (string.IsNullOrEmpty(sheetName)) {
        Console.WriteLine("Ошибка: Неверный выбор листа.");
        return;
      }

      if (!database.ContainsKey(sheetName)) {
        Console.WriteLine($"Ошибка: Лист '{sheetName}' не найден в базе данных.");
        return;
      }

      try {
        Console.Write("Введите ID записи для редактирования: ");
        string id = Console.ReadLine();

        var records = database[sheetName];
        var recordToEdit = records.FirstOrDefault(r => r.TryGetValue("ID", out var val) && val == id);

        if (recordToEdit == null) {
          Console.WriteLine($"Запись с ID {id} не найдена.");
          Logger.Log($"Попытка редактирования несуществующей записи с ID {id} в листе {sheetName}");
          return;
        }

        Console.WriteLine($"\nРедактирование записи с ID {id} в листе '{sheetName}':");

        foreach (var field in recordToEdit.Keys.ToList()) {
          Console.Write($"{field} (текущее значение: {recordToEdit[field]}): ");
          string newValue = Console.ReadLine();

          if (!string.IsNullOrEmpty(newValue) && newValue != recordToEdit[field]) {
            recordToEdit[field] = newValue;
            Logger.Log($"Изменено поле {field} с {recordToEdit[field]} на {newValue} в записи {id} листа {sheetName}");
          }
        }

        Console.WriteLine("Запись успешно отредактирована.");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при редактировании записи: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Обрабатывает процесс удаления записи из выбранного листа.
    /// </summary>
    private void DeleteRecordMenu() {
      Console.WriteLine("\nВыберите лист для удаления записи:");
      Console.WriteLine("1. Движение товаров");
      Console.WriteLine("2. Магазин");
      Console.WriteLine("3. Товар");
      Console.WriteLine("4. Категория");
      Console.Write("Выберите лист: ");

      string choice = Console.ReadLine();
      string sheetName = choice switch {
        "1" => "Движение товаров",
        "2" => "Магазин",
        "3" => "Товар",
        "4" => "Категория",
        _ => ""
      };

      if (string.IsNullOrEmpty(sheetName)) {
        Console.WriteLine("Ошибка: Неверный выбор листа.");
        return;
      }

      if (!database.ContainsKey(sheetName)) {
        Console.WriteLine($"Ошибка: Лист '{sheetName}' не найден в базе данных.");
        return;
      }

      try {
        Console.Write("Введите ID записи для удаления: ");
        string id = Console.ReadLine();

        var records = database[sheetName];
        int initialCount = records.Count;
        records.RemoveAll(r => r.TryGetValue("ID", out var val) && val == id);

        if (records.Count < initialCount) {
          Logger.Log($"Удалена запись с ID {id} из листа {sheetName}");
          Console.WriteLine("Запись успешно удалена.");
        } else {
          Logger.Log($"Попытка удаления несуществующей записи с ID {id} из листа {sheetName}");
          Console.WriteLine($"Запись с ID {id} не найдена.");
        }
      } catch (Exception ex) {
        Logger.Log($"Ошибка при удалении записи: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Отображает меню запросов и обрабатывает выбор пользователя.
    /// </summary>
    public void ExecuteQueries() {
      try {
        Logger.Log("Выполнение запросов к базе данных");

        while (true) {
          Console.WriteLine("\nМеню запросов:");
          Console.WriteLine("1. Топ-5 самых продаваемых товаров");
          Console.WriteLine("2. Товары с наибольшей скидкой");
          Console.WriteLine("3. Продажи по магазинам");
          Console.WriteLine("4. Товары без продаж");
          Console.WriteLine("5. Назад");
          Console.Write("Выберите запрос: ");

          string choice = Console.ReadLine();

          switch (choice) {
            case "1":
              QueryTopSellingProducts();
              break;
            case "2":
              QueryMostDiscountedProducts();
              break;
            case "3":
              QuerySalesByStores();
              break;
            case "4":
              QueryProductsWithoutSales();
              break;
            case "5":
              return;
            default:
              Console.WriteLine("Ошибка: Неверный выбор. Попробуйте снова.");
              break;
          }
        }
      } catch (Exception ex) {
        Logger.Log($"Ошибка при выполнении запросов: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Выполняет запрос для получения топ-5 самых продаваемых товаров.
    /// </summary>
    private void QueryTopSellingProducts() {
      try {
        if (!database.ContainsKey("Движение товаров") || !database.ContainsKey("Товар")) {
          Console.WriteLine("Ошибка: Необходимые листы не найдены в базе данных.");
          return;
        }

        var sales = database["Движение товаров"];
        var products = database["Товар"];

        var topProducts = sales
          .GroupBy(s => s["ID товара"])
          .Select(g => new {
            ProductId = g.Key,
            TotalQuantity = g.Sum(s => int.TryParse(s["Количество упаковок"], out var qty) ? qty : 0)
          })
          .OrderByDescending(p => p.TotalQuantity)
          .Take(5)
          .ToList();

        Console.WriteLine("\nТоп-5 самых продаваемых товаров:");
        Console.WriteLine("ID товара\tНазвание\tОбщее количество");

        foreach (var product in topProducts) {
          var productInfo = products.FirstOrDefault(p => p["ID"] == product.ProductId);
          string name = productInfo?["Название"] ?? "Неизвестно";
          Console.WriteLine($"{product.ProductId}\t{name}\t{product.TotalQuantity}");
        }

        Logger.Log("Выполнен запрос: Топ-5 самых продаваемых товаров");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при выполнении запроса топ-5 товаров: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Выполняет запрос для получения товаров с наибольшей скидкой.
    /// </summary>
    private void QueryMostDiscountedProducts() {
      try {
        if (!database.ContainsKey("Товар")) {
          Console.WriteLine("Ошибка: Лист 'Товар' не найден в базе данных.");
          return;
        }

        var products = database["Товар"]
          .Where(p => decimal.TryParse(p.TryGetValue("Цена за упаковку", out var priceStr) ? priceStr : "0", out var price) &&
                      decimal.TryParse(p.TryGetValue("Скидка", out var discountStr) ? discountStr : "0", out var discount))
          .OrderByDescending(p => decimal.Parse(p["Скидка"]))
          .Take(5)
          .ToList();

        Console.WriteLine("\nТовары с наибольшей скидкой:");
        Console.WriteLine("ID\tНазвание\tСкидка\tЦена");

        foreach (var product in products) {
          Console.WriteLine($"{product["ID"]}\t{product["Название"]}\t{product["Скидка"]}%\t{product["Цена за упаковку"]}");
        }

        Logger.Log("Выполнен запрос: Товары с наибольшей скидкой");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при выполнении запроса товаров со скидкой: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Выполняет запрос для получения информации о продажах по магазинам.
    /// </summary>
    private void QuerySalesByStores() {
      try {
        if (!database.ContainsKey("Движение товаров") || !database.ContainsKey("Магазин")) {
          Console.WriteLine("Ошибка: Необходимые листы не найдены в базе данных.");
          return;
        }

        var sales = database["Движение товаров"];
        var stores = database["Магазин"];

        var salesByStore = sales
          .GroupBy(s => s["ID магазина"])
          .Select(g => new {
            StoreId = g.Key,
            TotalSales = g.Sum(s => int.TryParse(s["Количество упаковок"], out var qty) ? qty : 0)
          })
          .OrderByDescending(s => s.TotalSales)
          .ToList();

        Console.WriteLine("\nПродажи по магазинам:");
        Console.WriteLine("ID магазина\tНазвание\tОбщие продажи");

        foreach (var store in salesByStore) {
          var storeInfo = stores.FirstOrDefault(s => s["ID"] == store.StoreId);
          string name = storeInfo?["Название"] ?? "Неизвестно";
          Console.WriteLine($"{store.StoreId}\t{name}\t{store.TotalSales}");
        }

        Logger.Log("Выполнен запрос: Продажи по магазинам");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при выполнении запроса продаж по магазинам: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }

    /// <summary>
    /// Выполняет запрос для получения товаров, которые не продавались.
    /// </summary>
    private void QueryProductsWithoutSales() {
      try {
        if (!database.ContainsKey("Товар") || !database.ContainsKey("Движение товаров")) {
          Console.WriteLine("Ошибка: Необходимые листы не найдены в базе данных.");
          return;
        }

        var products = database["Товар"];
        var sales = database["Движение товаров"];

        var soldProductIds = sales.Select(s => s["ID товара"]).Distinct().ToList();
        var allProductIds = products.Select(p => p["ID"]).ToList();

        var unsoldProducts = products
          .Where(p => !soldProductIds.Contains(p["ID"]))
          .ToList();

        Console.WriteLine("\nТовары без продаж:");
        Console.WriteLine("ID\tНазвание\tКатегория");

        foreach (var product in unsoldProducts) {
          Console.WriteLine($"{product["ID"]}\t{product["Название"]}\t{product["Категория"]}");
        }

        Logger.Log($"Выполнен запрос: Товары без продаж. Найдено {unsoldProducts.Count} товаров");
      } catch (Exception ex) {
        Logger.Log($"Ошибка при выполнении запроса товаров без продаж: {ex.Message}");
        Console.WriteLine($"Ошибка: {ex.Message}");
      }
    }
  }

  /// <summary>
  /// Класс для логирования операций программы.
  /// </summary>
  public class Logger {
    private readonly string _logFilePath;

    /// <summary>
    /// Инициализирует новый экземпляр класса Logger.
    /// </summary>
    /// <param name="logFilePath">Путь к файлу лога.</param>
    public Logger(string logFilePath) {
      _logFilePath = logFilePath;
      Log("Начало работы программы");
    }

    /// <summary>
    /// Записывает сообщение в лог-файл.
    /// </summary>
    /// <param name="message">Сообщение для записи в лог.</param>
    public void Log(string message) {
      string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
      try {
        File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
      } catch {
        Console.WriteLine($"Лог: {logMessage}");
      }
    }
  }
}
