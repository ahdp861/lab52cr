using System;

namespace RetailDatabaseManager
{
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
}
