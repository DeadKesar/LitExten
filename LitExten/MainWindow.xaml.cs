using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Xceed.Words.NET;



namespace WordExcelParser
{
    public partial class MainWindow : Window
    {
        private List<(string Discipline, List<string> Literature)> _processedData = new List<(string, List<string>)>();

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("DeadKesar");
        }

        private void LoadFilesButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word Documents (*.docx)|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                StatusText.Text = $"Загружено файлов: {openFileDialog.FileNames.Length}";
                ProcessFiles(openFileDialog.FileNames);
                ExportButton.IsEnabled = true; // Активируем кнопку экспорта
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "LiteratureExport"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                ExportToExcel(saveFileDialog.FileName);
                StatusText.Text = $"Файл сохранен: {saveFileDialog.FileName}";
            }
        }


        private void ProcessFiles(string[] filePaths)
        {
            _processedData.Clear();
            List<string> results = new List<string>();
            ProgressBar.Maximum = filePaths.Length;
            ProgressBar.Value = 0;
            ProgressBar.Visibility = Visibility.Visible;

            foreach (string filePath in filePaths)
            {
                try
                {
                    using (DocX document = DocX.Load(filePath))
                    {
                        // Парсинг литературы (шаги 1-4)
                        List<string> literatureLines = ExtractLiterature(document);

                        // Парсинг дисциплины с первой страницы (шаги 5-7)
                        string firstPageText = GetFirstPageText(document);
                        string discipline = ExtractDiscipline(firstPageText);

                        // Сохраняем данные для экспорта
                        _processedData.Add((discipline, literatureLines));

                        // Добавляем результат для отображения
                        results.Add($"Файл: {Path.GetFileName(filePath)}");
                        results.Add($"Дисциплина: {discipline}");
                        results.Add("Литература:");
                        results.AddRange(literatureLines);
                        results.Add("---");
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Ошибка при обработке {Path.GetFileName(filePath)}: {ex.Message}");
                }
                ProgressBar.Value++;
            }

            ResultListBox.ItemsSource = results;
            ProgressBar.Visibility = Visibility.Collapsed;
        }

        // Извлечение текста между последними "4.1 Литература" и "4.2 Периодические издания" как списка параграфов
        private List<string> ExtractLiterature(DocX document)
        {
            var paragraphs = document.Paragraphs;
            var markers = new[]
            {
                new { Start = "4.1 Литература", End = "4.2 Периодические издания" },
                new { Start = "Литература", End = "Периодические издания" }
            };
            foreach (var marker in markers)
            {
                List<int> startIndices = new();
                List<int> endIndices = new();

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    string text = paragraphs[i].Text;
                    if (text.Contains(marker.Start, StringComparison.OrdinalIgnoreCase))
                        startIndices.Add(i);
                    else if (text.Contains(marker.End, StringComparison.OrdinalIgnoreCase))
                        endIndices.Add(i);
                }

                if (startIndices.Count > 0 && endIndices.Count > 0)
                {
                    int lastStartIndex = startIndices.Last();
                    int lastEndIndex = endIndices.FirstOrDefault(end => end > lastStartIndex, -1);

                    if (lastEndIndex != -1 && lastStartIndex < lastEndIndex)
                    {
                        List<string> literatureLines = new();
                        for (int i = lastStartIndex + 1; i < lastEndIndex; i++)
                        {
                            string paragraphText = paragraphs[i].Text.Trim();
                            if (!string.IsNullOrWhiteSpace(paragraphText))
                                literatureLines.Add(paragraphText);
                        }
                        return literatureLines.Count > 0 ? literatureLines : new List<string> { "Список литературы пуст" };
                    }
                }
            }

            return new List<string> { "Список литературы не был найден" };
        }


        // Получение текста первой страницы
        private string GetFirstPageText(DocX document)
        {
            // DocX не всегда корректно разделяет страницы, поэтому будем брать текст до первого разрыва страницы
            var paragraphs = document.Paragraphs;
            string firstPageText = string.Join("\n", paragraphs.TakeWhile(p => !p.Text.Contains("\f"))
                                                              .Select(p => p.Text));
            return firstPageText;
        }

        // Извлечение дисциплины между "РАБОЧАЯ ПРОГРАММА" и "Уровень высшего образования"
        private string ExtractDiscipline(string text)
        {
            string startMarker = "РАБОЧАЯ ПРОГРАММА";
            string endMarker = "Уровень высшего образования";

            int startIndex = text.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase);
            int endIndex = text.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase);

            if (startIndex == -1 || endIndex == -1 || startIndex >= endIndex)
                return "Не удалось найти дисциплину";

            startIndex += startMarker.Length;
            string discipline = text.Substring(startIndex, endIndex - startIndex).Trim();

            // Очистка от лишних символов и слов
            discipline = Regex.Replace(discipline, @"ДИСЦИПЛИНЫ|[\n\r]+|\s+", " ").Trim();
            return discipline;
        }
        private void ExportToExcel(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Literature");

                // Установка заголовков
                worksheet.Cells[1, 1].Value = "№";
                worksheet.Cells[1, 2].Value = "Дисциплина";
                worksheet.Cells[1, 3].Value = "Литература";

                int currentRow = 2; // Начальная строка для данных
                int counter = 1;    // Счетчик дисциплин

                foreach (var (discipline, literature) in _processedData)
                {
                    if (literature.Count == 0) continue; // Пропускаем пустые списки

                    // Определяем диапазон для объединения
                    int startRow = currentRow;
                    int endRow = currentRow + literature.Count - 1;

                    // Записываем номер дисциплины и объединяем ячейки
                    worksheet.Cells[startRow, 1].Value = counter++;
                    if (literature.Count > 1)
                    {
                        worksheet.Cells[startRow, 1, endRow, 1].Merge = true; // Объединяем ячейки для "№"
                    }

                    // Записываем дисциплину и объединяем ячейки
                    worksheet.Cells[startRow, 2].Value = discipline;
                    if (literature.Count > 1)
                    {
                        worksheet.Cells[startRow, 2, endRow, 2].Merge = true; // Объединяем ячейки для "Дисциплина"
                    }

                    // Записываем элементы литературы
                    for (int i = 0; i < literature.Count; i++)
                    {
                        worksheet.Cells[currentRow + i, 3].Value = literature[i];
                    }

                    // Переходим к следующей строке после списка литературы
                    currentRow = endRow + 1;
                }

                // Автоматическая настройка ширины столбцов
                worksheet.Cells.AutoFitColumns();

                // Сохраняем файл
                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
        }
    }
    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source?.IndexOf(toCheck, comp) >= 0;
        }
    }
}

//string pattern = @"(?<![0-9-])\d{1,2}\.\s";