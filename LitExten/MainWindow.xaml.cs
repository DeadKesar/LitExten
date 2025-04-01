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
        private List<(string Discipline, List<string> Literature, List<string> MaterialSupport)> _processedData =
            new List<(string, List<string>, List<string>)>();

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
                ExportLiteratureButton.IsEnabled = true;
                ExportMaterialButton.IsEnabled = true;
            }
        }

        private void ExportLiteratureButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "LiteratureExport"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                ExportToExcel(saveFileDialog.FileName, "LiteratureExport", AppendLiteratureCheckBox.IsChecked == true);
                StatusText.Text = $"Литература сохранена: {saveFileDialog.FileName}";
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
                        List<string> materialSupportLines = ExtractMaterialSupport(document);

                        // Парсинг дисциплины с первой страницы (шаги 5-7)
                        string firstPageText = GetFirstPageText(document);
                        string discipline = ExtractDiscipline(firstPageText);

                        // Сохраняем данные для экспорта
                        _processedData.Add((discipline, literatureLines, materialSupportLines));

                        // Добавляем результат для отображения
                        results.Add($"Файл: {Path.GetFileName(filePath)}");
                        results.Add($"Дисциплина: {discipline}");
                        results.Add("Литература:");
                        results.AddRange(literatureLines);
                        results.Add("Обеспечение:");
                        results.AddRange(materialSupportLines);
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
                new { Start = "Литература", End = "Периодические издания" },
                new { Start = "Литература", End = "Интернет-рес" }
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

        private List<string> ExtractMaterialSupport(DocX document)
        {
            var paragraphs = document.Paragraphs;
            int startIndex = -1;
            int endIndex = -1;
            List<int> listIndices = new();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i].Text;
                if (text.Contains("5. Материально-техническое обеспечение дисциплины", StringComparison.OrdinalIgnoreCase))
                {
                    startIndex = i;
                }
                else if (text.Contains("ЛИСТ", StringComparison.OrdinalIgnoreCase))
                {
                    listIndices.Add(i);
                }
                else if (startIndex != -1 && text.Contains("\f") && i > startIndex)
                {
                    endIndex = i;
                    if (i + 1 < paragraphs.Count && paragraphs[i + 1].Text.Contains("ЛИСТ", StringComparison.OrdinalIgnoreCase))
                    {
                        break; // Разрыв страницы с последующим "ЛИСТ" — конец раздела
                    }
                }
            }
            if (startIndex == -1)
            {
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    string text = paragraphs[i].Text;
                    if (text.Contains("Материально-техническое обеспечение дисциплины", StringComparison.OrdinalIgnoreCase))
                    {
                        startIndex = i;
                    }
                    else if (text.Contains("ЛИСТ", StringComparison.OrdinalIgnoreCase))
                    {
                        listIndices.Add(i);
                    }
                    else if (startIndex != -1 && text.Contains("\f") && i > startIndex)
                    {
                        endIndex = i;
                        if (i + 1 < paragraphs.Count && paragraphs[i + 1].Text.Contains("ЛИСТ", StringComparison.OrdinalIgnoreCase))
                        {
                            break; // Разрыв страницы с последующим "ЛИСТ" — конец раздела
                        }
                    }
                }
            }

            if (startIndex == -1)
                return new List<string> { "Раздел обеспечения не найден" };

            if (endIndex == -1 && listIndices.Any())
                endIndex = listIndices.Last(); // Последний "ЛИСТ" как конец, если нет разрыва с "ЛИСТ"

            if (endIndex == -1 || endIndex <= startIndex)
                endIndex = paragraphs.Count; // Если конец не найден, берем до конца документа

            List<string> materialSupportLines = new();
            for (int i = startIndex + 1; i < endIndex; i++)
            {
                string paragraphText = paragraphs[i].Text.Trim();
                if (!string.IsNullOrWhiteSpace(paragraphText))
                    materialSupportLines.Add(paragraphText);
            }

            return materialSupportLines.Count > 0 ? materialSupportLines : new List<string> { "Список обеспечения пуст" };
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
        private void ExportMaterialButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                FileName = "MaterialSupportExport"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                ExportToExcel(saveFileDialog.FileName, "MaterialSupportExport", AppendMaterialCheckBox.IsChecked == true);
                StatusText.Text = $"Обеспечение сохранено: {saveFileDialog.FileName}";
            }
        }
        private void ExportToExcel(string filePath, string exportType, bool appendMode)
        {
            ExcelPackage package;
            ExcelWorksheet worksheet;
            int startRow;
            int counter;

            bool isAppending = appendMode && File.Exists(filePath);
            string sheetName = exportType == "LiteratureExport" ? "Literature" : "MaterialSupport";

            if (isAppending)
            {
                package = new ExcelPackage(new FileInfo(filePath));
                worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName) ?? package.Workbook.Worksheets.Add(sheetName);

                startRow = worksheet.Dimension?.End.Row + 1 ?? 2;
                if (startRow == 2 && worksheet.Cells[1, 1].Value?.ToString() != "№")
                {
                    worksheet.Cells[1, 1].Value = "№";
                    worksheet.Cells[1, 2].Value = "Дисциплина";
                    worksheet.Cells[1, 3].Value = exportType == "LiteratureExport" ? "Литература" : "Обеспечение";
                }

                counter = 1;
                for (int row = 2; row < startRow; row++)
                {
                    var cellValue = worksheet.Cells[row, 1].Value;
                    if (cellValue != null && int.TryParse(cellValue.ToString(), out int num))
                        counter = Math.Max(counter, num + 1);
                }
            }
            else
            {
                package = new ExcelPackage();
                worksheet = package.Workbook.Worksheets.Add(sheetName);

                worksheet.Cells[1, 1].Value = "№";
                worksheet.Cells[1, 2].Value = "Дисциплина";
                worksheet.Cells[1, 3].Value = exportType == "LiteratureExport" ? "Литература" : "Обеспечение";
                startRow = 2;
                counter = 1;
            }

            using (var range = worksheet.Cells[1, 1, 1, 3])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }

            int currentRow = startRow;
            var data = exportType == "LiteratureExport" ?
                _processedData.Select(d => (d.Discipline, d.Literature)) :
                _processedData.Select(d => (d.Discipline, d.MaterialSupport));

            foreach (var (discipline, items) in data)
            {
                if (items.Count == 0) continue;

                int rowStart = currentRow;
                int rowEnd = currentRow + items.Count - 1;

                worksheet.Cells[rowStart, 1].Value = counter++;
                if (items.Count > 1)
                {
                    worksheet.Cells[rowStart, 1, rowEnd, 1].Merge = true;
                    worksheet.Cells[rowStart, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

                worksheet.Cells[rowStart, 2].Value = discipline;
                if (items.Count > 1)
                {
                    worksheet.Cells[rowStart, 2, rowEnd, 2].Merge = true;
                    worksheet.Cells[rowStart, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

                for (int i = 0; i < items.Count; i++)
                {
                    worksheet.Cells[currentRow + i, 3].Value = items[i];
                }

                currentRow = rowEnd + 1;
            }

            using (var range = worksheet.Cells[1, 1, currentRow - 1, 3])
            {
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            worksheet.Cells.AutoFitColumns();
            File.WriteAllBytes(filePath, package.GetAsByteArray());
            package.Dispose();
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