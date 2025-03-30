using Microsoft.Win32;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using Xceed.Words.NET;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using Xceed.Document.NET;
using System.Xml.Linq;




namespace WordExcelParser
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
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
            }
        }


        private void ProcessFiles(string[] filePaths)
        {
            List<string> results = new List<string>();

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

                        // Добавляем результат для отображения
                        results.Add($"Файл: {filePath}");
                        results.Add($"Дисциплина: {discipline}");
                        results.Add("Литература:");
                        results.AddRange(literatureLines);
                        results.Add("---");
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Ошибка при обработке {filePath}: {ex.Message}");
                }
            }

            ResultListBox.ItemsSource = results;
        }


        // Извлечение текста между последними "4.1 Литература" и "4.2 Периодические издания" как списка параграфов
        private List<string> ExtractLiterature(DocX document)
        {
            string startMarker = "4.1 Литература";
            string endMarker = "4.2 Периодические издания";
            var paragraphs = document.Paragraphs;

            // Основная ветка: поиск маркеров как текста
            List<int> startIndices = new List<int>();
            List<int> endIndices = new List<int>();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i].Text;
                if (text.Contains(startMarker, StringComparison.OrdinalIgnoreCase))
                {
                    startIndices.Add(i);
                }
                else if (text.Contains(endMarker, StringComparison.OrdinalIgnoreCase))
                {
                    endIndices.Add(i);
                }
            }

            // Если маркеры найдены как текст
            if (startIndices.Count > 0 && endIndices.Count > 0)
            {
                int lastStartIndex = startIndices.Last();
                int lastEndIndex = endIndices.FirstOrDefault(end => end > lastStartIndex, -1);

                if (lastEndIndex != -1 && lastStartIndex < lastEndIndex)
                {
                    List<string> literatureLines = new List<string>();
                    for (int i = lastStartIndex + 1; i < lastEndIndex; i++)
                    {
                        string paragraphText = paragraphs[i].Text.Trim();
                        if (!string.IsNullOrWhiteSpace(paragraphText))
                        {
                            literatureLines.Add(paragraphText);
                        }
                    }
                    return literatureLines.Count > 0 ? literatureLines : new List<string> { "Список литературы пуст" };
                }
            }

            startMarker = "Литература";
            endMarker = "Периодические издания";
          

            // Основная ветка: поиск маркеров как текста
            startIndices = new List<int>();
            endIndices = new List<int>();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string text = paragraphs[i].Text;
                if (text.Contains(startMarker, StringComparison.OrdinalIgnoreCase))
                {
                    startIndices.Add(i);
                }
                else if (text.Contains(endMarker, StringComparison.OrdinalIgnoreCase))
                {
                    endIndices.Add(i);
                }
            }

            // Если маркеры найдены как текст
            if (startIndices.Count > 0 && endIndices.Count > 0)
            {
                int lastStartIndex = startIndices.Last();
                int lastEndIndex = endIndices.FirstOrDefault(end => end > lastStartIndex, -1);

                if (lastEndIndex != -1 && lastStartIndex < lastEndIndex)
                {
                    List<string> literatureLines = new List<string>();
                    for (int i = lastStartIndex + 1; i < lastEndIndex; i++)
                    {
                        string paragraphText = paragraphs[i].Text.Trim();
                        if (!string.IsNullOrWhiteSpace(paragraphText))
                        {
                            literatureLines.Add(paragraphText);
                        }
                    }
                    return literatureLines.Count > 0 ? literatureLines : new List<string> { "Список литературы пуст" };
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
            {
                return "Не удалось найти дисциплину";
            }

            startIndex += startMarker.Length;
            string discipline = text.Substring(startIndex, endIndex - startIndex).Trim();

            // Очистка от лишних переходов строк и пробелов
            discipline = Regex.Replace(discipline, @"ДИСЦИПЛИНЫ", " ");
            discipline = Regex.Replace(discipline, @"\s+", " ");
            
            return discipline;
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