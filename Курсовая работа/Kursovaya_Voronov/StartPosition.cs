using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using System.Text.RegularExpressions;

namespace Kursovaya_rabota
{
    public partial class kursovaya : Form
    {
        Year y = new Year();
        Publish p = new Publish();
        public kursovaya()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }
        private Dictionary<string, int> coauthorsCount = new Dictionary<string, int>();

        private void button3_Click(object sender, EventArgs e)
        {
            // Открыть диалог для выбора файла Word
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.doc;*.docx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Вызываем метод конвертации
                ConvertWordToExcel(openFileDialog.FileName);
            }
        }
        private void ConvertWordToExcel(string wordFilePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false; // Открываем Word в фоновом режиме
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(wordFilePath, ReadOnly: true);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false; // Открываем Excel в фоновом режиме
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();

            int row = 1; // Начальная строка для записи в Excel

            // Перебираем все таблицы в документе
            for (int tableIndex = 1; tableIndex <= doc.Tables.Count; tableIndex++)
            {
                Microsoft.Office.Interop.Word.Table table = doc.Tables[tableIndex];

                // Проверяем, если в таблице более 4 строк (1 заголовок + 3 записи)
                if (table.Rows.Count > 4)
                {
                    // Обработка таблицы, если она удовлетворяет условию
                    row = WriteTableToExcel(table, workbook, row);
                }
            }

            string excelFilePath = System.IO.Path.ChangeExtension(wordFilePath, ".xlsx");

            // Сохранение Excel файла
            workbook.SaveAs(excelFilePath);

            workbook.Close(false); // Закрываем книгу без сохранения
            excelApp.Quit();
            doc.Close(false); // Закрываем документ без сохранения
            wordApp.Quit();

            // Освобождаем ресурсы
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

            MessageBox.Show($"Конвертация завершена! Файл сохранен как: {excelFilePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Dictionary<string, int> newCoauthorsCount = GetUniqueCoauthorsCount(excelFilePath);
            UpdateChart(newCoauthorsCount);

            ConvertExcelToCsv(excelFilePath); //конвертация excel в csv

            y.UpdatePublicationCounts(excelFilePath);
            p.UpdateVolumeCounts(excelFilePath);

        }
        
        private int WriteTableToExcel(Microsoft.Office.Interop.Word.Table table, Microsoft.Office.Interop.Excel.Workbook workbook, int startRow)
        {
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

            // Заполняем заголовки, если это первая таблица, которую мы обрабатываем
            if (startRow == 1)
            {
                worksheet.Cells[1, 1] = "№ п/п";
                worksheet.Cells[1, 2] = "Название";
                worksheet.Cells[1, 3] = "Печатный или на правах рукописи";
                worksheet.Cells[1, 4] = "Издательство, журнал (название, год, номер)";
                worksheet.Cells[1, 5] = "Количество печатных листов или страниц";
                worksheet.Cells[1, 6] = "Фамилия соавторов";
                startRow++;
            }

            // Заполняем данные таблицы
            for (int i = 2; i <= table.Rows.Count; i++)
            {
                // Получаем ячейки строки
                string cell1 = table.Cell(i, 1).Range.Text.Trim('\r', '\a').Replace(".", ""); // Заменяем точки на пустую строку
                var cell2 = table.Cell(i, 2).Range.Text.Trim('\r', '\a');
                string cell3 = table.Cell(i, 3).Range.Text.Trim('\r', '\a');
                string cell4 = table.Cell(i, 4).Range.Text.Trim('\r', '\a');
                string cell5 = table.Cell(i, 5).Range.Text.Trim('\r', '\a');
                string cell6 = table.Cell(i, 6).Range.Text.Trim('\r', '\a');

                if (cell2 == "2")
                {
                    continue; // Пропускаем запись для текущей строки
                }
                // Заполнение ячеек Excel
                worksheet.Cells[startRow, 1] = cell1;
                worksheet.Cells[startRow, 2] = cell2;
                worksheet.Cells[startRow, 3] = cell3;
                worksheet.Cells[startRow, 4] = cell4;
                worksheet.Cells[startRow, 5] = cell5;
                worksheet.Cells[startRow, 6] = cell6;

                startRow++;
            }

            return startRow; // Возвращаем следующую строку для записи
        }

        private string NormalizeCoauthorName(string name)
        {
            // Удаляем пробелы и приводим к нижнему регистру
            string normalized = name.ToLower().Trim();

            // Разделяем имя на части
            var nameParts = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            // Проверяем, чтобы было хотя бы две части
            if (nameParts.Length >= 2)
            {
                if (nameParts[0].Contains(".") && !nameParts[nameParts.Length - 1].Contains("."))
                {
                    // Форматируем как "Фамилия Инициалы"
                    string initials = string.Join(" ", nameParts.Take(nameParts.Length - 1));
                    string surname = nameParts[nameParts.Length - 1];
                    return $"{surname} {initials}";
                }

                return normalized; 
            }

            return normalized; // На случай, если имя состоит только из одной части
        }

        private Dictionary<string, List<string>> coauthorsDictionary = new Dictionary<string, List<string>>(); // Словарь для хранения соавторов

        private Dictionary<string, int> GetUniqueCoauthorsCount(string excelFilePath)
        {
            HashSet<string> uniqueCoauthors = new HashSet<string>();
            List<string> coauthorsList = new List<string>();

            // Открываем Excel и получаем первый лист
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            // Получаем количество строк в листе
            int lastRow = worksheet.Cells[worksheet.Rows.Count, 6].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;

            for (int i = 2; i <= lastRow; i++)
            {
                string coauthorsCell = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[i, 6]).Text;

                if (!string.IsNullOrWhiteSpace(coauthorsCell) && coauthorsCell.ToLower() != "нет")
                {
                    // Разделяем имена по запятой и переносу строки
                    string[] coauthors = coauthorsCell.Split(new[] { ',', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var coauthor in coauthors)
                    {
                        // Нормализуем имя соавтора
                        string coauthorName = NormalizeCoauthorName(coauthor.Trim());

                        // Добавляем нормализованное имя в HashSet
                        uniqueCoauthors.Add(coauthorName);
                        coauthorsList.Add(coauthorName); // Добавляем в общий список
                    }
                }
            }

            // Закрываем Excel
            workbook.Close(false);
            excelApp.Quit();

            // Освобождаем ресурсы
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(excelFilePath);

            // Добавляем соавторов в словарь
            coauthorsDictionary[fileNameWithoutExtension] = coauthorsList;

            return new Dictionary<string, int> { { fileNameWithoutExtension, uniqueCoauthors.Count } };
        }

        // Обновляем Chart
        private void UpdateChart(Dictionary<string, int> newCoauthorsCount)
        {
            foreach (KeyValuePair<string, int> entry in newCoauthorsCount)
            {
                if (!coauthorsCount.ContainsKey(entry.Key))
                {
                    coauthorsCount[entry.Key] = 0;
                }
                coauthorsCount[entry.Key] += entry.Value;
            }

            chart1.ChartAreas[0].AxisY.Interval = 1;
            chart1.ChartAreas[0].AxisY.Title = "Количество уникальных соавторов";
            chart1.ChartAreas[0].AxisX.Title = "";

            foreach (KeyValuePair<string, int> entry in newCoauthorsCount)
            {
                System.Windows.Forms.DataVisualization.Charting.Series series = chart1.Series.FirstOrDefault(s => s.Name == entry.Key);
                if (series == null)
                {
                    series = new System.Windows.Forms.DataVisualization.Charting.Series(entry.Key)
                    {
                        ChartType = SeriesChartType.Column
                    };
                    chart1.Series.Add(series);
                }
                series.Points.AddXY(entry.Key, entry.Value);
            }

            chart1.ChartAreas[0].AxisX.Title = "Уникальные соавторы";

            foreach (DataPoint point in chart1.Series.SelectMany(s => s.Points))
            {
                point.AxisLabel = "";
            }

            chart1.Invalidate();
        }
        // Обработчик клика по диаграмме
        private void chart1_Click(object sender, EventArgs e)
        {
            System.Drawing.Point mousePosition = PointToClient(MousePosition);
            HitTestResult hitTestResult = chart1.HitTest(mousePosition.X, mousePosition.Y);

            if (hitTestResult.ChartElementType == ChartElementType.DataPoint ||
                hitTestResult.ChartElementType == ChartElementType.PlottingArea)
            {
                System.Windows.Forms.DataVisualization.Charting.Series series = hitTestResult.Series;
                int pointIndex = hitTestResult.PointIndex;

                if (series != null && pointIndex >= 0)
                {
                    string seriesName = series.Name; // Получаем имя серии

                    if (coauthorsDictionary.TryGetValue(seriesName, out List<string> coauthors))
                    {
                        // Используем HashSet для уникальных соавторов
                        HashSet<string> uniqueCoauthors = new HashSet<string>(coauthors);
                        listBox1.Items.Clear(); // Очищаем список перед добавлением новых элементов (если это необходимо)
                        foreach (var coauthor in uniqueCoauthors)
                        {
                            if (coauthor != "")
                            {
                                listBox1.Items.Add(FormatCoauthor(coauthor)); // Добавляем каждого соавтора в ListBox
                            }
                        }
                        y.UpdateChartForYears(seriesName);
                        p.UpdateVolumeChart(seriesName);
                    }
                }
                else if (hitTestResult.ChartElementType == ChartElementType.PlottingArea)
                {
                    foreach (System.Windows.Forms.DataVisualization.Charting.Series s in chart1.Series)
                    {
                        if (s.Points.Count > 0)
                        {
                            string seriesName = s.Name;

                            if (coauthorsDictionary.TryGetValue(seriesName, out List<string> coauthors))
                            {
                                // Используем HashSet для уникальных соавторов
                                HashSet<string> uniqueCoauthors = new HashSet<string>(coauthors);
                                listBox1.Items.Clear(); // Очищаем список перед добавлением новых элементов (если это необходимо)
                                foreach (string coauthor in uniqueCoauthors)
                                {
                                    if (coauthor != "")
                                    {
                                        listBox1.Items.Add(FormatCoauthor(coauthor)); // Добавляем каждого соавтора в ListBox
                                    }
                                }
                                y.UpdateChartForYears(seriesName);
                                p.UpdateVolumeChart(seriesName);
                                break;
                            }
                        }
                    }
                }
            }
        }
        private string FormatCoauthor(string name)
        {
            // Нормализуем имя
            string normalized = NormalizeCoauthorName(name);

            // Разбиваем на части
            string[] parts = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
                return name; // Если нет частей, возвращаем исходное имя

            // Предполагаем, что первая часть - фамилия, остальные части - инициалы
            string surname = parts[0]; // Фамилия
            string initials = string.Empty;

            // Форматируем фамилию с заглавной буквы
            string formattedSurname = char.ToUpper(surname[0]) + surname.Substring(1).ToLower(); // первая буква заглавная, остальные строчные

            // Проверяем наличие инициала
            if (parts.Length > 1) // Если есть инициалы
            {
                initials += char.ToUpper(parts[1][0]) + "."; // Первая буква инициала
                // Проверка на наличие второй буквы
                if (parts[1].Length > 2 && parts[1][1] == '.') // Проверяем, что после первой буквы есть точка
                {
                    initials += char.ToUpper(parts[1][2]) + "."; // Вторая буква инициала
                }

            }
            // Проверяем наличие третьего элемента
            if (parts.Length > 2) // Если есть третий элемент
            {
                initials += char.ToUpper(parts[2][0]) + "."; // Используем первую букву третьего элемента в качестве инициала
            }

            // Возвращаем форматированную строку
            return $"{formattedSurname} {initials.Trim()}".Trim();
        }


        private void ConvertExcelToCsv(string excelFilePath)
        {
            // Создаем экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false; // Открываем Excel в фоновом режиме

            // Открываем существующий Excel файл
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

            // Получаем первый лист в книге
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

            // Определяем путь для сохранения CSV файла
            string csvFilePath = System.IO.Path.ChangeExtension(excelFilePath, ".csv");

            // Сохраняем текущий лист как CSV файл
            worksheet.SaveAs(csvFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

            // Закрываем работу с Excel
            workbook.Close(false); // Закрываем workbook без сохранения
            excelApp.Quit(); // Закрываем приложение Excel

            // Освобождаем объекты
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // Подтверждаем успешное создание CSV файла
            MessageBox.Show($"Конвертация в CSV завершена! Файл сохранен как: {csvFilePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            y.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            p.ShowDialog();
        }
    }

}