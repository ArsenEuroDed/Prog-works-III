using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;

namespace Kursovaya_rabota
{
    public partial class Year : Form
    {
        public Year()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }
        private Dictionary<string, Dictionary<int, int>> publicationCountsBySeries = new Dictionary<string, Dictionary<int, int>>();

        // Новый метод для обновления количества публикаций
        public void UpdatePublicationCounts(string excelFilePath)
        {
            // Получаем имя файла без расширения
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelFilePath);

            // Инициализируем словарь для этой серии, если его еще нет
            if (!publicationCountsBySeries.ContainsKey(fileNameWithoutExtension))
            {
                publicationCountsBySeries[fileNameWithoutExtension] = new Dictionary<int, int>();
            }

            // Открываем Excel и получаем данные из листа
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            int lastRow = worksheet.Cells[worksheet.Rows.Count, 4].End[XlDirection.xlUp].Row;

            int currentYear = DateTime.Now.Year; // Получаем текущий год

            for (int i = 2; i <= lastRow; i++)
            {
                string publicationCell = ((Range)worksheet.Cells[i, 4]).Text;

                // Находим год в строке
                Match yearMatch = Regex.Match(publicationCell, @"\b(18|19|20)\d{2}\b");
                if (yearMatch.Success)
                {
                    int year = int.Parse(yearMatch.Value);

                    // Проверяем, что год находится в диапазоне от 1800 до текущего года
                    if (year >= 1800 && year <= currentYear)
                    {
                        if (publicationCountsBySeries[fileNameWithoutExtension].ContainsKey(year))
                        {
                            publicationCountsBySeries[fileNameWithoutExtension][year]++;
                        }
                        else
                        {
                            publicationCountsBySeries[fileNameWithoutExtension][year] = 1;
                        }
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
        }

        // Метод для обновления chart2 на основе словаря
        public void UpdateChartForYears(string seriesName)
        {
            chart1.Series.Clear();
            System.Windows.Forms.DataVisualization.Charting.Series series = new System.Windows.Forms.DataVisualization.Charting.Series(seriesName)
            {
                ChartType = SeriesChartType.Column
            };
            chart1.Series.Add(series);

            if (publicationCountsBySeries.TryGetValue(seriesName, out Dictionary<int, int> publicationCounts))
            {
                foreach (KeyValuePair<int, int> entry in publicationCounts.OrderBy(kvp => kvp.Key))
                {
                    series.Points.AddXY(entry.Key, entry.Value);
                }
            }

            chart1.Invalidate();
        }
    }
}
