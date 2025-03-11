using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Kursovaya_rabota
{
    public partial class Publish : Form
    {
        public Publish()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }
        private Dictionary<string, Dictionary<string, List<double>>> volumeCountsBySeries = new Dictionary<string, Dictionary<string, List<double>>>();

        public void UpdateVolumeCounts(string excelFilePath)
        {
            // Получаем имя файла без расширения
            string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(excelFilePath);

            // Инициализируем словарь для этой серии, если его еще нет
            if (!volumeCountsBySeries.ContainsKey(fileNameWithoutExtension))
            {
                volumeCountsBySeries[fileNameWithoutExtension] = new Dictionary<string, List<double>>();
            }

            // Открываем Excel и получаем данные из листа
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            int lastRow = worksheet.Cells[worksheet.Rows.Count, 5].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row; // Столбец E

            for (int i = 2; i <= lastRow; i++)
            {
                string titleCell = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[i, 2]).Text; // Столбец B
                string volumeCell = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[i, 5]).Text; // Столбец E

                double volumeValue;
                if (TryParseVolume(volumeCell, out volumeValue))
                {
                    if (!volumeCountsBySeries[fileNameWithoutExtension].ContainsKey(titleCell))
                    {
                        volumeCountsBySeries[fileNameWithoutExtension][titleCell] = new List<double>(); // Инициализируем новый список
                    }
                    volumeCountsBySeries[fileNameWithoutExtension][titleCell].Add(volumeValue); // Добавляем значение в список
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

        // Проверяет и извлекает объём из строки
        public bool TryParseVolume(string input, out double volume)
        {

            // Убираем пробелы и нечисловые символы
            input = input.Trim();
            input = input.Replace("стр.", "").Replace("с.", "").Replace("п.л.", "").Replace(" ", ""); // Нормализация формата
            // Пробуем конвертировать в число
            return double.TryParse(input, out volume);
        }
        public void UpdateVolumeChart(string seriesName)
        {
            chart1.Series.Clear();
            System.Windows.Forms.DataVisualization.Charting.Series series = new System.Windows.Forms.DataVisualization.Charting.Series(seriesName)
            {
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true // Включаем отображение значения над столбцом
            };
            chart1.Series.Add(series);

            if (volumeCountsBySeries.TryGetValue(seriesName, out Dictionary<string, List<double>> volumes))
            {
                foreach (KeyValuePair<string, List<double>> entry in volumes)
                {
                    double totalVolume = entry.Value.Sum(); // Суммируем все объемы для данной работы
                    series.Points.AddXY(entry.Key, totalVolume); // entry.Key - название работы, totalVolume - сумма объемов
                }
            }

            chart1.ChartAreas[0].AxisX.Title = "Название работы";
            chart1.ChartAreas[0].AxisY.Title = "Объем в печатных страницах";

            // Установка формата меток оси Y
            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "F2"; // Формат с двумя знаками после запятой

            // Настройки осей
            chart1.ChartAreas[0].AxisX.Interval = 1; // Отображаем все метки по оси X
            chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; // Отключаем сетку

            // Установите минимальное и максимальное значение оси Y, если необходимо
            chart1.ChartAreas[0].AxisY.Minimum = 0; // Минимальное значение
            chart1.ChartAreas[0].AxisY.Maximum = 10; // Установите максимальное значение, соответствующее вашим данным

            chart1.Invalidate(); // Обновляем график
        }

    }
}
