using System.Windows.Forms;
using System;

namespace ExcelStudent
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private DataGridView dataGridView1;
        private Button btnExport;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.dataGridView1 = new DataGridView();
            this.btnExport = new Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();

            // dataGridView1
            this.dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(776, 250);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.ColumnCount = 12; // 1 фамилия + 11 оценок
            string[] columnHeaders = new string[]
            {
                "Фамилия   /   Умения", // 1-й столбец
                "Читать, записывать и сравнивать числа в пределах 1000000", // 2-й столбец
                "Складывать, вычитать, умножать и делить числа в пределах 1000000", // 3-й столбец
                "Находить значения выражений в 2-4 действия", // 4-й столбец
                "Сравнивать именованные числа и выполнять 4 арифметических действия с ними", // 5-й столбец
                "Читать и записывать именованные числа (длина, площадь, масса, объем)", // 6-й столбец
                "Читать информацию, заданную с помощью столбчатых, линейных и круговых диаграмм, таблиц, графов", // 7-й столбец
                "Переносить информацию из таблицы в линейные и столбчатые диаграммы", // 8-й столбец
                "Находить значение выражений с переменной", // 9-й столбец
                "Находить среднее арифметическое двух чисел", // 10-й столбец
                "Определять время по часам (до минуты)", // 11-й столбец
                "Сравнивать и упорядочивать объекты по разным признакам (длина, масса, объем)" // 12-й столбец
            };

            // Назначаем новые заголовки для столбцов
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                dataGridView1.Columns[i].Name = columnHeaders[i];
            }

            // btnExport
            this.btnExport.Location = new System.Drawing.Point(12, 280);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(150, 40);
            this.btnExport.TabIndex = 1;
            this.btnExport.Text = "Экспортировать в Excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new EventHandler(this.btnExport_Click);

            // Form1
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "ExcelStudent";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
        }
    }
}
