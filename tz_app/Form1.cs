using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Collections;
using System.Linq;

namespace tz_app
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection;
        private SqlDataAdapter sqlDataAdapter;
        private DataTable dataTable;

        public Form1()
        {
            InitializeComponent();

            sqlConnection = new SqlConnection("Data Source=DESKTOP-97LV1Q7\\SQLEXPRESS;Initial Catalog=intermech;Integrated Security=True;TrustServerCertificate=True");

            FillComboBox();

            comboBox1.SelectedIndexChanged += (s, e) =>
            {
                if (comboBox1.SelectedItem != null)
                {
                    FillDataGridView();
                }
            };
        }

        private void FillComboBox()
        {
            // using для автоматического закрытия соединений
            using (SqlConnection conn1 = new SqlConnection("Data Source=DESKTOP-97LV1Q7\\SQLEXPRESS;Initial Catalog=intermech;Integrated Security=True;TrustServerCertificate=True"))
            {
                conn1.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT TABLE_NAME FROM intermech.INFORMATION_SCHEMA.TABLES", conn1);
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                ArrayList list = new ArrayList();
                while (sqlDataReader.Read())
                {
                    string tableName = sqlDataReader[0].ToString();
                    // Два соединения, одно для чтения имен таблиц, а другое для подсчета строк 
                    using (SqlConnection conn2 = new SqlConnection("Data Source=DESKTOP-97LV1Q7\\SQLEXPRESS;Initial Catalog=intermech;Integrated Security=True;TrustServerCertificate=True"))
                    {
                        conn2.Open();
                        SqlCommand countCommand = new SqlCommand($"SELECT COUNT(*) FROM [{tableName}]", conn2);
                        int rowCount = (int)countCommand.ExecuteScalar();
                        if (rowCount > 0)
                        {
                            list.Add(tableName);
                        }
                    }
                }
                list.Sort();
                foreach (var item in list)
                {
                    comboBox1.Items.Add(item);
                }
            }
        }

        private void FillDataGridView()
        {
            button2.Visible = false;
            sqlDataAdapter = new SqlDataAdapter($"SELECT * FROM [{comboBox1.SelectedItem}]", sqlConnection);
            dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;


            sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand($"SELECT COUNT(*) FROM [{comboBox1.SelectedItem}]", sqlConnection);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            while (sqlDataReader.Read())
            {
                label1.Text = "Всего: " + sqlDataReader[0].ToString();
            }
            sqlConnection.Close();
        }

        private void SearchAndFillDataGridView(string searchValue)
        {
            string query = "SELECT R.F_PROJ_ID, O.F_OBJECT_ID, R.F_RELATION_TYPE, T.F_DESCRIPTION, T.F_TYPE_NAME, T.F_NOTE, O.F_OBJECT_TYPE, OT.F_OBJ_TYPE_NAME, A.F_ATTRIBUTE_LIST, B.F_VALUE_LIST " +
                           "FROM [IMS_RELATIONS] R " +
                           "INNER JOIN [IMS_RELATION_TYPES] T ON R.F_RELATION_TYPE = T.F_RELATION_TYPE " +
                           "INNER JOIN [IMS_OBJECTS] O ON R.F_PART_ID = O.F_ID " +
                           "INNER JOIN [IMS_OBJECT_TYPES] OT ON O.F_OBJECT_TYPE = OT.F_OBJECT_TYPE " +
                           "LEFT JOIN (" +
                           "    SELECT AO.F_OBJECT_ID, STRING_AGG(AT.F_NAME, ', ') AS F_ATTRIBUTE_LIST " +
                           "    FROM [IMS_OBJECT_ATTRS] AO " +
                           "    INNER JOIN [IMS_ATTRIBUTES] AT ON AO.F_ATTRIBUTE_ID = AT.F_ATTRIBUTE_ID " +
                           "    GROUP BY AO.F_OBJECT_ID " +
                           ") A ON O.F_OBJECT_ID = A.F_OBJECT_ID " +
                           "LEFT JOIN (" +
                           "    SELECT AO.F_OBJECT_ID, STRING_AGG(CONCAT(AO.F_INTEGER_VALUE, ', ', AO.F_STRING_VALUE), '; ') AS F_VALUE_LIST " +
                           "    FROM [IMS_OBJECT_ATTRS] AO " +
                           "    GROUP BY AO.F_OBJECT_ID " +
                           ") B ON O.F_OBJECT_ID = B.F_OBJECT_ID " +
                           "WHERE R.[F_PROJ_ID] = @searchValue";


            SqlCommand command = new SqlCommand(query, sqlConnection);
            command.Parameters.AddWithValue("@searchValue", searchValue);

            sqlDataAdapter = new SqlDataAdapter(command);
            dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;
            sqlConnection.Close();


            sqlConnection.Open();
            SqlCommand sqlCommand = new SqlCommand($"SELECT COUNT(*) FROM [IMS_RELATIONS] WHERE [F_PROJ_ID] = '{searchValue}'", sqlConnection);
            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            while (sqlDataReader.Read())
            {
                label1.Text = "Всего: " + sqlDataReader[0].ToString();
            }
            sqlConnection.Close();
            button2.Visible = true;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text))
            {
                SearchAndFillDataGridView(textBox2.Text);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);

            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Лист1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Объект";

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        string cellValue = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        if (dataGridView1.Columns[j].Name == "F_VALUE_LIST")
                        {
                            string[] elements = cellValue.Split(new[] { "; , " }, StringSplitOptions.None);
                            elements = elements.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
                            cellValue = string.Join("; , ", elements);
                        }
                        worksheet.Cells[i + 2, j + 1] = cellValue;
                    }
                    else
                    {
                        worksheet.Cells[i + 2, j + 1] = "";
                    }
                }
            }


            dataGridView1.Columns["F_PROJ_ID"].HeaderText = "Айди родительского объекта";
            dataGridView1.Columns["F_OBJECT_ID"].HeaderText = "Дочерние объекты";
            dataGridView1.Columns["F_RELATION_TYPE"].HeaderText = "Номер типа связи";
            dataGridView1.Columns["F_DESCRIPTION"].HeaderText = "Описание типа связи";
            dataGridView1.Columns["F_TYPE_NAME"].HeaderText = "Тип связи";
            dataGridView1.Columns["F_NOTE"].HeaderText = "Примечание к типу связи";
            dataGridView1.Columns["F_OBJECT_TYPE"].HeaderText = "Номер типа объекта";
            dataGridView1.Columns["F_OBJ_TYPE_NAME"].HeaderText = "Имя объекта";
            dataGridView1.Columns["F_ATTRIBUTE_LIST"].HeaderText = "Список атрибутов объекта";
            dataGridView1.Columns["F_VALUE_LIST"].HeaderText = "Список значений атрибутов объекта";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }


            excel.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Range range = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[dataGridView1.Rows.Count + 1, 1]];
            range.Merge();
            excel.DisplayAlerts = true;


            worksheet.Columns.AutoFit();
            worksheet.Range["A:J"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range["A:J"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet.Columns["F:F"].ColumnWidth = 60;
            worksheet.Columns["I:I"].ColumnWidth = 60;
            worksheet.Columns["J:J"].ColumnWidth = 60;

            worksheet.Columns["F:F"].WrapText = true;
            worksheet.Columns["I:I"].WrapText = true;
            worksheet.Columns["J:J"].WrapText = true;

            worksheet.Columns["F:F"].EntireRow.AutoFit();
            worksheet.Columns["I:I"].EntireRow.AutoFit();
            worksheet.Columns["J:J"].EntireRow.AutoFit();


            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel Files|*.xlsx";
            saveDialog.Title = "Сохранить как Excel файл";
            saveDialog.ShowDialog();

            if (saveDialog.FileName != "")
            {
                workbook.SaveAs(saveDialog.FileName);
            }

            excel.Quit();
        }
    }
}
