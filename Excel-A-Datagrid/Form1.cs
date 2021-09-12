using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Excel_A_Datagrid
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string ExcelConnection(string fileName)
        {
            return @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   @"Data Source=" + fileName + ";" +
                   @"Extended Properties=" + Convert.ToChar(34).ToString() +
                   @"Excel 12.0" + Convert.ToChar(34).ToString() + ";";
        }

        DataView DataImport(string filename)
        {
            var excelconnection = ExcelConnection(filename);
            //string connection = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", filename);
            OleDbConnection connector = new OleDbConnection(excelconnection);
            connector.Open();
            OleDbCommand query = new OleDbCommand("select * from [Hoja1$]", connector);
            OleDbDataAdapter adapter = new OleDbDataAdapter
            {
                SelectCommand = query
            };
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            connector.Close();
            return dataset.Tables[0].DefaultView;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfiledialog = new OpenFileDialog
            {
                Filter = "Excel | *.xls;*xlsx;*.csv;",
                Title = "Seleccionar Archivo"
            };

            if (openfiledialog.ShowDialog() == DialogResult.OK)
            {
                DGVDataImport.DataSource = DataImport(openfiledialog.FileName);
            }
        }
    }
}
