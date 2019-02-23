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
using System.Data.SqlClient;
using System.Configuration;
using System.Data.Common;
using System.IO;

namespace test005
{
    public partial class Form1 : Form
    {
        
        OleDbConnection OleDbcon;
        string fl;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";

           
            openFileDialog.ShowDialog();

            string fileName = Path.GetFileName(openFileDialog.FileName);
            fl = fileName.Substring(0, fileName.IndexOf("."));
            



            if (!string.IsNullOrEmpty(openFileDialog.FileName))
               
            {

                OleDbcon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog.FileName + ";Extended Properties=Excel 12.0;");
                
                OleDbcon.Open();

                DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                OleDbcon.Close();

                comboBox1.Items.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)

                {

                    String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
                   
                    sheetName = sheetName.Substring(0, sheetName.Length - 1);

                    comboBox1.Items.Add(sheetName);

                }

            }

        }
        private DataTable Get_table()
        {
            OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from [" + comboBox1.Text + "$]", OleDbcon);

            DataTable dt = new DataTable();

            oledbDa.Fill(dt);

            
            return dt;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            dataGridView1.DataSource = Get_table(); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataTable salesData = Get_table();

            using (SqlConnection sqlconnection = new SqlConnection(@"Data Source=localhost;User ID=sa;Password=sa123"))
            {
                sqlconnection.Open();

                string createTableQuery = ""; 
                try
                {

                    createTableQuery = @"Create Table "+comboBox1.Text+fl+" ( " ; 
                   

                            for (int j = 0; j < salesData.Columns.Count; j++)
                            {
                                createTableQuery += salesData.Columns[j].ColumnName.Trim() + " nvarchar(1000)";
                             
                                if (j < salesData.Columns.Count - 1) 
                                {
                                    createTableQuery += ",";
                                   

                                }
                                if (j == salesData.Columns.Count -1)
                                {
                                    createTableQuery += ")";
                                }
                 
                            }
                            

                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error");
                }
                
          
                     
                SqlCommand command = new SqlCommand(createTableQuery, sqlconnection);

                command.ExecuteNonQuery();

                
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(sqlconnection))
                {
                    sqlBulkCopy.DestinationTableName = comboBox1.Text+fl;

                    foreach (var column in salesData.Columns)
                        sqlBulkCopy.ColumnMappings.Add(column.ToString(), column.ToString());

                    sqlBulkCopy.WriteToServer(salesData);
                    MessageBox.Show("Upload Success Fully");
                    
                    sqlconnection.Close();

                    if (System.Windows.Forms.Application.MessageLoop)
                    {
                        
                        System.Windows.Forms.Application.Exit();
                    }
                    else
                    {
                        
                        System.Environment.Exit(1);
                    }
                }
            }

        }
        private static DataTable GetSalesData()
        {
            DataTable salesHistory = new DataTable("SalesHistory");

            // Create Column 1: SaleDate
            DataColumn dateColumn = new DataColumn();
            dateColumn.DataType = Type.GetType("System.DateTime");
            dateColumn.ColumnName = "SaleDate";

            // Create Column 2: ItemName
            DataColumn productNameColumn = new DataColumn();
            productNameColumn.ColumnName = "ItemName";

            // Create Column 3: ItemsCount
            DataColumn totalSalesColumn = new DataColumn();
            totalSalesColumn.DataType = Type.GetType("System.Int32");
            totalSalesColumn.ColumnName = "ItemsCount";

            // Add the columns to the SalesHistory DataTable
            salesHistory.Columns.AddRange(new DataColumn[] { dateColumn, productNameColumn, totalSalesColumn });


            // Let's populate the datatable with our stats.
            // You can add as many rows as you want here!

            // Create a new row
            DataRow dailyProductSalesRow = salesHistory.NewRow();
            dailyProductSalesRow["SaleDate"] = DateTime.Now.Date;
            dailyProductSalesRow["ItemName"] = "Nike Shoe-32";
            dailyProductSalesRow["ItemsCount"] = 10;

            // Add the row to the SalesHistory DataTable
            salesHistory.Rows.Add(dailyProductSalesRow);
            return salesHistory;
        }

    }
}











