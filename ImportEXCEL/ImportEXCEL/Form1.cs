using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using SpreadsheetLight;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace ImportEXCEL
{
    public partial class Form1 : Form
    {      
        private string path = @"C:\Users\Usuario\Downloads\ImportEXCEL\Inventario de refacciones.xlsx";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SLDocument sLDocument = new SLDocument(path);
            int iRow = 2;
            List<ExcelViewModel> lst = new List<ExcelViewModel>();
            while (!string.IsNullOrEmpty(sLDocument.GetCellValueAsString(iRow, 1)))
            {
                ExcelViewModel ObjexcelViewModel = new ExcelViewModel();
                ObjexcelViewModel.Part_Number = sLDocument.GetCellValueAsString(iRow, 1);
                ObjexcelViewModel.Supplier_part_number = sLDocument.GetCellValueAsString(iRow, 2);
                ObjexcelViewModel.Description_Material = sLDocument.GetCellValueAsString(iRow, 3);
                ObjexcelViewModel.Supplier = sLDocument.GetCellValueAsInt32(iRow, 4);
                ObjexcelViewModel.Min_Stock = sLDocument.GetCellValueAsInt32(iRow, 5);
                ObjexcelViewModel.Max_Stock = sLDocument.GetCellValueAsInt32(iRow, 6);
                lst.Add(ObjexcelViewModel);
                iRow++;               
            }
           
            dataGridView1.DataSource = lst;
            using (SqlConnection conn = new SqlConnection("Data Source = DESKTOP-HH29OP1\\SQLEXPRESS;initial Catalog=Proyecto_Seguridad_Web; Integrated Security=True;"))
            {            
                conn.Open();

                string query = "INSERT INTO Info_Part_Numbers (Part_Number, Supplier_part_number, Description_Material, Supplier, Min_Stock, Max_Stock) VALUES(@Part_Number, @Supplier_part_number, @Description_Material, @Supplier, @Min_Stock, @Max_Stock)";
                SqlCommand cmd = new SqlCommand();
                string cadena = "";
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    cadena = row.Cells[0].Value.ToString();
                    cadena = cadena.Replace("-", string.Empty);
                    cadena = "00000" + cadena;
                    cmd = new SqlCommand(query, conn);
                    cmd.Parameters.Add("@Part_Number", SqlDbType.VarChar).Value = cadena;
                    cmd.Parameters.Add("@Supplier_part_number", SqlDbType.VarChar).Value = row.Cells[1].Value;
                    cmd.Parameters.Add("@Description_Material", SqlDbType.VarChar).Value = row.Cells[2].Value;
                    cmd.Parameters.Add("@Supplier", SqlDbType.VarChar).Value = row.Cells[3].Value;
                    cmd.Parameters.Add("@Min_Stock", SqlDbType.Int).Value = row.Cells[4].Value;
                    cmd.Parameters.Add("@Max_Stock", SqlDbType.Int).Value = row.Cells[5].Value;
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
            }           
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }
    }
}
