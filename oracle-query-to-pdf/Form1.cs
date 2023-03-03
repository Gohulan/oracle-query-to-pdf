using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace oracle_query_to_pdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OracleConnection conn = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=mcrspos)));User Id=dbcreate;Password=Sup3rAdm1n#;");
            conn.Open();

            OracleCommand cmd = new OracleCommand("SELECT * FROM INTERFACE", conn);

            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);

            Document doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream("output.pdf", FileMode.Create));
            doc.Open();

            PdfPTable table = new PdfPTable(ds.Tables[0].Columns.Count);

            // Set table headers
            foreach (DataColumn col in ds.Tables[0].Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(col.ColumnName));
                table.AddCell(cell);
            }

            // Add rows
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                foreach (DataColumn col in ds.Tables[0].Columns)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(row[col].ToString()));
                    table.AddCell(cell);
                }
            }

            doc.Add(table);
            doc.Close();
            conn.Close();

            MessageBox.Show("PDF created successfully.");




        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
    }

