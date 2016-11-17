using Microsoft.Office.Interop.Excel;
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

namespace TidyBudgetData
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox1.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox2.Text = openFileDialog1.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox3.Text = openFileDialog1.FileName;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sErrorMessage = "";
            if (textBox1.Text == "")
                sErrorMessage = "Favor selecionar uma planilha que relacione os projetos aos tipos de ativo.";
            if (textBox2.Text == "")
                if(sErrorMessage=="")
                    sErrorMessage = "Favor selecionar uma planilha com os valores orçados.";
                else
                    sErrorMessage = sErrorMessage + Environment.NewLine + "Favor selecionar uma planilha com os valores orçados.";
            if (textBox3.Text == "")
                if (sErrorMessage == "")
                    sErrorMessage = "Favor selecionar uma planilha com os valores realizados.";
                else
                    sErrorMessage = sErrorMessage + Environment.NewLine + "Favor selecionar uma planilha com os valores realizados.";

            if(sErrorMessage!="")
                MessageBox.Show(sErrorMessage,"Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int iRowIndex = 0;
                Cursor.Current = Cursors.WaitCursor;
                string sCurrProj = "";
                string sCurrType = "";

                System.Data.DataTable tblProjTypeOfAsset = new System.Data.DataTable();
                tblProjTypeOfAsset.Columns.Add("Code", typeof(string));
                tblProjTypeOfAsset.Columns.Add("Desc", typeof(string));
                tblProjTypeOfAsset.Columns.Add("Type", typeof(string));

                DirectoryInfo rootFolder = new DirectoryInfo(textBox1.Text);
                Microsoft.Office.Interop.Excel.Application xlSourceApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Application xlTargetApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlSourceWorkBook;
                Workbook xlTargetWorkBook;
                Worksheet xlSourceWorkSheet;
                Worksheet xlTargetWorkSheet;

                // Delete file if it exists, and create a new, empty one
                if (File.Exists("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx"))
                {
                    File.Delete("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx");
                }

                xlTargetWorkBook = xlTargetApp.Workbooks.Add();
                xlTargetWorkBook.SaveAs("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx");

                // Reading worksheet that relates projects / actions to types of asset
                xlSourceWorkBook = xlSourceApp.Workbooks.Open(textBox1.Text);
                xlSourceWorkSheet = xlSourceWorkBook.Worksheets[1];

                for (iRowIndex = 1; iRowIndex <= xlSourceWorkSheet.UsedRange.Rows.Count; iRowIndex++)
                {
                    sCurrProj = xlSourceWorkSheet.Cells[iRowIndex, 1].Value;
                    if(sCurrProj.Substring(2,1)=="_")   
                    {
                        sCurrType = xlSourceWorkSheet.Cells[iRowIndex, 2].Value;
                        tblProjTypeOfAsset.Rows.Add(sCurrProj.Substring(0, 5), sCurrProj.Substring(8), sCurrType);
                    }
                }

                xlSourceApp.Quit();
                xlTargetWorkBook.Save();
                xlTargetApp.Quit();
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Data tidied up", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
