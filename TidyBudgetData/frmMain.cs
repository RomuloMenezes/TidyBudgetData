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
                int iYearIndex = 0;
                Cursor.Current = Cursors.WaitCursor;
                string sCurrCell = "";
                string sCurrProj = "";
                string sCurrType = "";
                string sCurrAction = "";
                int iCurrYear = 0;
                DateTime dtCurrDate;
                double dCurrValue = 0;
                int iNbOfYears = 0;

                textBox4.Text = "Inicializando - criando estruturas auxiliares.";
                textBox4.Refresh();

                DirectoryInfo rootFolder = new DirectoryInfo(textBox1.Text);
                Microsoft.Office.Interop.Excel.Application xlSourceApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Application xlTargetApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlTypesOfAssetWorkBook;
                Worksheet xlTypesOfAssetWorkSheet;
                Workbook xlBudgetWorkBook;
                Worksheet xlBudgetWorkSheet;
                Workbook xlActualExpensesWorkBook;
                Worksheet xlActualExpensesWorkSheet;
                Workbook xlTargetWorkBook;
                Worksheet xlTargetWorkSheet;

                System.Data.DataTable tblProjTypeOfAsset = new System.Data.DataTable();
                tblProjTypeOfAsset.Columns.Add("Code", typeof(string));
                tblProjTypeOfAsset.Columns.Add("Desc", typeof(string));
                tblProjTypeOfAsset.Columns.Add("Type", typeof(string));

                System.Data.DataTable tblActualExpenses = new System.Data.DataTable();
                tblActualExpenses.Columns.Add("Project", typeof(string));
                tblActualExpenses.Columns.Add("Action", typeof(string));
                tblActualExpenses.Columns.Add("Year", typeof(DateTime));
                tblActualExpenses.Columns.Add("Value", typeof(float));

                // Delete file if it exists, and create a new, empty one
                if (File.Exists("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx"))
                {
                    File.Delete("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx");
                }

                xlTargetWorkBook = xlTargetApp.Workbooks.Add();
                xlTargetWorkBook.SaveAs("D:\\_GIT\\Projetos\\GIT - Orçamento\\Plano de Ação\\TidyData.xlsx");

                // Reading worksheet that relates projects / actions to types of asset
                textBox4.Text = "Lendo tipos de ativo.";
                textBox4.Refresh();

                xlTypesOfAssetWorkBook = xlSourceApp.Workbooks.Open(textBox1.Text);
                xlTypesOfAssetWorkSheet = xlTypesOfAssetWorkBook.Worksheets[1];

                for (iRowIndex = 1; iRowIndex <= xlTypesOfAssetWorkSheet.UsedRange.Rows.Count; iRowIndex++)
                {
                    sCurrProj = xlTypesOfAssetWorkSheet.Cells[iRowIndex, 1].Value;
                    if(sCurrProj.Substring(2,1)=="_")   
                    {
                        sCurrType = xlTypesOfAssetWorkSheet.Cells[iRowIndex, 2].Value;
                        tblProjTypeOfAsset.Rows.Add(sCurrProj.Substring(0, 5), sCurrProj.Substring(8), sCurrType);
                    }
                }

                System.Data.DataTable tblBudget = new System.Data.DataTable();
                tblBudget.Columns.Add("Project", typeof(string));
                tblBudget.Columns.Add("Action", typeof(string));
                tblBudget.Columns.Add("Year", typeof(DateTime));
                tblBudget.Columns.Add("Value", typeof(float));

                // Reading worksheet with the budgeted values
                textBox4.Text = "Lendo itens orçados.";
                textBox4.Refresh();

                xlBudgetWorkBook= xlSourceApp.Workbooks.Open(textBox2.Text);
                xlBudgetWorkSheet = xlBudgetWorkBook.Worksheets[1];
                iNbOfYears = xlBudgetWorkSheet.UsedRange.Columns.Count - 2;

                for (iRowIndex = 1; iRowIndex <= xlBudgetWorkSheet.UsedRange.Rows.Count; iRowIndex++)
                {
                    sCurrCell = xlBudgetWorkSheet.Cells[iRowIndex, 1].Value;
                    if (sCurrCell != null)
                    {
                        if (sCurrCell.Substring(2, 1) == "_" || sCurrCell.Substring(4, 1) == "-") // Either a project or an action
                        {
                            if (sCurrCell.Substring(2, 1) == "_") // It's a project
                                sCurrProj = sCurrCell;
                            else // It's an action
                            {
                                sCurrAction = sCurrCell;
                                for (iYearIndex = 0; iYearIndex < iNbOfYears; iYearIndex++)
                                {
                                    iCurrYear = Convert.ToInt16(xlBudgetWorkSheet.Cells[1, 2 + iYearIndex].Value);
                                    dtCurrDate = new DateTime(iCurrYear, 1, 1);
                                    if (xlBudgetWorkSheet.Cells[iRowIndex, 2 + iYearIndex].Value != null)
                                        dCurrValue = xlBudgetWorkSheet.Cells[iRowIndex, 2 + iYearIndex].Value;
                                    else
                                        dCurrValue = 0;
                                    tblBudget.Rows.Add(sCurrProj, sCurrAction, dtCurrDate, dCurrValue);
                                }
                            }
                        }
                    }
                }

                // Reading worksheet with the actual expenses
                textBox4.Text = "Lendo itens realizados.";
                textBox4.Refresh();

                xlActualExpensesWorkBook = xlSourceApp.Workbooks.Open(textBox3.Text);
                xlActualExpensesWorkSheet = xlActualExpensesWorkBook.Worksheets[1];
                iNbOfYears = xlActualExpensesWorkSheet.UsedRange.Columns.Count - 2;

                for (iRowIndex = 1; iRowIndex <= xlActualExpensesWorkSheet.UsedRange.Rows.Count; iRowIndex++)
                {
                    sCurrCell = xlActualExpensesWorkSheet.Cells[iRowIndex, 1].Value;
                    if (sCurrCell != null)
                    {
                        if (sCurrCell.Substring(2, 1) == "_" || sCurrCell.Substring(4, 1) == "-") // Either a project or an action
                        {
                            if (sCurrCell.Substring(2, 1) == "_") // It's a project
                                sCurrProj = sCurrCell;
                            else // It's an action
                            {
                                sCurrAction = sCurrCell;
                                for (iYearIndex = 0; iYearIndex < iNbOfYears; iYearIndex++)
                                {
                                    iCurrYear = Convert.ToInt16(xlActualExpensesWorkSheet.Cells[1, 2 + iYearIndex].Value);
                                    dtCurrDate = new DateTime(iCurrYear, 1, 1);
                                    if (xlActualExpensesWorkSheet.Cells[iRowIndex, 2 + iYearIndex].Value != null)
                                        dCurrValue = xlActualExpensesWorkSheet.Cells[iRowIndex, 2 + iYearIndex].Value;
                                    else
                                        dCurrValue = 0;
                                    tblActualExpenses.Rows.Add(sCurrProj, sCurrAction, dtCurrDate, dCurrValue);
                                }
                            }
                        }
                    }
                }

                textBox4.Text = "";
                textBox4.Refresh();
                xlSourceApp.Quit();
                xlTargetWorkBook.Save();
                xlTargetApp.Quit();
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Data tidied up", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
