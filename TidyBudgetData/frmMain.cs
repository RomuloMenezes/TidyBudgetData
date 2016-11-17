using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
            else{

            }
        }
    }
}
