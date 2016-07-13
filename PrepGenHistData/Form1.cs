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

namespace PrepGenHistData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                MessageBox.Show("Please select a source folder", "Error", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error);
            else
            {
                DirectoryInfo rootFolder = new DirectoryInfo(textBox1.Text);
                Microsoft.Office.Interop.Excel.Application xlSourceApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Application xlTargetApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlSourceWorkBook;
                Workbook xlTargetWorkBook;
                Worksheet xlCurrSourceWorkSheet;
                Worksheet xlTargetWorkSheet;
                int iIndex = 0;
                int xIndex, yIndex;
                int year;
                int targetWBRowIndex = 2;
                int yUpperLimit;
                int yOffset;
                string currMessage = "";

                xlTargetWorkBook = xlTargetApp.Workbooks.Open("D:\\_GIT\\Projetos\\Novo Site\\planilhas\\geração\\SínteseGeração.xlsx");
                xlTargetWorkSheet = xlTargetWorkBook.Worksheets[1];
                xlTargetWorkSheet.Cells[1, 1] = "Data";
                xlTargetWorkSheet.Cells[1, 2] = "Tipo de geração";
                xlTargetWorkSheet.Cells[1, 3] = "Região";
                xlTargetWorkSheet.Cells[1, 4] = "Unidade";
                xlTargetWorkSheet.Cells[1, 5] = "Montante";

                foreach(DirectoryInfo currFolder in rootFolder.GetDirectories("geracao_*")){
                    if (currFolder.Name.IndexOf("nuclear") > 0){
                        yUpperLimit = 2;
                        yOffset = 5;
                    }
                    else{
                        yUpperLimit = 7;
                        yOffset = 10;
                    }
                    xlSourceWorkBook = xlSourceApp.Workbooks.Open(currFolder.GetFiles()[0].FullName);
                    for (iIndex = 1; iIndex <= xlSourceWorkBook.Worksheets.Count; iIndex++) {
                        xlCurrSourceWorkSheet = xlSourceWorkBook.Worksheets[iIndex];
                        if (xlCurrSourceWorkSheet.Cells[1, 1].Value != null){
                            year = Convert.ToInt16(xlCurrSourceWorkSheet.Cells[2, 1].Value);
                            currMessage = Convert.ToString(year) + Environment.NewLine + currFolder.Name.Substring(8);
                            textBox2.Text = currMessage;
                            textBox2.Refresh();
                            for (yIndex = 1; yIndex <= yUpperLimit; yIndex++) {
                                for (xIndex = 1; xIndex <= 12; xIndex++) {
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = Convert.ToString(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name.Substring(8);
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 1].Value;
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2].Value;
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2 + xIndex].Value;
                                    targetWBRowIndex++;
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value + "/" + Convert.ToString(year);
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name.Substring(8);
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[yOffset + yIndex, 1].Value;
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[yOffset + yIndex, 2].Value;
                                    xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[yOffset + yIndex, 2 + xIndex].Value;
                                    targetWBRowIndex++;
                                }
                            }
                        }
                    }
                }
                xlSourceApp.Quit();
                xlTargetWorkBook.Save();
                xlTargetApp.Quit();
            }
        }
    }
}
