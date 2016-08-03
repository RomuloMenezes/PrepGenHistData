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
                MessageBox.Show("Please select a source folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                if (!radioButton1.Checked && !radioButton2.Checked && !radioButton3.Checked)
                {
                    MessageBox.Show("Please select a type of historical data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    Cursor.Current = Cursors.WaitCursor;

                    DirectoryInfo rootFolder = new DirectoryInfo(textBox1.Text);
                    Microsoft.Office.Interop.Excel.Application xlSourceApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Application xlTargetApp = new Microsoft.Office.Interop.Excel.Application();
                    Workbook xlSourceWorkBook;
                    Workbook xlTargetWorkBook;
                    Worksheet xlCurrSourceWorkSheet;
                    Worksheet xlTargetWorkSheet;

                    int yUpperLimit;
                    int iIndex = 0;
                    int xIndex, yIndex;
                    int year;
                    int targetWBRowIndex = 2;
                    string currMessage = "";

                    // Delete file if it exists, and create a new, empty one
                    if (File.Exists("D:\\_GIT\\Projetos\\Novo Site\\Planilhas histórico\\Tableau\\TidyData.xlsx"))
                    {
                        File.Delete("D:\\_GIT\\Projetos\\Novo Site\\Planilhas histórico\\Tableau\\TidyData.xlsx");
                    }

                    xlTargetWorkBook = xlTargetApp.Workbooks.Add();
                    xlTargetWorkBook.SaveAs("D:\\_GIT\\Projetos\\Novo Site\\Planilhas histórico\\Tableau\\TidyData.xlsx");

                    xlTargetWorkSheet = xlTargetWorkBook.Worksheets[1];
                    xlTargetWorkSheet.Cells[1, 1] = "Data";
                    xlTargetWorkSheet.Cells[1, 2] = "Medida";
                    xlTargetWorkSheet.Cells[1, 3] = "Região";
                    xlTargetWorkSheet.Cells[1, 4] = "Unidade";
                    xlTargetWorkSheet.Cells[1, 5] = "Montante";

                    if (radioButton1.Checked) // Geração
                    {
                        int yOffset;
                        foreach (DirectoryInfo currFolder in rootFolder.GetDirectories("geracao_*"))
                        {
                            if (currFolder.Name.IndexOf("nuclear") > 0)
                            {
                                yUpperLimit = 2;
                                yOffset = 5;
                            }
                            else
                            {
                                if (currFolder.Name.IndexOf("hidraulica") > 0) // Inclui o subsistema Itaipu, portanto um a mais que os demais tipos de geração
                                {
                                    yUpperLimit = 8;
                                    yOffset = 11;
                                }
                                else
                                {
                                    yUpperLimit = 7;
                                    yOffset = 10;
                                }
                            }
                            xlSourceWorkBook = xlSourceApp.Workbooks.Open(currFolder.GetFiles()[0].FullName);
                            for (iIndex = 1; iIndex <= xlSourceWorkBook.Worksheets.Count; iIndex++)
                            {
                                xlCurrSourceWorkSheet = xlSourceWorkBook.Worksheets[iIndex];
                                if (xlCurrSourceWorkSheet.Cells[1, 1].Value != null)
                                {
                                    year = Convert.ToInt16(xlCurrSourceWorkSheet.Cells[2, 1].Value);
                                    currMessage = Convert.ToString(year) + Environment.NewLine + currFolder.Name.Substring(8);
                                    textBox2.Text = currMessage;
                                    textBox2.Refresh();
                                    for (yIndex = 1; yIndex <= yUpperLimit; yIndex++)
                                    {
                                        for (xIndex = 1; xIndex <= 12; xIndex++)
                                        {
                                            // Data - Medida (emergencial, eólica, hidráulica, nuclear, térmica, térmica a gás) - Região - Unidade - Montante
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name.Substring(8);
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 1].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2 + xIndex].Value;
                                            targetWBRowIndex++;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
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
                    }
                    else if (radioButton2.Checked) // Energia armazenada
                    {
                        int GWhOffset;
                        int MWMesOffset;
                        yUpperLimit = 6;

                        foreach (DirectoryInfo currFolder in rootFolder.GetDirectories("*_armazenada"))
                        {
                            xlSourceWorkBook = xlSourceApp.Workbooks.Open(currFolder.FullName + "\\energia_armazenada_mensal.xls");
                            for (iIndex = 1; iIndex <= xlSourceWorkBook.Worksheets.Count; iIndex++)
                            {
                                xlCurrSourceWorkSheet = xlSourceWorkBook.Worksheets[iIndex];
                                if (xlCurrSourceWorkSheet.Cells[1, 1].Value != null)
                                {
                                    year = Convert.ToInt16(xlCurrSourceWorkSheet.Name);

                                    if (year < 2004)
                                    {
                                        GWhOffset = 9;
                                        MWMesOffset = 16;
                                    }
                                    else
                                    {
                                        GWhOffset = 10;
                                        MWMesOffset = 18;
                                    }

                                    currMessage = Convert.ToString(year) + Environment.NewLine + currFolder.Name;
                                    textBox2.Text = currMessage;
                                    textBox2.Refresh();
                                    for (yIndex = 1; yIndex <= yUpperLimit; yIndex++)
                                    {
                                        for (xIndex = 1; xIndex <= 12; xIndex++)
                                        {
                                            // Data - Medida (energia armazenada) - Região - Unidade - Montante
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 1].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[2 + yIndex, 2 + xIndex].Value;
                                            targetWBRowIndex++;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[GWhOffset + yIndex, 1].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[GWhOffset + yIndex, 2].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[GWhOffset + yIndex, 2 + xIndex].Value;
                                            targetWBRowIndex++;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[2, 2 + xIndex].Value) + "/" + Convert.ToString(year);
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = currFolder.Name;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = xlCurrSourceWorkSheet.Cells[MWMesOffset + yIndex, 1].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = xlCurrSourceWorkSheet.Cells[MWMesOffset + yIndex, 2].Value;
                                            xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[MWMesOffset + yIndex, 2 + xIndex].Value;
                                            targetWBRowIndex++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if(radioButton3.Checked) // Histórico de vazões
                    {
                        int monthIndex;
                        int innerUsinaYIndex;
                        string strYear = "";
                        string nomeUsina = "";

                        foreach (DirectoryInfo currFolder in rootFolder.GetDirectories("*Histórico de Vazões*"))
                        {
                            xlSourceWorkBook = xlSourceApp.Workbooks.Open(currFolder.FullName + "\\Vazões_Mensais_1931_2014.xls");
                            xlCurrSourceWorkSheet = xlSourceWorkBook.Worksheets["Tabela"];

                            // yUpperLimit = xlCurrSourceWorkSheet.UsedRange.Rows.Count;
                            yUpperLimit = 17088;
                            yIndex = 1;
                            innerUsinaYIndex = 1;

                            while(yIndex<=yUpperLimit)
                            {
                                nomeUsina = xlCurrSourceWorkSheet.Cells[yIndex, 1].Value;
                                nomeUsina = nomeUsina.Substring(0, nomeUsina.IndexOf("(") - 1);
                                yIndex++; monthIndex = yIndex;
                                yIndex++;
                                while (innerUsinaYIndex <= 84) // 84 anos de histórico
                                {
                                    // Data - Medida ("Vazão") - Região (usina) - Unidade ("m3/s") - Montante
                                    strYear = Convert.ToString(xlCurrSourceWorkSheet.Cells[yIndex, 1].Value);
                                    currMessage = strYear + Environment.NewLine + nomeUsina;
                                    textBox2.Text = currMessage;
                                    textBox2.Refresh();
                                    for (xIndex = 2; xIndex <= 13; xIndex++) // Notar que xIndex começa em 2, e por isso termina em 13
                                    {
                                        xlTargetWorkSheet.Cells[targetWBRowIndex, 1].Value = correctMonthName(xlCurrSourceWorkSheet.Cells[monthIndex, xIndex].Value) + "/" + strYear;
                                        xlTargetWorkSheet.Cells[targetWBRowIndex, 2].Value = "Vazão";
                                        xlTargetWorkSheet.Cells[targetWBRowIndex, 3].Value = nomeUsina;
                                        xlTargetWorkSheet.Cells[targetWBRowIndex, 4].Value = "m3/s";
                                        xlTargetWorkSheet.Cells[targetWBRowIndex, 5].Value = xlCurrSourceWorkSheet.Cells[yIndex, xIndex].Value;
                                        targetWBRowIndex++;
                                    }
                                    yIndex++;
                                    innerUsinaYIndex++;
                                }
                                yIndex += 3; // Pulando MIN, MED e MAX
                                innerUsinaYIndex = 1;
                            }
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

        private string correctMonthName(string inputName)
        {
            string returnValue;

            switch (inputName.ToUpper())
            {
                case "FEV":
                    returnValue = "Feb";
                    break;
                case "ABR":
                    returnValue = "Apr";
                    break;
                case "MAI":
                    returnValue = "May";
                    break;
                case "AGO":
                    returnValue = "Aug";
                    break;
                case "SET":
                    returnValue = "Sep";
                    break;
                case "OUT":
                    returnValue = "Oct";
                    break;
                case "DEZ":
                    returnValue = "Dec";
                    break;
                default:
                    returnValue = inputName;
                    break;
            }

            return (returnValue);
        }
    }
}
