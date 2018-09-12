using OfficeOpenXml;
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

namespace PhoneEditorApp
{
    public partial class Form1 : Form
    {
        string excelad;
        List<string> sütunlar = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Excel Dosyası Yükle";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            excelad = textBox1.Text;
            FileInfo newfile = new FileInfo(textBox1.Text);
            ExcelPackage package = new ExcelPackage(newfile);

            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

            //var rowCnt = worksheet.Dimension.End.Row;  
            var colCnt = worksheet.Dimension.End.Column;

            for (int j = 1; j <= colCnt; j++)
            {
                string deger = worksheet.Cells[1, j].Text;
                sütunlar.Add(deger);
            }

            for (int i = 0; i < colCnt; i++)
            {
                comboBox1.Items.Add(sütunlar[i]);
                comboBox3.Items.Add(sütunlar[i]);
                comboBox4.Items.Add(sütunlar[i]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox6.Text = "İşlem Başladı.";

            List<string> phone_list = new List<string>();
            List<string> alfabe = new List<string> { "a", "b", "c", "ç", "d", "e", "f", "g", "ğ", "h", "ı", "i", "j", "k", "l", "m", "n", "o", "ö", "p", "r", "s", "ş", "t", "u", "ü", "v", "y", "z", "x", "w" };
            int column = 2;
            int sutun = comboBox1.SelectedIndex;
            int il = comboBox3.SelectedIndex;
            int ilçe = comboBox4.SelectedIndex;
            FileInfo newfile = new FileInfo(textBox1.Text);
            ExcelPackage package = new ExcelPackage(newfile);


            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

            var rowCnt = worksheet.Dimension.End.Row;
            var colCnt = worksheet.Dimension.End.Column;
            worksheet.Cells[1, colCnt + 1].Value = "Telefon Edit";

            if (checkBox1.Checked == false && checkBox2.Checked == true && checkBox3.Checked == false)
            {
                for (int i = 2; i <= rowCnt; i++)
                {
                    string phone = worksheet.Cells[i, sutun + 1].Text.ToString().ToUpper().Trim();
                    string kodil = worksheet.Cells[i, il + 1].Text.ToString().ToUpper().Trim();
                    string kodilçe = worksheet.Cells[i, ilçe + 1].Text.ToString().ToUpper().Trim();

                    string alankodu = ilKodu(kodil, kodilçe);

                    int indexOfDahili = phone.IndexOf("DAHİLİ");
                    if (indexOfDahili >= 0)
                        phone = phone.Remove(indexOfDahili);

                    int indexOfDahili2 = phone.IndexOf("DAHILI");
                    if (indexOfDahili2 >= 0)
                        phone = phone.Remove(indexOfDahili2);
                    if (phone.Contains("HAT"))
                    {
                        try
                        {
                            phone = phone.Substring(0, phone.IndexOf("HAT") - 2);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    phone = phone.Replace("/", "").Replace("  ", "").Replace(":", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Trim();

                    for (int j = 0; j < alfabe.Count(); j++)
                    {
                        phone = phone.Replace(alfabe[j], "").Replace(alfabe[j].ToUpper(), "").Trim();
                    }
                    if (phone.Contains("-") || phone.Contains("&")  || phone.Contains("–") || phone.Contains("\"") || phone.Contains("|") || phone.Contains("/") && phone.Length >= 11)
                    {
                        phone = phone.Replace("   ","").Replace("\"","").Replace("&","").Replace("/","").Replace("|","").Replace("–","").Replace("-", "").Replace(" - ", "").Replace("   ", "").Replace("  ", "").Replace(")", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Replace("-", "").Trim();
                        if (phone.Length == 10)
                        {
                            phone = phone.Substring(0, 10);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 3);
                            string phone7 = phone.Substring(16, 2);
                            string phone8 = phone.Substring(18, 2);

                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) == "0")
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);

                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 21)
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string phone5 = phone.Substring(8, 3);
                            string phone6 = phone.Substring(11, 3);
                            string phone7 = phone.Substring(14, 2);
                            string phone8 = phone.Substring(16, 2);

                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 22)
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);
                            string phone8 = phone.Substring(20, 2);

                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if(phone.Length == 18 && phone.Substring(0, 1) == "0")
                        {
                            phone = phone.Substring(1);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 17 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length > 10 && phone.Length < 20)
                        {
                            if (phone.Substring(0, 1) == "0")
                            {
                                phone = phone.Substring(1, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                            else
                            {
                                phone = phone.Substring(0, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                        }
                        else
                        {
                            worksheet.Cells[i, colCnt + 1].Value = "Manuel Bakın.";
                            column++;;
                        }
                    }
                    else if (phone.Length == 13)
                    {
                        string phone1 = phone.Substring(3, 3);
                        string phone2 = phone.Substring(6, 3);
                        string phone3 = phone.Substring(9, 2);
                        string phone4 = phone.Substring(11, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 14)
                    {
                        if (phone.StartsWith("0"))
                            phone = phone.Substring(1);
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 12)
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 11)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 10)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 7)
                    {
                        if (phone.Substring(0, 3) == "444")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = "(0" + alankodu + ") " + phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                    }
                    else if (phone.Length == 20)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 3);
                        string phone7 = phone.Substring(16, 2);
                        string phone8 = phone.Substring(18, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 17)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 2);
                        string phone7 = phone.Substring(15, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + phone5 + " " + phone6 + " " + phone7;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 22)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string phone5 = phone.Substring(12, 3);
                        string phone6 = phone.Substring(15, 3);
                        string phone7 = phone.Substring(18, 2);
                        string phone8 = phone.Substring(20, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 24)//38476
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string phone5 = phone.Substring(14, 3);
                        string phone6 = phone.Substring(17, 3);
                        string phone7 = phone.Substring(20, 2);
                        string phone8 = phone.Substring(22, 2);

                        string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else
                    {
                        worksheet.Cells[i, colCnt + 1].Value = "";
                        column++;
                    }
                }
            }
            else if (checkBox1.Checked == true && checkBox2.Checked == false && checkBox3.Checked == false)
            {
                for (int i = 2; i <= rowCnt; i++)
                {
                    string phone = worksheet.Cells[i, sutun + 1].Text.ToString().ToUpper().Trim();
                    string kodil = worksheet.Cells[i, il + 1].Text.ToString().ToUpper().Trim();
                    string kodilçe = worksheet.Cells[i, ilçe + 1].Text.ToString().ToUpper().Trim();

                    string alankodu = ilKodu(kodil, kodilçe);

                    int indexOfDahili = phone.IndexOf("DAHİLİ");
                    if (indexOfDahili >= 0)
                        phone = phone.Remove(indexOfDahili);

                    int indexOfDahili2 = phone.IndexOf("DAHILI");
                    if (indexOfDahili2 >= 0)
                        phone = phone.Remove(indexOfDahili2);
                    if (phone.Contains("HAT"))
                    {
                        try
                        {
                            phone = phone.Substring(0, phone.IndexOf("HAT") - 2);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    phone = phone.Replace("/", "").Replace("  ", "").Replace(":", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Trim();

                    for (int j = 0; j < alfabe.Count(); j++)
                    {
                        phone = phone.Replace(alfabe[j], "").Replace(alfabe[j].ToUpper(), "").Trim();
                    }
                    if (phone.Contains("-") || phone.Contains("&") || phone.Contains("–") || phone.Contains("\"") || phone.Contains("|") || phone.Contains("/") && phone.Length >= 11)
                    {
                        phone = phone.Replace("   ", "").Replace("\"", "").Replace("&", "").Replace("/", "").Replace("|", "").Replace("–", "").Replace("-", "").Replace(" - ", "").Replace("   ", "").Replace("  ", "").Replace(")", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Replace("-", "").Trim();
                        if (phone.Length == 10)
                        {
                            phone = phone.Substring(0, 10);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 3);
                            string phone7 = phone.Substring(16, 2);
                            string phone8 = phone.Substring(18, 2);

                            string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) == "0")
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);

                            string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 21)
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string phone5 = phone.Substring(8, 3);
                            string phone6 = phone.Substring(11, 3);
                            string phone7 = phone.Substring(14, 2);
                            string phone8 = phone.Substring(16, 2);

                            string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 22)
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);
                            string phone8 = phone.Substring(20, 2);

                            string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 18 && phone.Substring(0, 1) == "0")
                        {
                            phone = phone.Substring(1);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 17 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length > 10 && phone.Length < 20)
                        {
                            if (phone.Substring(0, 1) == "0")
                            {
                                phone = phone.Substring(1, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                            else
                            {
                                phone = phone.Substring(0, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                        }
                        else
                        {
                            worksheet.Cells[i, colCnt + 1].Value = "Manuel Bakın.";
                            column++; ;
                        }
                    }
                    else if (phone.Length == 13)
                    {
                        string phone1 = phone.Substring(3, 3);
                        string phone2 = phone.Substring(6, 3);
                        string phone3 = phone.Substring(9, 2);
                        string phone4 = phone.Substring(11, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 14)
                    {
                        if (phone.StartsWith("0"))
                            phone = phone.Substring(1);
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 12)
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 11)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 10)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 7)
                    {
                        if (phone.Substring(0, 3) == "444")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = "+90 (" + alankodu + ") " + phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                    }
                    else if (phone.Length == 20)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 3);
                        string phone7 = phone.Substring(16, 2);
                        string phone8 = phone.Substring(18, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 17)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 2);
                        string phone7 = phone.Substring(15, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + phone5 + " " + phone6 + " " + phone7;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 22)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string phone5 = phone.Substring(12, 3);
                        string phone6 = phone.Substring(15, 3);
                        string phone7 = phone.Substring(18, 2);
                        string phone8 = phone.Substring(20, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 24)//38476
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string phone5 = phone.Substring(14, 3);
                        string phone6 = phone.Substring(17, 3);
                        string phone7 = phone.Substring(20, 2);
                        string phone8 = phone.Substring(22, 2);

                        string resultphone = "+90 (" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else
                    {
                        worksheet.Cells[i, colCnt + 1].Value = "";
                        column++;
                    }
                }
            }
            else if (checkBox1.Checked == false && checkBox2.Checked == false && checkBox3.Checked == true)
            {
                for (int i = 2; i <= rowCnt; i++)
                {
                    string phone = worksheet.Cells[i, sutun + 1].Text.ToString().ToUpper().Trim();
                    string kodil = worksheet.Cells[i, il + 1].Text.ToString().ToUpper().Trim();
                    string kodilçe = worksheet.Cells[i, ilçe + 1].Text.ToString().ToUpper().Trim();

                    string alankodu = ilKodu(kodil, kodilçe);

                    int indexOfDahili = phone.IndexOf("DAHİLİ");
                    if (indexOfDahili >= 0)
                        phone = phone.Remove(indexOfDahili);

                    int indexOfDahili2 = phone.IndexOf("DAHILI");
                    if (indexOfDahili2 >= 0)
                        phone = phone.Remove(indexOfDahili2);
                    if (phone.Contains("HAT"))
                    {
                        try
                        {
                            phone = phone.Substring(0, phone.IndexOf("HAT") - 2);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    phone = phone.Replace("/", "").Replace("  ", "").Replace(":", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Trim();

                    for (int j = 0; j < alfabe.Count(); j++)
                    {
                        phone = phone.Replace(alfabe[j], "").Replace(alfabe[j].ToUpper(), "").Trim();
                    }
                    if (phone.Contains("-") || phone.Contains("&")  || phone.Contains("–") || phone.Contains("\"") || phone.Contains("|") || phone.Contains("/") && phone.Length >= 11)
                    {
                        phone = phone.Replace("   ", "").Replace("\"","").Replace("&","").Replace("/","").Replace("|","").Replace("–","").Replace("-", "").Replace(" - ", "").Replace("   ", "").Replace("  ", "").Replace(")", "").Replace("-", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("&", "").Replace(",", "").Replace(".", "").Replace("+90", "").Replace("-", "").Trim();
                        if (phone.Length == 10)
                        {
                            phone = phone.Substring(0, 10);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 3);
                            string phone7 = phone.Substring(16, 2);
                            string phone8 = phone.Substring(18, 2);

                            string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 20 && phone.Substring(0, 1) == "0")
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);

                            string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 21)
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string phone5 = phone.Substring(8, 3);
                            string phone6 = phone.Substring(11, 3);
                            string phone7 = phone.Substring(14, 2);
                            string phone8 = phone.Substring(16, 2);

                            string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 22)
                        {
                            string phone1 = phone.Substring(1, 3);
                            string phone2 = phone.Substring(4, 3);
                            string phone3 = phone.Substring(7, 2);
                            string phone4 = phone.Substring(9, 2);

                            string phone5 = phone.Substring(12, 3);
                            string phone6 = phone.Substring(15, 3);
                            string phone7 = phone.Substring(18, 2);
                            string phone8 = phone.Substring(20, 2);

                            string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 18 && phone.Substring(0, 1) == "0")
                        {
                            phone = phone.Substring(1);
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length == 17 && phone.Substring(0, 1) != "0")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 3);
                            string phone3 = phone.Substring(6, 2);
                            string phone4 = phone.Substring(8, 2);

                            string phone5 = phone.Substring(10, 3);
                            string phone6 = phone.Substring(13, 2);
                            string phone7 = phone.Substring(15, 2);
                            string resultphone = "(0" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + "/ " + phone5 + " " + phone6 + " " + phone7;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else if (phone.Length > 10 && phone.Length < 20)
                        {
                            if (phone.Substring(0, 1) == "0")
                            {
                                phone = phone.Substring(1, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                            else
                            {
                                phone = phone.Substring(0, 10);
                                string phone1 = phone.Substring(0, 3);
                                string phone2 = phone.Substring(3, 3);
                                string phone3 = phone.Substring(6, 2);
                                string phone4 = phone.Substring(8, 2);

                                string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                                worksheet.Cells[i, colCnt + 1].Value = resultphone;
                                column++;
                            }
                        }
                        else
                        {
                            worksheet.Cells[i, colCnt + 1].Value = "Manuel Bakın.";
                            column++; ;
                        }
                    }
                    else if (phone.Length == 13)
                    {
                        string phone1 = phone.Substring(3, 3);
                        string phone2 = phone.Substring(6, 3);
                        string phone3 = phone.Substring(9, 2);
                        string phone4 = phone.Substring(11, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 14)
                    {
                        if (phone.StartsWith("0"))
                            phone = phone.Substring(1);
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 12)
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 11)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 10)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 7)
                    {
                        if(phone.Contains("00000") || phone.Contains("11111") || phone.Contains("22222") || phone.Contains("33333") ||
                            phone.Contains("55555") || phone.Contains("66666") || phone.Contains("77777") || phone.Contains("88888") ||
                            phone.Contains("99999"))
                        {
                            worksheet.Cells[i, colCnt + 1].Value = "Geçersiz Numara";
                            column++;
                        }
                        else if (phone.Substring(0, 3) == "444")
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                        else
                        {
                            string phone1 = phone.Substring(0, 3);
                            string phone2 = phone.Substring(3, 2);
                            string phone3 = phone.Substring(5, 2);

                            string resultphone = "(" + alankodu + ") " + phone1 + " " + phone2 + " " + phone3;
                            worksheet.Cells[i, colCnt + 1].Value = resultphone;
                            column++;
                        }
                    }
                    else if (phone.Length == 20)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 3);
                        string phone7 = phone.Substring(16, 2);
                        string phone8 = phone.Substring(18, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 17)
                    {
                        string phone1 = phone.Substring(0, 3);
                        string phone2 = phone.Substring(3, 3);
                        string phone3 = phone.Substring(6, 2);
                        string phone4 = phone.Substring(8, 2);

                        string phone5 = phone.Substring(10, 3);
                        string phone6 = phone.Substring(13, 2);
                        string phone7 = phone.Substring(15, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + phone5 + " " + phone6 + " " + phone7;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 22)
                    {
                        string phone1 = phone.Substring(1, 3);
                        string phone2 = phone.Substring(4, 3);
                        string phone3 = phone.Substring(7, 2);
                        string phone4 = phone.Substring(9, 2);

                        string phone5 = phone.Substring(12, 3);
                        string phone6 = phone.Substring(15, 3);
                        string phone7 = phone.Substring(18, 2);
                        string phone8 = phone.Substring(20, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else if (phone.Length == 24)//38476
                    {
                        string phone1 = phone.Substring(2, 3);
                        string phone2 = phone.Substring(5, 3);
                        string phone3 = phone.Substring(8, 2);
                        string phone4 = phone.Substring(10, 2);

                        string phone5 = phone.Substring(14, 3);
                        string phone6 = phone.Substring(17, 3);
                        string phone7 = phone.Substring(20, 2);
                        string phone8 = phone.Substring(22, 2);

                        string resultphone = "(" + phone1 + ") " + phone2 + " " + phone3 + " " + phone4 + " / " + "(0" + phone5 + ") " + phone6 + " " + phone7 + " " + phone8;
                        worksheet.Cells[i, colCnt + 1].Value = resultphone;
                        column++;
                    }
                    else
                    {
                        worksheet.Cells[i, colCnt + 1].Value = "";
                        column++;
                    }
                }
            }
            else
                MessageBox.Show("Lütfen sadece bir metod seçiniz.");

            textBox1.Text = textBox1.Text.Replace(".xlsx", "");
            FileInfo newfile2 = new FileInfo(textBox1.Text + "_edit" + ".xlsx");
            package.SaveAs(newfile2);
            MessageBox.Show("İşlem Tamamlandı. Çalışmanın olduğu klasöre kaydedildi.");
        }
        public static int sutunBul(string harf)
        {
            if (harf == "A")
                return 1;
            else if (harf == "B")
                return 2;
            else if (harf == "C")
                return 3;
            else if (harf == "D")
                return 4;
            else if (harf == "E")
                return 5;
            else if (harf == "F")
                return 6;
            else if (harf == "G")
                return 7;
            else if (harf == "H")
                return 8;
            else if (harf == "I")
                return 9;
            else if (harf == "J")
                return 10;
            else if (harf == "K")
                return 11;
            else if (harf == "L")
                return 12;
            else if (harf == "M")
                return 13;
            else if (harf == "N")
                return 14;
            else if (harf == "O")
                return 15;
            else if (harf == "P")
                return 16;
            else if (harf == "Q")
                return 17;
            else if (harf == "R")
                return 18;
            else if (harf == "S")
                return 19;
            else if (harf == "T")
                return 20;
            else if (harf == "U")
                return 21;
            else if (harf == "V")
                return 22;
            else if (harf == "W")
                return 23;
            else if (harf == "X")
                return 24;
            else if (harf == "Y")
                return 25;
            else if (harf == "Z")
                return 26;
            else
                return 0;
        }
        public static string ilKodu(string il, string ilce)// Yeni ilçeler buraya yazılacak alan kodu için (212) (216)
        {
            if (il == "İSTANBUL")
            {
                if (ilce == "ARNAVUTKÖY" || ilce == "AVCILAR" || ilce == "BAĞCILAR" ||
                    ilce == "BAHÇELİEVLER" || ilce == "BAKIRKÖY" || ilce == "BAŞAKŞEHİR" ||
                    ilce == "BAYRAMPAŞA" || ilce == "BEŞİKTAŞ" || ilce == "BEYLİKDÜZÜ" ||
                    ilce == "BEYOĞLU" || ilce == "BÜYÜKÇEKMECE" || ilce == "ÇATALCA" ||
                    ilce == "ESENLER" || ilce == "ESENYURT" || ilce == "EYÜPSULTAN" ||
                    ilce == "FATİH" || ilce == "GAZİOSMANPAŞA" || ilce == "GÜNGÖREN" ||
                    ilce == "KAĞITHANE" || ilce == "KÜÇÜKÇEKMECE" || ilce == "SARIYER" ||
                    ilce == "SİLİVRİ" || ilce == "SULTANGAZİ" || ilce == "ŞİŞLİ" ||
                    ilce == "ZEYTİNBURNU" || ilce == "B.ÇEKMECE" || ilce == "B.DÜZÜ" ||
                    ilce == "G.OSMANPAŞA" || ilce == "GOP" || ilce == "K.ÇEKMECE" ||
                    ilce == "EYÜP" || ilce == "Z.BURNU" || ilce == "BEŞİKTAŞ (MERKEZ)" ||
                    ilce == "BEYOĞLU (MERKEZ)" || ilce == "EYÜP (EYÜPSULTAN)" || ilce == "FATİH (MERKEZ, EMİNÖNÜ" ||
                    ilce == "ŞİŞLİ (MERKEZ)")
                {
                    return "212";
                }
                else
                    return "216";
            }
            else if (il == "ADANA")
                return "322";
            else if (il == "ADIYAMAN")
                return "416";
            else if (il == "AFYONKARAHİSAR" || il == "AFYON")
                return "272";
            else if (il == "AĞRI")
                return "472";
            else if (il == "AKSARAY")
                return "382";
            else if (il == "AMASYA")
                return "358";
            else if (il == "ANKARA")
                return "312";
            else if (il == "ANTALYA")
                return "242";
            else if (il == "ARDAHAN")
                return "478";
            else if (il == "ARTVİN")
                return "466";
            else if (il == "AYDIN")
                return "256";
            else if (il == "BALIKESİR")
                return "266";
            else if (il == "BARTIN")
                return "378";
            else if (il == "BATMAN")
                return "488";
            else if (il == "BAYBURT")
                return "458";
            else if (il == "BİLECİK")
                return "228";
            else if (il == "BİNGÖL")
                return "426";
            else if (il == "BİTLİS")
                return "434";
            else if (il == "BOLU")
                return "374";
            else if (il == "BURDUR")
                return "248";
            else if (il == "BURSA")
                return "224";
            else if (il == "ÇANAKKALE" || il == "Ç.KALE")
                return "286";
            else if (il == "ÇANKIRI")
                return "376";
            else if (il == "ÇORUM")
                return "364";
            else if (il == "DENİZLİ")
                return "258";
            else if (il == "DİYARBAKIR")
                return "412";
            else if (il == "DÜZCE")
                return "380";
            else if (il == "EDİRNE")
                return "284";
            else if (il == "ELAZIĞ")
                return "424";
            else if (il == "ERZİNCAN")
                return "446";
            else if (il == "ERZURUM")
                return "442";
            else if (il == "ESKİŞEHİR")
                return "222";
            else if (il == "GAZİANTEP" || il == "ANTEP" || il == "G.ANTEP")
                return "342";
            else if (il == "GİRESUN")
                return "454";
            else if (il == "GÜMÜŞHANE")
                return "456";
            else if (il == "HAKKARİ")
                return "438";
            else if (il == "HATAY")
                return "326";
            else if (il == "IĞDIR")
                return "476";
            else if (il == "ISPARTA")
                return "246";
            else if (il == "İZMİR")
                return "232";
            else if (il == "KAHRAMANMARAŞ" || il == "K.MARAŞ" || il == "MARAŞ")
                return "344";
            else if (il == "KARABÜK")
                return "370";
            else if (il == "KARAMAN")
                return "338";
            else if (il == "KARS")
                return "474";
            else if (il == "KASTAMONU")
                return "366";
            else if (il == "KAYSERİ")
                return "352";
            else if (il == "KIRIKKALE")
                return "318";
            else if (il == "KIRKLARELİ")
                return "288";
            else if (il == "KIRŞEHİR")
                return "386";
            else if (il == "KİLİS")
                return "348";
            else if (il == "KOCAELİ" || il == "İZMİT")
                return "262";
            else if (il == "KONYA")
                return "332";
            else if (il == "KÜTAHYA")
                return "274";
            else if (il == "MALATYA")
                return "422";
            else if (il == "MANİSA")
                return "236";
            else if (il == "MARDİN")
                return "482";
            else if (il == "MERSİN" || il == "İÇEL")
                return "324";
            else if (il == "MUĞLA")
                return "252";
            else if (il == "MUŞ")
                return "436";
            else if (il == "NEVŞEHİR")
                return "384";
            else if (il == "NİĞDE")
                return "388";
            else if (il == "ORDU")
                return "452";
            else if (il == "OSMANİYE")
                return "328";
            else if (il == "RİZE")
                return "464";
            else if (il == "SAKARYA" || il == "ADAPAZARI")
                return "264";
            else if (il == "SAMSUN")
                return "362";
            else if (il == "SİİRT")
                return "484";
            else if (il == "SİNOP")
                return "368";
            else if (il == "SİVAS")
                return "346";
            else if (il == "ŞANLIURFA" || il == "URFA" || il == "Ş.URFA")
                return "414";
            else if (il == "ŞIRNAK")
                return "486";
            else if (il == "TEKİRDAĞ")
                return "282";
            else if (il == "TOKAT")
                return "356";
            else if (il == "TRABZON")
                return "462";
            else if (il == "TUNCELİ")
                return "428";
            else if (il == "UŞAK")
                return "276";
            else if (il == "VAN")
                return "432";
            else if (il == "YALOVA")
                return "226";
            else if (il == "YOZGAT")
                return "354";
            else if (il == "ZONGULDAK")
                return "372";
            else
                return "---";
        }

    }
}
