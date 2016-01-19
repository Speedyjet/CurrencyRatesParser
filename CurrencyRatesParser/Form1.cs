using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Fizzler.Systems.HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace CurrencyRatesParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var objExcel = new Application();
            Workbook ObjWorkBook;
            Worksheet ObjWorkSheet;
            ObjWorkBook = objExcel.Workbooks.Add(Missing.Value);
            ObjWorkSheet = (Worksheet) ObjWorkBook.Sheets[1];
            objExcel.Visible = true;
            objExcel.UserControl = true;
            ObjWorkSheet.Cells[2, 1] = "Дата";
            ObjWorkSheet.Cells[2, 2] = "Валюта";
            ObjWorkSheet.Cells[1, 3] = "09:30 - 11.00";
            ObjWorkSheet.Cells[2, 3] = "Покупка";
            ObjWorkSheet.Cells[2, 4] = "Продажа";
            ObjWorkSheet.Cells[1, 5] = "11:00 - 16.00";
            ObjWorkSheet.Cells[2, 5] = "Покупка";
            ObjWorkSheet.Cells[2, 6] = "Продажа";
            ObjWorkSheet.Cells[1, 7] = "С 16:00";
            ObjWorkSheet.Cells[2, 7] = "Покупка";
            ObjWorkSheet.Cells[2, 8] = "Продажа";
            var today = DateTime.Today;
            var firstDay = DateTime.Parse("1.1.2014").Date;
            var j = 2;
            try
            {
                for (var i = firstDay; i <= today; i = i.AddDays(1))
                {
                    var client = new WebClient {Encoding = Encoding.UTF8};

                    var reply =
                        client.DownloadString(
                            "http://ru.kkb.kz/page/RatesConvertingOld?day=" + i.Day + "&month=" + i.Month + "&year=" +
                            i.Year);
                    var document = new HtmlDocument();
                    document.LoadHtml(reply);
                    var nodes = document.DocumentNode.QuerySelectorAll("table.tbl_text2 tr");
                    var table =
                        nodes.Select(x => x.QuerySelectorAll("td").Select(y => y.InnerText).ToArray()).Skip(5).Take(23);
                    //.Skip(1)
                    //.Take(9);
                    var application = new Application
                    {
                        DisplayAlerts = false
                    };
                    
                    foreach (var row in table.Where(row => row[0].StartsWith("1 ")))
                    {
                        j += 1;
                        ObjWorkSheet.Cells[j, 1] = i;
                        ObjWorkSheet.Cells[j, 2] = row[0];
                        ObjWorkSheet.Cells[j, 3] = row[1];
                        ObjWorkSheet.Cells[j, 4] = row[2];
                        ObjWorkSheet.Cells[j, 5] = row[3];
                        ObjWorkSheet.Cells[j, 6] = row[4];
                        ObjWorkSheet.Cells[j, 7] = row[5];
                        ObjWorkSheet.Cells[j, 8] = row[6];
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
