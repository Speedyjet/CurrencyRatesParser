using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Fizzler.Systems.HtmlAgilityPack;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;


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
            dateTimePicker1.CustomFormat ="d.M.yyyy";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var objExcel = new Microsoft.Office.Interop.Excel.Application();
            objExcel.DisplayAlerts = false;
            var ObjWorkBook = objExcel.Workbooks.Add(Missing.Value);
            var ObjWorkSheet = (Worksheet) ObjWorkBook.Sheets[1];
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
            var firstDay = DateTime.Parse(Convert.ToString(dateTimePicker1.Value, CultureInfo.CurrentCulture)).Date;
            var j = 2;
            var client = new WebClient { Encoding = Encoding.UTF8 };
           
            var document = new HtmlAgilityPack.HtmlDocument();
            try
            {
                for (var i = firstDay; i <= today; i = i.AddDays(1))
                {
                    var reply =
                client.DownloadString(
                    "http://ru.kkb.kz/page/RatesConvertingOld?day=" + i.Day + "&month=" + i.Month + "&year=" +
                    i.Year);
                    document.LoadHtml(reply);

                    if (i<DateTime.Parse("10.1.2015"))
                    {
                        var nodes = document.DocumentNode.QuerySelectorAll("table.tbl_text2 tr");
                        var table =
                            nodes.Select(x => x.QuerySelectorAll("td").Select(y => y.InnerText).ToArray())
                        .Skip(5)
                        .Take(23);

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
                    else
                    {
                        var nodes = document.DocumentNode.QuerySelectorAll("table.tbl_text2 tr");
                        var table = nodes.Select(x => x.QuerySelectorAll("td").Select(y => y.InnerText).ToArray()).Skip(11);

                        foreach (var row in table)
                        {
                            j += 1;
                            ObjWorkSheet.Cells[j, 1] = i;
                            ObjWorkSheet.Cells[j, 2] = row[0].PadRight(30);
                            ObjWorkSheet.Cells[j, 3] = row[1];
                            ObjWorkSheet.Cells[j, 4] = row[2];
                            ObjWorkSheet.Cells[j, 5] = row[3].Replace("&nbsp;", "").PadLeft(10);
                            ObjWorkSheet.Cells[j, 6] = row[4];
                            ObjWorkSheet.Cells[j, 7] = row[5];
                            ObjWorkSheet.Cells[j, 8] = row[6];

                            //Console.WriteLine("{0}: {1}", row[0].PadRight(30), row[3].Replace("&nbsp;", "").PadLeft(10));
                        }
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
