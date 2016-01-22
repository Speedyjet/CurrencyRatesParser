using System;
using System.Diagnostics;
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
            ObjWorkSheet.Cells[1, 1] = "Дата";
            ObjWorkSheet.Cells[1, 2] = "Валюта";
            ObjWorkSheet.Cells[1, 3] = "Курс Казкоммерцбанк";
            ObjWorkSheet.Cells[1, 4] = "курс Национальный Банк РК";
            ObjWorkSheet.Cells[1, 5] = "Единицы";
            var today = DateTime.Today;
            var firstDay = DateTime.Parse("12.01.2015").Date;
            var j = 1;
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

                    
                        var nodes = document.DocumentNode.QuerySelectorAll("table.tbl_text2 tr");
                        var table = nodes.Select(x => x.QuerySelectorAll("td").Select(y => y.InnerText).ToArray()).Skip(11);

                        foreach (var row in table)
                        {
                            j += 1;
                            ObjWorkSheet.Cells[j, 1] = i;
                            ObjWorkSheet.Cells[j, 2] = row[0].PadRight(30); //Код валюты
                            ObjWorkSheet.Cells[j, 3] = row[3].Replace("&nbsp;", "").PadLeft(10);//Курс Казкоммерцбанк   //3] = row[1]; //единицы
                            objExcel.Cells[j, 4] = row[6].Replace("&nbsp;", "").PadLeft(10);//row[5]; //курс Национальный Багк РК
                            ObjWorkSheet.Cells[j, 5] = row[4];
                        }
                    }
            }catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            objExcel.Application.Quit();
            Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (Process p2 in ps2)
            {
                p2.Kill();
            }

        }
    }
}
