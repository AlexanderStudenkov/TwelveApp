using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TwelveApp
{
    public partial class MainWindow : Window
    {
        private String mainUrl = "https://api.twelvedata.com/";
        private String app_key = "64b2a35e505942929879e90957ee5530";
        private String output = 1.ToString();
        private String symbol = "ETH/USD";
        private String interval = "1min";
        private String format = "JSON";
        private String[] indexes = { "add", "adx", "adxr", "apo", "aroonosc" };

        private bool started = false;

        System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();
        Microsoft.Office.Interop.Excel.Workbook workBook = null;
        Microsoft.Office.Interop.Excel._Worksheet workSheet = null;
        string fileName = "C:\\data\\TwelveData.xls";
        private int v = 0;

        Dictionary<String, float> prev = new Dictionary<String, float>();

        public MainWindow()
        {
            InitializeComponent();

            for(int i=0; i<indexes.Length; i++)
            {
                prev.Add(indexes[i], float.MinValue);
            }

            timer.Tick += new EventHandler(timerTask);
            timer.Interval = new TimeSpan(0, 1, 0);

            Stop.IsEnabled = false;
        }

        
        private void acceptButton_Click(object sender, RoutedEventArgs e)
        {
            Start.IsEnabled = false;
            Stop.IsEnabled = true;

            GetOrCreateFile();

            timer.Start();
        }

        private void GetOrCreateFile()
        {
            var exApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();
                var workBooks = excel.Workbooks;

                if (File.Exists(fileName))
                {
                    workBook = workBooks.Open(fileName);
                    workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;
                }
                else
                {
                    workBook = workBooks.Add();
                    workSheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;
                    
                    workSheet.Cells[1, "A"] = "Code";
                    workSheet.Cells[1, "B"] = "Name";
                    workSheet.Cells[1, "C"] = "Value";
                    workSheet.Cells[1, "D"] = "Percent";
                    workSheet.Cells[1, "E"] = "Datetime";
                }

                v = workSheet.UsedRange.Rows.Count;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.ToString());
            }
        }

        private void timerTask(object sender, EventArgs e)
        {
            for(int i=0; i<indexes.Length; i++)
            {
                GetData(indexes[i]);
            }
        }

        private void GetData(String index)
        {
            string URL = mainUrl + index + "?symbol=" + symbol + "&interval=" + interval + "&outputsize=" + output + "&format=" + format + "&apikey=" + app_key;
            System.Net.WebRequest req = System.Net.WebRequest.Create(URL);
            System.Net.WebResponse resp = req.GetResponse();
            System.IO.Stream stream = resp.GetResponseStream();
            System.IO.StreamReader sr = new System.IO.StreamReader(stream);
            string responce = sr.ReadToEnd();
            sr.Close();

            JObject o = JObject.Parse(responce);

            String code = index;
            String name = o["meta"]["indicator"]["name"].ToString();
            String volume = o["values"][0][index].ToString();

            float newValue = float.Parse(volume, CultureInfo.InvariantCulture.NumberFormat);
            if (prev[index] != float.MinValue)
            {
                float prevValue = prev[index];
                float percent = (newValue - prevValue) / prevValue * 100;
                workSheet.Cells[v + 1, "D"] = percent;
            }
            prev[index] = newValue;

            String datetime = o["values"][0]["datetime"].ToString();

            workSheet.Cells[v + 1, "A"] = code;
            workSheet.Cells[v + 1, "B"] = name;
            workSheet.Cells[v + 1, "C"] = volume;
            
            workSheet.Cells[v + 1, "E"] = datetime;

            v++;

            LastRequest.Content = datetime;
        }

        private void escButton_Click(object sender, RoutedEventArgs e)
        {
            Start.IsEnabled = true;
            Stop.IsEnabled = false;

            timer.Stop();
            if (workBook != null)
            {
                workBook.SaveAs(fileName);
                workBook.Close();
            }
        }

    }
}