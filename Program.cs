using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AIO2
{
    class Program
    {
        static System.Data.DataTable table = new System.Data.DataTable();
        static Dictionary<string, string> hash = new Dictionary<string, string>
                {
                    { "0", "1111011111011101110111110111110" },
                    { "1", "1111111110" },
                    { "2", "101111101111101111110111110" },
                    { "3", "11101111011111011111110" },
                    { "4", "1101111011101111111111010" },
                    { "5", "110111111011110111110" },
                    { "6", "11110111101111011110111110" },
                    { "7", "1011110111111011111010" },
                    { "8", "11101111111101111011111110111110" },
                    { "9", "1110111011101111101111110" },
                    { "/", "11011011011010" }
                };
        static int index = 0;
        static async Task<string> GetSource(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();            
            string data = "";
            //Console.WriteLine("Getting " + url);
            if (response.StatusCode == HttpStatusCode.OK)
            {
                await Task.Run(() => {
                    Stream receiveStream = response.GetResponseStream();
                    StreamReader readStream;

                    if (string.IsNullOrWhiteSpace(response.CharacterSet))
                        readStream = new StreamReader(receiveStream);
                    else
                        readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));

                    data = readStream.ReadToEnd();

                    response.Close();
                    readStream.Close();
                });         
            }
            //Console.WriteLine("Done getting " + url);
            return data;
        }
        static string ToNumber(Bitmap img)
        {
            List<string> numbers = new List<string>();

            int[,] map = new int[1000, 1000];
            int black = 70;
            for (int i = 0; i < img.Width; i++)
            {
                for (int j = 0; j < img.Height; j++)
                {
                    if (img.GetPixel(i, j).R <= black && img.GetPixel(i, j).G <= black && img.GetPixel(i, j).B <= black)
                        map[i, j] = 1;
                }
            }

            string num = "";
            string ans = "";
            bool trigger;
            for (int a = 0; a < img.Width; a++)
            {
                trigger = false;
                for (int b = 0; b < img.Height; b++)
                    if (map[a, b] == 1)
                    {
                        trigger = true;
                        num += "1";
                    }

                if (!trigger && num != "")
                {
                    numbers.Add(num);
                    ans += hash.FirstOrDefault(x => x.Value == num).Key;
                    num = "";
                }
                else if (trigger) num += "0";
            }
            return ans;
        }
        static async Task GetData(string pageUrl, int i)
        {
            var pageDoc = new HtmlDocument();
            pageDoc.LoadHtml(await GetSource(pageUrl));
            if (pageDoc.ParsedText == "Cannot Connect To MySQL Server")
            {
                Console.WriteLine("Please use a vpn or proxy to continue.");
                return;
            }
            //========================================
            //Getting text data==============================
            string dnName = pageDoc.DocumentNode.SelectSingleNode("//span[@title]").InnerText;
            string process = pageDoc.DocumentNode.SelectSingleNode("//div[@class='jumbotron']").InnerText;
            int dateIndex = process.IndexOf("Ngày cấp giấy phép");
            string dnDate = process.Substring(dateIndex + 20, 10);
            int t;
            string dnLaw = "", dnAddress = "";
            try
            {
                t = process.IndexOf("Địa chỉ") + 9;
                while (process[t] != '\r')
                {
                    dnAddress = dnAddress + process[t];
                    t = t + 1;
                }
            }
            catch { }

            try
            {
                t = process.IndexOf("Đại diện pháp luật") + 20;               
                while (process[t] != ':')
                {
                    dnLaw = dnLaw + process[t];
                    t = t + 1;
                }
                dnLaw = dnLaw.Substring(0, dnLaw.IndexOf("Ngày cấp giấy phép") - 36);
            }
            catch
            {
                dnLaw = " ";
            }
            //Done getting text data===============================
            //OCR-ing==========================================
            var images = pageDoc.DocumentNode.SelectNodes("//div[@class='jumbotron']//img");
            string mstBase64 = images[0].Attributes["src"].Value.Substring(22);
            Bitmap pic;
            byte[] bytes = Convert.FromBase64String(mstBase64);
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                pic = (Bitmap)Image.FromStream(ms);
            }
            string dnMst = ToNumber(pic);

            string dnSdt = "";
            if (images.ToArray().Length > 1)
            {
                string sdtBase64 = images[1].Attributes["src"].Value.Substring(22);
                bytes = Convert.FromBase64String(sdtBase64);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    pic = (Bitmap)Image.FromStream(ms);
                }
                dnSdt = ToNumber(pic);
            }
            //Done OCR-ing=========================================
         
            
            table.Rows.Add(i, dnMst, dnSdt, dnDate, dnLaw, dnName, dnAddress);
            Console.WriteLine("Done writing no." + i);
        }
        public static async Task GetPage(string urlAddress)
        {
            List<Task> tasks = new List<Task>();
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(await GetSource(urlAddress));
            if (htmlDoc.ParsedText == "Cannot Connect To MySQL Server")
            { 
                Console.WriteLine("Please use a vpn or proxy to continue.");
            }
            var datas = htmlDoc.DocumentNode.SelectNodes("//div[@class='search-results']");
            foreach (var data in datas)
            {
                string pageUrl = data.ChildNodes["a"].Attributes["href"].Value;
                tasks.Add(GetData(pageUrl, ++index));
            }
      
            await Task.WhenAll(tasks);
        }
        static async Task Run()
        {
            List<Task> tasks = new List<Task>();

            Console.OutputEncoding = Encoding.UTF8;
            Console.Write("Lay tu trang: ");
            int a = int.Parse(Console.ReadLine());
            Console.Write("Den trang: ");
            int page = int.Parse(Console.ReadLine());
            //Init table========================
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Mã số thuế", typeof(string));
            table.Columns.Add("Số điện thoại", typeof(string));
            table.Columns.Add("Ngày Đăng Ký", typeof(string));
            table.Columns.Add("Người đại diện", typeof(string));
            table.Columns.Add("Tên công ty", typeof(string));
            table.Columns.Add("Địa chỉ", typeof(string));

            for (;a<=page;a++)
            {
                string urlAddress = "https://www.thongtincongty.com/thanh-pho-ha-noi/?page=" + a;
                //Writing data to complete.xlsx
                //Mã số thuế;Số điện thoại;Ngày Đăng Ký;Người đại diện;Tên công ty;Địa chỉ
                tasks.Add(GetPage(urlAddress));
                //Console.WriteLine("P" + a);
            }
            await Task.WhenAll(tasks);
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(table, "Complete");
                workbook.SaveAs("Complete.xlsx");
            }
            
            Console.WriteLine("All done UwU");
            Console.ReadKey();
        }
        static async Task Main(string[] args)
        {
            await Run();
        }
    }
}
