using Limilabs.FTP.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ADGTools.App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {

            ISecretInformation secret = new SecretInformation();

            var df = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "DownloadedFiles");

            if (System.IO.Directory.Exists(df)) System.IO.Directory.Delete(df, true);

            System.IO.Directory.CreateDirectory(df);

            var options = new OpenQA.Selenium.Chrome.ChromeOptions();
            options.AddUserProfilePreference("download.default_directory", df);
            options.AddUserProfilePreference("download.prompt_for_download", false);

            var driver = new OpenQA.Selenium.Chrome.ChromeDriver(options);

            driver.Navigate().GoToUrl(@"http://idrottonline.se/AlingsasDGK");

            OpenQA.Selenium.IWebElement elm = null;

            while (elm == null)
            {
                elm = driver.FindElementById("IdrottOnline-LoginBoxTarget");
                System.Threading.Thread.Sleep(1000);
            }

            elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementById("ioui-access-username");
            elm.SendKeys(secret.iotUser);

            elm = driver.FindElementById("ioui-access-password");
            elm.SendKeys(secret.iotPassword);

            elm = driver.FindElementById("ioui-access-login");
            elm = elm.FindElement(OpenQA.Selenium.By.TagName("button"));

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            driver.Navigate().GoToUrl(@"https://ioa.idrottonline.se");

            System.Threading.Thread.Sleep(5000);

            driver.Navigate().GoToUrl(@"https://person.idrottonline.se/search");

            elm = null;

            while (elm == null)
            {
                elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Sök")).FirstOrDefault();
                System.Threading.Thread.Sleep(1000);
            }

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Excel")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            driver.Navigate().GoToUrl(@"https://feeadmin.idrottonline.se/Member/Payments");

            elm = null;

            while (elm == null)
            {
                try
                {
                    elm = driver.FindElementById("paymentsTable").FindElements(OpenQA.Selenium.By.TagName("select")).FirstOrDefault();
                }
                catch { }
                System.Threading.Thread.Sleep(1000);
            }

            elm.SendKeys("Alla");

            System.Threading.Thread.Sleep(5000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Excel, välj")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            driver.Quit();

            var ps = Library.Convert.ExcelToPersons(System.IO.Path.Combine(df, "ExportedPersons.xls"));
            Library.Convert.AddFeesFromExcel(ps, System.IO.Path.Combine(df, "ExportFile.xls"));
            ps = ps.Where(o => o.IsMember).ToList();

            var dta = new DataModel
            {
                date = DateTime.Today.ToString("yyyy-MM-dd"),
                feeYear = DateTime.Today.Year - (DateTime.Today.Month < 5 ? 1 : 0),
                members = ps
            };

            var js = Newtonsoft.Json.JsonConvert.SerializeObject(dta);
            
            textBox1.Text = js;

            using(var ftp = new Ftp())
            {

                ftp.Connect("ftp.alingsasdiscgolf.se");

                ftp.Login(secret.ftpUser, secret.ftpPassword);

                ftp.ChangeFolder("api/members");
                ftp.DeleteFile("memberlist.json");
                ftp.Upload("memberlist.json", System.Text.Encoding.UTF8.GetBytes(js));

                ftp.Close();

            }

        }

    }
}
