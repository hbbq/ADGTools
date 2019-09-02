using ArxOne.Ftp;
using System;
using System.Linq;
using System.Windows.Forms;

namespace ADGTools.App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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

            //elm.Click();

            System.Threading.Thread.Sleep(1000);

            elm = null;

            while (elm == null)
            {
                try
                {
                    elm = driver.FindElementByName("userName");
                }
                catch
                {
                    // ignored
                }
                System.Threading.Thread.Sleep(1000);
            }

            elm.SendKeys(secret.iotUser);

            elm = driver.FindElementByName("password");
            elm.SendKeys(secret.iotPassword);

            //elm = driver.FindElementById("ioui-access-login");
            elm = driver.FindElement(OpenQA.Selenium.By.TagName("button"));

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

            elm?.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Excel")).FirstOrDefault();

            elm?.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm?.Click();

            System.Threading.Thread.Sleep(5000);

            driver.Navigate().GoToUrl(@"https://feeadmin.idrottonline.se/Member/Payments");

            elm = null;

            while (elm == null)
            {
                try
                {
                    elm = driver.FindElementById("content").FindElements(OpenQA.Selenium.By.TagName("select")).FirstOrDefault(o => o.Location.Y > 0);
                }
                catch
                {
                    // ignored
                }
                System.Threading.Thread.Sleep(1000);
            }

            elm.SendKeys("Alla");

            System.Threading.Thread.Sleep(1000);

            elm = null;

            while (elm == null)
            {
                elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Sök")).FirstOrDefault();
                System.Threading.Thread.Sleep(1000);
            }

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();

            elm?.Click();

            System.Threading.Thread.Sleep(1000);

            elm = driver.FindElementsByTagName("a").Where(o => o.Text.Contains("Excel, välj")).FirstOrDefault();

            elm?.Click();

            System.Threading.Thread.Sleep(1000);

            elm = null;

            while (elm == null)
            {
                elm = driver.FindElementsByTagName("button").Where(o => o.Text.Contains("Exportera")).FirstOrDefault();
                System.Threading.Thread.Sleep(1000);
            }

            elm.Click();

            System.Threading.Thread.Sleep(5000);

            driver.Quit();

            var ps = Library.Convert.ExcelToPersons(System.IO.Path.Combine(df, "ExportedPersons.xlsx"));
            Library.Convert.AddFeesFromExcel(ps, System.IO.Path.Combine(df, "ExportFile.xlsx"));
            ps = ps.Where(o => o.IsMember).ToList();

            var dta = new DataModel<Library.Models.Person>
            {
                date = DateTime.Today.ToString("yyyy-MM-dd"),
                feeYear = DateTime.Today.Year - (DateTime.Today.Month < 5 ? 1 : 0),
                members = ps
            };

            var js = Newtonsoft.Json.JsonConvert.SerializeObject(dta);

            textBox1.Text = js;

            using (var ftp = new FtpClient(new Uri("ftp://ftp.alingsasdiscgolf.se"), new System.Net.NetworkCredential(secret.ftpUser, secret.ftpPassword)))
            using (var str = ftp.Stor("api/members/memberlistfull.json"))
            using (var sw = new System.IO.BinaryWriter(str))
            {
                sw.Write(System.Text.Encoding.UTF8.GetBytes(js));
            }

            var dta2 = new DataModel<Library.Models.Restricted.Person>
            {
                date = DateTime.Today.ToString("yyyy-MM-dd"),
                feeYear = DateTime.Today.Year - (DateTime.Today.Month < 5 ? 1 : 0),
                members = Library.Convert.PersonsToRestrictedPersons(ps).ToList()
            };

            var js2 = Newtonsoft.Json.JsonConvert.SerializeObject(dta2);

            textBox1.Text = js2;

            using (var ftp = new FtpClient(new Uri("ftp://ftp.alingsasdiscgolf.se"), new System.Net.NetworkCredential(secret.ftpUser, secret.ftpPassword)))
            using (var str = ftp.Stor("api/members/memberlist.json"))
            using (var sw = new System.IO.BinaryWriter(str))
            {
                sw.Write(System.Text.Encoding.UTF8.GetBytes(js2));
            }

            var f1 = System.IO.Path.Combine(df, "memberlistfull.json");
            var f2 = System.IO.Path.Combine(df, "memberlist.json");

            System.IO.File.WriteAllText(f1, js);
            System.IO.File.WriteAllText(f2, js2);

        }

    }
}
