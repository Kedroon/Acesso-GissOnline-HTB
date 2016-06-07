using System;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Acesso_GissOnline_HTB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string month = DateTime.Now.ToString("MM");
            string year = DateTime.Now.Year.ToString();
            comboBox1.SelectedIndex = int.Parse(month)-1;
            textBox1.Text = year;
            Environment.SetEnvironmentVariable("webdriver.chrome.driver", "chromedriver.exe");

        }

        private void acessar_Click(object sender, EventArgs e)
        {
            int i = 1 + comboBox1.SelectedIndex;
            string month;
            if (i<10)
            {
                month = "0" + i;
            }
            else
            {
                month = i.ToString();
            }
            string year = textBox1.Text;
            Console.WriteLine(month);
            Console.WriteLine(year);
            ChromeDriver js = driver();
            Cookie cookie1 = new Cookie("PID", "2524");
            Cookie cookie2 = new Cookie("MOBI", "560801");
            Cookie cookie3 = new Cookie("TIPO", "0");
            Cookie cookie4 = new Cookie("SUBTIPO", "");
            Cookie cookie5 = new Cookie("mes", month);
            Cookie cookie6 = new Cookie("ano", year);

            js.Navigate().GoToUrl("https://www3.gissonline.com.br/interna/default.cfms");
            js.Manage().Cookies.AddCookie(cookie1);
            js.Manage().Cookies.AddCookie(cookie2);
            js.Manage().Cookies.AddCookie(cookie3);
            js.Manage().Cookies.AddCookie(cookie4);
            js.Manage().Cookies.AddCookie(cookie5);
            js.Manage().Cookies.AddCookie(cookie6);

            js.Navigate().GoToUrl("https://www3.gissonline.com.br/recebidas/listaNotas.cfm?modalidade=T");
        }

        public ChromeDriver driver() {

            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            return new ChromeDriver(chromeDriverService, new ChromeOptions());

        }
    }
}
