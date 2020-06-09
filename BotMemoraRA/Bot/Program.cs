using Microsoft.Extensions.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Text.RegularExpressions;
using System.Threading;

namespace Bot
{
    class Program
    {
        public static IWebDriver driver;

        static void Main(string[] args)
        {
            var regexRaiz = new Regex(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");

            var diretorioRaiz = regexRaiz.Match(AppDomain.CurrentDomain.BaseDirectory).Value;

            var builder = new ConfigurationBuilder().SetBasePath(diretorioRaiz).AddJsonFile("appsettings.json");

            var configuration = builder.Build();

            driver = new ChromeDriver($"{diretorioRaiz}\\WebDriver");

            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(0));

                driver.Manage().Window.Maximize();

                driver.Navigate().GoToUrl("https://portal.memora.com.br/apex/f?p=1205:LOGIN_DESKTOP");

                var usuario = configuration["memora:usuario"];
                var senha = configuration["memora:senha"];
                var atividade = configuration["memora:atividade"];

                wait.Until(ExpectedConditions.ElementExists(By.Name("P101_USERNAME"))).SendKeys(usuario);
                wait.Until(ExpectedConditions.ElementExists(By.Name("P101_PASSWORD"))).SendKeys(senha);
                wait.Until(ExpectedConditions.ElementExists(By.Id("P101_LOGIN"))).Click();

                Thread.Sleep(TimeSpan.FromSeconds(6));

                if (IsElementPresent(By.Id("P1_GERAR_RA")))
                {
                    wait.Until(ExpectedConditions.ElementExists(By.Id("P1_GERAR_RA"))).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(6));
                }

                wait.Until(ExpectedConditions.ElementExists(By.Id("P1_PREENCHER_RA"))).Click();

                Thread.Sleep(TimeSpan.FromSeconds(6));

                var listaAtividades = driver.FindElements(By.CssSelector("td[headers='DES_ATIVIDADE'] input"));

                foreach (var item in listaAtividades)
                {
                    var valorItem = item.GetAttribute("value");

                    if (string.IsNullOrEmpty(valorItem))
                        item.SendKeys(atividade);
                }

                Thread.Sleep(TimeSpan.FromSeconds(2));
                ((IJavaScriptExecutor)driver).ExecuteScript("apex.submit({request:'Salvar'});");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                //driver.Close();
                //driver.Dispose();
            }
            Console.ReadKey();
        }

        private static bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
    }
}
