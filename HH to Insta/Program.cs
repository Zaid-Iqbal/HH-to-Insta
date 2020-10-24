using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoIt;

namespace HH_to_Insta
{
    class Program
    {
        public static IWebDriver driver = new ChromeDriver(GetMobileView());
        public static WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
        public static IWorkbook Twb = new XSSFWorkbook(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Twitter.xlsx");
        public static ISheet Tws = Twb.GetSheetAt(0);

        static void Main(string[] args)
        {
            String ID = getNextID();
            GoToInsta();
            Post(ID);
            driver.Close();
        }

        public static ChromeOptions GetMobileView()
        {
            ChromeOptions chromeOptions = new ChromeOptions();
            chromeOptions.EnableMobileEmulation("iPhone 6");
            return chromeOptions;
        }

        public static void GoToInsta()
        {
            driver.Navigate().GoToUrl("https://www.instagram.com/");
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@type='button']")));

            driver.FindElement(By.XPath("//button [@type='button']")).Click();
            driver.FindElement(By.XPath("//input[@aria-label='Phone number, username, or email']")).SendKeys("EMAIL");
            driver.FindElement(By.XPath("//input[@aria-label='Password']")).SendKeys("PASSWORD");
            driver.FindElement(By.XPath("//div [contains(text(),'Log In')]")).Click();
            try
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'eiCW-']")));
                MessageBox.Show("Unable to connect to Internet");
            }
            catch (WebDriverTimeoutException)
            {
            }
            try
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'cmbtv']")))[0].Click();
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@class = 'aOOlW   HoLwm ']")))[0].Click();
            }
            catch (WebDriverTimeoutException)
            {
            }           
        }

        public static void Post(string ID)
        {
            try
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'q02Nz _0TPg']")));
            }
            catch (WebDriverTimeoutException)
            {
            }
            driver.FindElement(By.XPath("//div [@class = 'q02Nz _0TPg']")).Click();
            try
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@class = 'aOOlW   HoLwm ']")))[0].Click();
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'q02Nz _0TPg']")));
            }
            catch (WebDriverTimeoutException)
            {
            }
            driver.FindElement(By.XPath("//div [@class = 'q02Nz _0TPg']")).Click();
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            try
            {
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'mt3GC']")));
            }
            catch (WebDriverTimeoutException)
            {
            }
            driver.FindElement(By.XPath("//div [@class = 'mXkkY KDuQp']")).Click();
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'NfvXc']")));
            driver.FindElement(By.XPath("//div [@class = 'NfvXc']")).Click();
            driver.FindElement(By.XPath("//textarea")).SendKeys(getInstaBody(ID));
            driver.FindElement(By.XPath("//div [@class = 'mXkkY KDuQp']")).Click();
        }

        public static string getNextID()
        {
            for (int x = 1; x <= Tws.LastRowNum; x++)
            {
                if (Tws.GetRow(x).Cells[2].ToString() == "No")
                {
                    return Tws.GetRow(x - 1).Cells[1].ToString();
                }
            }
            MessageBox.Show("Excel.getNextID() Error: Unsent ID not found (Check if there are unsent tweets scheduled in Twitter.xlsx)");
            return "Not Found";

        }

        public static string getInstaBody(string ID)
        {
            for (int x = 1; x <= Tws.LastRowNum; x++)
            {
                if (Tws.GetRow(x).Cells[1].ToString() == ID)
                {
                    String send = Tws.GetRow(x).Cells[5].ToString();
                    return send.Substring(send.IndexOf(":") + 2);
                }
            }
            MessageBox.Show("Excel.getNextID() Error: Body not found (Check if there are unsent tweets scheduled in Twitter.xlsx)");
            return "Not Found";
        }
    }
}
