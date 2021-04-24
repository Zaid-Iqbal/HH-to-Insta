using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
using InstagramApiSharp;
using InstagramApiSharp.API;
using InstagramApiSharp.Classes;
using InstagramApiSharp.API.Builder;
using InstagramApiSharp.Logger;
using InstagramApiSharp.Classes.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using InstaSharp.Models;
using System.Drawing;

namespace HH_to_Insta
{
    class Program
    {
        public static IWorkbook Twb = new XSSFWorkbook(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Twitter.xlsx");
        public static ISheet Tws = Twb.GetSheetAt(0);

        //public static IWebDriver driver = new ChromeDriver(GetMobileView());
        //public static WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));


        public static UserSessionData userSession = getCreds();
        public static IInstaApi api = InstaApiBuilder.CreateBuilder()
                                        .SetUser(userSession)
                                        .UseLogger(new DebugLogger(InstagramApiSharp.Logger.LogLevel.All))
                                        .Build();
        static void Main(string[] args)
        {
            //GoToInsta();
            Post(getNextID());
            //driver.Close();
        }

        public static UserSessionData getCreds()
        {
            String[] lines = System.IO.File.ReadAllLines(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Hardware Hub Access Tokens (Instagram).txt");
            return new UserSessionData
            {
                UserName = lines[0].Substring(lines[0].IndexOf(':') + 2),
                Password = lines[1].Substring(lines[1].IndexOf(':') + 2)
            };                  
        }

        public static string getNextID()
        {
            for (int x = 1; x <= Tws.LastRowNum; x++)
            {
                if (Tws.GetRow(x).Cells[2].ToString() == "No")
                {
                    return Tws.GetRow(x).Cells[1].ToString();
                }
            }
            MessageBox.Show("Excel.getNextID() Error: Unsent ID not found (Check if there are unsent tweets scheduled in Twitter.xlsx)");
            return "Not Found";

        }

        /// <summary>
        /// Post to instagram account
        /// </summary>
        /// <param name="ID"></param>
        public static void Post(string ID)
        {
            if (!Login().Result)
            {
                return;
            }

            System.Drawing.Image img = System.Drawing.Image.FromFile(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
            var mediaImage = new InstaImageUpload
            {
                // leave zero, if you don't know how height and width is it.
                Height = 0,
                Width = 0,
                ImageBytes = ImageToByteArray(img)
            };
            api.MediaProcessor.UploadPhotoAsync(mediaImage, getInstaBody(ID)).Wait();

            #region old Selenium Post approach
            //try
            //{
            //    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'q02Nz _0TPg']")));
            //}
            //catch (WebDriverTimeoutException)
            //{
            //}
            //driver.FindElement(By.XPath("//div [@class = 'q02Nz _0TPg']")).Click();

            //Thread.Sleep(500);
            //System.Windows.Forms.SendKeys.SendWait(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
            //Thread.Sleep(500);
            //System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            //Thread.Sleep(500);
            //System.Windows.Forms.SendKeys.SendWait(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
            //Thread.Sleep(500);
            //System.Windows.Forms.SendKeys.SendWait("{ENTER}");
            //try
            //{
            //    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'mt3GC']")));
            //}
            //catch (WebDriverTimeoutException)
            //{
            //}
            //driver.FindElement(By.XPath("//div [@class = 'mXkkY KDuQp']")).Click();
            //wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'NfvXc']")));
            //driver.FindElement(By.XPath("//div [@class = 'NfvXc']")).Click();
            //driver.FindElement(By.XPath("//textarea")).SendKeys(getInstaBody(ID));
            //driver.FindElement(By.XPath("//div [@class = 'mXkkY KDuQp']")).Click();
            #endregion
        }

        /// <summary>
        /// Get the Body of the text to be added to post
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Check if the creds work
        /// </summary>
        /// <returns></returns>
        public static async Task<bool> Login()
        {
            if (!api.IsUserAuthenticated)
            {
                // login
                Console.WriteLine($"Logging in as {userSession.UserName}");
                var logInResult = await api.LoginAsync();
                if (!logInResult.Succeeded)
                {
                    MessageBox.Show($"Unable to login: {logInResult.Info.Message}");
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Converts Image to Byte Array. Workaround that allows me to upload png's instead of manually converting to jpeg
        /// </summary>
        /// <param name="imageIn"></param>
        /// <returns></returns>
        public static byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        //public static void GoToInsta()
        //{
        //    driver.Navigate().GoToUrl("https://www.instagram.com/");
        //    wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@type='button']")));

        //    driver.FindElement(By.XPath("//button [@type='button']")).Click();
        //    driver.FindElement(By.XPath("//input[@aria-label='Phone number, username, or email']")).SendKeys("hardwarehubdeals@gmail.com");
        //    driver.FindElement(By.XPath("//input[@aria-label='Password']")).SendKeys("exioite33");
        //    driver.FindElement(By.XPath("//div [contains(text(),'Log In')]")).Click();
        //    try
        //    {
        //        wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'eiCW-']")));
        //        MessageBox.Show("Unable to connect to Internet");
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //    }
        //    try
        //    {
        //        wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//button [@class = 'aOOlW   HoLwm ']")))[0].Click();
        //        wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//div [@class = 'cmbtv']")))[0].Click();
        //    }
        //    catch (WebDriverTimeoutException)
        //    {
        //    }
        //}
        //public static ChromeOptions GetMobileView()
        //{
        //    ChromeOptions chromeOptions = new ChromeOptions();
        //    chromeOptions.EnableMobileEmulation("iPhone 6");
        //    return chromeOptions;
        //}
    }




}
