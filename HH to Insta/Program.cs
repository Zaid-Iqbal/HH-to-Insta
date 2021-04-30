using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using InstagramApiSharp.API;
using InstagramApiSharp.Classes;
using InstagramApiSharp.API.Builder;
using InstagramApiSharp.Logger;
using InstagramApiSharp.Classes.Models;
using System.Data.SqlClient;

namespace HH_to_Insta
{
    class Program
    {
        public static UserSessionData userSession = getCreds();
        public static IInstaApi api = InstaApiBuilder.CreateBuilder()
                                        .SetUser(userSession)
                                        .UseLogger(new DebugLogger(InstagramApiSharp.Logger.LogLevel.All))
                                        .Build();

        public static String cs = System.IO.File.ReadAllLines(@"C:\Users\email\Desktop\Hardware Hub\SQL Connection String.txt")[0];
        public static SqlConnection con = new SqlConnection(cs);

        static void Main(string[] args)
        {
            //GoToInsta();
            Post(getNextID());
            //driver.Close();
        }

        /// <summary>
        /// Gets auth tokens needed to interact with instagram api
        /// </summary>
        /// <returns></returns>
        public static UserSessionData getCreds()
        {
            String[] lines = System.IO.File.ReadAllLines(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Hardware Hub Access Tokens (Instagram).txt");
            return new UserSessionData
            {
                UserName = lines[0].Substring(lines[0].IndexOf(':') + 2),
                Password = lines[1].Substring(lines[1].IndexOf(':') + 2)
            };                  
        }

        /// <summary>
        /// Gets the ID of the next product that needs to be tweeted. Looks for the product with todays date
        /// </summary>
        /// <returns></returns>
        public static string getNextID()
        {
            con.Open();

            SqlCommand cmd = new SqlCommand("SELECT [ID] FROM Posts WHERE [Date] = CAST(GETDATE() AS DATE)", con);
            SqlDataReader Reader = cmd.ExecuteReader();

            if (Reader.Read())
            {
                String ID = Reader[0].ToString();
                Reader.Close();
                return ID;
            }
            Reader.Close();
            return "None";
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
            SqlCommand cmd = new SqlCommand($"SELECT [Body] FROM Posts WHERE [ID] = '{ID}'", con);
            SqlDataReader Reader = cmd.ExecuteReader();

            if (Reader.Read())
            {
                String str = Reader.GetString(0);
                str = str.Substring(str.IndexOf(':') + 2);
                Reader.Close();
                return str;
            }
            else
            {
                Reader.Close();
                return "None";
            }
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
    }




}
