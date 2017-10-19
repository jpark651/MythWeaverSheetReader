using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using System.IO;
using System.Security;
using OfficeOpenXml;

namespace MythWeaverSheetReader
{
    public partial class MainWindow : Window
    {
        List<string> db;
        CookieContainer cookies;
        string path;
        Uri mwUri;

        public MainWindow()
        {
            InitializeComponent();

            path = "\\myth-weaver.db";
            path = Directory.GetCurrentDirectory() + path;
            bool stored = false;
            mwUri = new Uri("http://www.myth-weavers.com/");
            cookies = new CookieContainer();

            if (!File.Exists(path))
            {
                File.WriteAllText(path, "0");
            }
            else
            {
                string[] dbArray = File.ReadLines(path).ToArray();
                db = dbArray.ToList<string>();
                if (db != null && db.Count > 0)
                {
                    if (db[0] != null && db[0].Equals("1"))
                    {
                        stored = true;
                    }
                }
                else
                {
                    File.WriteAllText(path, "0");
                }
            }

            if (stored)
            {
                cookies.SetCookies(mwUri, db[1]);
                FinishSubmit();
            }
        }

        private static SecureString ToSecureString(string input)
        {
            SecureString secure = new SecureString();
            foreach (char c in input)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        private void submit_Click(object sender, RoutedEventArgs e)
        {
            WebClient webClient = new WebClient();
            webClient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            string LOGIN_URL = "http://www.myth-weavers.com/login.php?do=login";
            string postData = "vb_login_username=" + usernameBox.Text + "&vb_login_password=" + passwordBox.Password + "&cookieuser=1&securitytoken=guest&do=login";

            // Post to the login form
            HttpWebRequest webRequest = WebRequest.Create(LOGIN_URL) as HttpWebRequest;
            webRequest.Method = "POST";
            webRequest.ContentType = "application/x-www-form-urlencoded";
            webRequest.CookieContainer = cookies;

            // Writes the form values into the request message
            StreamWriter requestWriter = new StreamWriter(webRequest.GetRequestStream());
            requestWriter.Write(postData);
            requestWriter.Close();
            postData = "";
            passwordBox.Clear();
            usernameBox.Clear();
            webRequest.GetResponse().Close();

            dbUpdate(0, "1");
            string cookieFormat = cookies.GetCookieHeader(mwUri).Replace(';', ',');
            dbUpdate(1, cookieFormat);
            File.WriteAllLines(path, db);
            FinishSubmit();
        }

        private void dbUpdate(int index, string value)
        {
            if (db == null)
            {
                db = new List<string>();
            }
            while (db.Count <= index)
            {
                db.Add("");
            }
            db[index] = value;
        }

        private void FinishSubmit()
        {
            LoggedIn newWindow = new LoggedIn(db, cookies, path);
            newWindow.Show();
            this.Close();
        }
    }
}
