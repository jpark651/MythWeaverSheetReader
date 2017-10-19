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
using System.Windows.Shapes;
using System.Net;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Threading;
using OfficeOpenXml.Style;


namespace MythWeaverSheetReader
{
    public partial class LoggedIn : Window
    {
        string path;
        string comparePath;
        List<string> db;
        List<string> unparsed;
        List<string>[] parsed;
        List<string> oldDiff;
        List<string> newDiff;
        CookieContainer cookies;
        ExcelPackage pck;
        bool updating;
        int oldVer;
        int newVer = 2; // Current db version

        public LoggedIn(List<string> db, CookieContainer cookies, string path)
        {
            InitializeComponent();
            this.db = db;
            this.cookies = cookies;
            this.path = path;
            pck = new ExcelPackage();

            updating = true;
            bool first = true;
            for (int i = 4; i < db.Count; i++)
            {
                if (!first && !db[i].Equals(""))
                {
                    sheetBox.Text += "\n";
                }
                first = false;
                sheetBox.Text += db[i];
            }
            if (!sheetBox.Text.Equals("") && !db[db.Count-1].Equals("\n"))
            {
                sheetBox.Text += "\n";
            }

            if (db.Count > 2)
            {
                outputBox.Text = db[2];
            }
            if (db.Count > 3)
            {
                compareBox.Text = db[3];
            }
            updating = false;
        }

        private void GetSheetData()
        {
            unparsed = new List<string>();
            for (int i = 3; i < db.Count; i++)
            {
                string[] url = new string[1];
                url[0] = db[i];
                string[] split = new string[1];
                split[0] = "id=";
                url = url[0].Split(split, StringSplitOptions.RemoveEmptyEntries);
                if (url.Length > 1)
                {
                    url[0] = "";
                    int result = 0;
                    bool done = false;
                    for (int j = 0; !done && j < url[1].Length; j++)
                    {
                        if (int.TryParse(url[1][j].ToString(), out result))
                        {
                            url[0] += result;
                        }
                        else
                        {
                            done = true;
                        }
                    }
                    string sheetData = Request(url[0]);
                    if (sheetData != null && sheetData.Length > 0)
                    {
                        unparsed.Add(sheetData);
                    }
                }
            }
        }

        private void ParseData()
        {
            int length = 0;
            if (unparsed != null)
            {
                length = unparsed.Count;
            }
            parsed = new List<string>[length];
            for (int i = 0; i < length; i++)
            {
                parsed[i] = new List<string>();
                string parse = unparsed[i];
                var matches = Regex.Matches(parse, "\"[^,](.*?)\":\"(.*?)\"[,}]", RegexOptions.Singleline).GetEnumerator();
                int sheetCount = 0;
                while (matches.MoveNext() && sheetCount <= 1)
                {
                    var current = matches.Current;
                    if (current != null)
                    {
                        string currentString = current.ToString();
                        if (currentString.Contains("sheet_template"))
                        {
                            // Ensures only one sheet of information is read
                            sheetCount++;
                        }
                        else
                        {
                            if (currentString.StartsWith("\"error\":"))
                            {
                                var nameMatch = Regex.Matches(currentString, "\"name\":\"(.*?)\",").GetEnumerator();
                                nameMatch.MoveNext();
                                currentString = nameMatch.Current.ToString();
                            }
                            currentString = currentString.Remove(currentString.Length - 1);
                            parsed[i].Add(currentString);
                        }
                    }
                }
                // Unwanted lines are removed
                // First line must be character name
                parsed[i].RemoveRange(2, 3);
            }
        }

        private void CreateExcel()
        {
            {
                string fileName;
                if (outputBox.Text.Equals(""))
                {
                    fileName = "MW Sheets";
                }
                else
                {
                    fileName = outputBox.Text;
                }
                string currentDir = Directory.GetCurrentDirectory() + "\\";
                string filePath = currentDir + fileName + ".xlsx";
                if (!File.Exists(filePath))
                {
                    FileInfo newFile = new FileInfo(filePath);
                    pck = new ExcelPackage(newFile);
                }
                else
                {
                    bool stop = false;
                    for (int i = 1; !stop; i++)
                    {
                        filePath = currentDir + fileName + " (" + i + ").xlsx";
                        if (!File.Exists(filePath))
                        {
                            stop = true;
                        }
                    }
                    FileInfo newFile = new FileInfo(filePath);
                    pck = new ExcelPackage(newFile);
                }

                // Starts adding sheets
                int length = oldDiff.Count;
                string name = "";
                bool first = true;
                int changes = 0;
                int counter = 2;
                ExcelWorksheet ws = null;
                for (int i = 0; i < length; i++)
                {
                    if (oldDiff[i].Equals("NAME~") && !newDiff[i].Equals("\"\":\"\""))
                    {
                        if (changes == 0 && !first )
                        {
                            pck.Workbook.Worksheets.Delete(ws);
                        }
                        first = false;
                        name = newDiff[i];
                        name = name.Substring(8);
                        name = name.Remove(name.Length - 1);
                        ws = pck.Workbook.Worksheets.Add(name);
                        ws.Cells["B1"].Value = "Old";
                        ws.Cells["D1"].Value = "New";
                        ws.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells["D1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Row(1).Style.Font.Bold = true;
                        ws.Column(1).Style.Font.Bold = true;
                        ws.Column(3).Style.Font.Bold = true;
                        ws.Column(1).Width = 12;
                        ws.Column(2).Width = 40;
                        ws.Column(3).Width = 12;
                        ws.Column(4).Width = 40;
                        changes = 0;
                        counter = 2;
                    }
                    else
                    {
                        string[] split = new string[1];
                        split[0] = "\":\"";
                        string[] oldSplit = oldDiff[i].Split(split, StringSplitOptions.None);
                        string[] newSplit = newDiff[i].Split(split, StringSplitOptions.None);
                        if (oldSplit[0].Length > 1)
                        {
                            oldSplit[0] = oldSplit[0].Substring(1);
                        }
                        if (newSplit[0].Length > 1)
                        {
                            newSplit[0] = newSplit[0].Substring(1);
                        }
                        if (oldSplit.Length > 1)
                        {
                            if (oldSplit[1].Length > 0)
                            {
                                oldSplit[1] = oldSplit[1].Remove(oldSplit[1].Length - 1);
                            }
                        }
                        else
                        {
                            string[] tempSplit = new string[2];
                            oldSplit.CopyTo(tempSplit, 0);
                            oldSplit = tempSplit;
                            oldSplit[1] = "";
                        }
                        if (newSplit.Length > 1)
                        {
                            if (newSplit[1].Length > 0)
                            {
                                newSplit[1] = newSplit[1].Remove(newSplit[1].Length - 1);
                            }
                        }
                        else
                        {
                            string[] tempSplit = new string[2];
                            newSplit.CopyTo(tempSplit, 0);
                            newSplit = tempSplit;
                            newSplit[1] = "";
                        }

                        ws.Cells[counter, 1].Value = oldSplit[0];
                        ws.Cells[counter, 2].Value = oldSplit[1];
                        ws.Cells[counter, 3].Value = newSplit[0];
                        ws.Cells[counter, 4].Value = newSplit[1];
                        changes++;
                        counter++;
                    }
                }
                if (changes == 0)
                {
                    pck.Workbook.Worksheets.Delete(ws);
                }
                if (pck.Workbook.Worksheets.Count > 0)
                {
                    pck.Save();
                }
                else
                {
                    pck.Workbook.Worksheets.Add("No Change");
                    pck.Save();
                }
            }
        }

        private string Request(string id)
        {
            string PAGE_URL = "http://www.myth-weavers.com/api/v1/sheets/sheets/" + id + "/";

            // Sends out cookie along with request for the protected page
            StreamReader responseReader = null;
            string responseData = "";

            HttpWebRequest webRequest = WebRequest.Create(PAGE_URL) as HttpWebRequest;
            webRequest.CookieContainer = cookies;
            webRequest.Accept = "*/*";
            webRequest.Method = "GET";
            webRequest.UserAgent = "Foo";
            responseReader = new StreamReader(webRequest.GetResponse().GetResponseStream());
            // Reads the response
            responseData = responseReader.ReadToEnd();

            if (responseReader != null)
            {
                responseReader.Close();
            }
            return responseData;
        }

        private void sheetBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!updating)
            {
                char[] ret = new char[1];
                ret[0] = '\n';
                string[] update = sheetBox.Text.Split(ret);
                if (update == null || update[0] == null)
                {
                    update = new string[1];
                    update[0] = "";
                }
                for (int i = 0, j = 4; i < update.Length; i++, j++)
                {
                    dbUpdate(j, update[i]);
                }
                int last = update.Length + 4;
                while (last < db.Count) 
                {
                    db.RemoveAt(last);
                }
                File.WriteAllLines(path, db);
            }
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

        private void outputBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!updating)
            {
                string update = outputBox.Text;
                if (update == null)
                {
                    update = "";
                }
                dbUpdate(2, update);
                File.WriteAllLines(path, db);
            }
        }
        
        private void compareBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!updating)
            {
                string update = compareBox.Text;
                if (update == null)
                {
                    update = "";
                }
                dbUpdate(3, update);
                File.WriteAllLines(path, db);
            }
        }

        private void buttonGenerate_LostFocus(object sender, RoutedEventArgs e)
        {
            labelComplete.Visibility = System.Windows.Visibility.Hidden;
        }

        private void labelComplete_MouseDown(object sender, MouseButtonEventArgs e)
        {
            labelComplete.Visibility = System.Windows.Visibility.Hidden;
        }

        private void WriteData()
        {
            if (parsed != null)
            {
                string[] version = new string[2];
                version[0] = "ver";
                version[1] = newVer.ToString();
                File.WriteAllLines(comparePath, version);
                if (parsed.Length > 0)
                {
                    File.AppendAllLines(comparePath, parsed[0]);
                }
                for (int i = 1; i < parsed.Length; i++)
                {
                    // Each array in 'parsed' should contain ~950 strings
                    File.AppendAllLines(comparePath, parsed[i]);
                }
            }
        }

        private void CompareData()
        {
            oldDiff = new List<string>();
            newDiff = new List<string>();
            string fileName;
            if (compareBox.Text.Equals(""))
            {
                fileName = "myth-weaver-compare";
            }
            else
            {
                fileName = compareBox.Text;
            }
            string currentDir = Directory.GetCurrentDirectory() + "\\";
            comparePath = currentDir + fileName + ".txt";
            if (File.Exists(comparePath))
            {
                string[] oldData = File.ReadAllLines(comparePath);
                bool done = false;
                for (int i = 0, j = 0, index = 0, length = parsed[0].Count; (i < oldData.Length || j < parsed[index].Count) && !done; i++, j++)
                {
                    if (i == 0 && oldData[0].Equals("ver"))
                    {
                        // Version handling: oldVer
                        int.TryParse(oldData[1], out oldVer);
                        i = 2;
                    }

                    if (j == length)
                    {
                        j = 0;
                        index++;
                        if (index < parsed.Length)
                        {
                            length = parsed[index].Count;
                        }
                        else
                        {
                            done = true;
                        }
                    }

                    if (!done)
                    {
                        if (i == oldData.Length - 1)
                        {
                            {
                            }
                        }
                        string[] split = new string[1];
                        split[0] = "\":\"";
                        string[] oldSplit = new string[2];
                        string[] newSplit = new string[2];
                        if (i < oldData.Length)
                        {
                            oldSplit = oldData[i].Split(split, StringSplitOptions.None);
                            oldSplit[0] = oldSplit[0].Substring(1);
                        }
                        else
                        {
                            oldSplit[0] = "";
                        }
                        if (j < parsed[index].Count)
                        {
                            newSplit = parsed[index][j].Split(split, StringSplitOptions.None);
                            newSplit[0] = newSplit[0].Substring(1);
                        }
                        else
                        {
                            newSplit[0] = "";
                        }

                        if (!oldSplit[0].Equals(newSplit[0]))
                        {
                            bool skip = false;
                            int check = i + 1;
                            if (check < oldData.Length)
                            {
                                string[] oldCheck = oldData[check].Split(split, StringSplitOptions.None);
                                oldCheck[0] = oldCheck[0].Substring(1);
                                if (oldCheck[0].Equals(newSplit[0]))
                                {
                                    newDiff.Add("");
                                    oldDiff.Add(oldData[i]);
                                    i++;
                                    skip = true;
                                }
                            }
                            check = j + 1;
                            if (!skip && check < parsed[index].Count)
                            {
                                string[] newCheck = parsed[index][check].Split(split, StringSplitOptions.None);
                                newCheck[0] = newCheck[0].Substring(1);
                                if (newCheck[0].Equals(oldSplit[0]))
                                {
                                    if (j == 0)
                                    {
                                        oldDiff.Add("NAME~");
                                    }
                                    else
                                    {
                                        oldDiff.Add("");
                                    }
                                    newDiff.Add(parsed[index][j]);
                                    j++;
                                }
                            }
                        }

                        if (j == 0)
                        {
                            oldDiff.Add("NAME~");
                            newDiff.Add(parsed[index][j]);
                        }
                        if (i < oldData.Length && j < parsed[index].Count && !oldData[i].Equals(parsed[index][j]))
                        {
                            oldDiff.Add(oldData[i]);
                            newDiff.Add(parsed[index][j]);
                        }
                        else if (i < oldData.Length && !(j < parsed[index].Count))
                        {
                            oldDiff.Add(oldData[i]);
                            newDiff.Add("");
                        }
                        else if (!(i < oldData.Length) && j < parsed[index].Count)
                        {
                            oldDiff.Add("");
                            newDiff.Add(parsed[index][j]);
                        }
                    }
                }
            }
        }

        private void buttonGenerate_Click(object sender, RoutedEventArgs e)
        {
            GetSheetData();
            ParseData();
            CompareData();
            if (File.Exists(comparePath))
            {
                CreateExcel();
            }
            WriteData();

            labelComplete.Visibility = System.Windows.Visibility.Visible;
        }
    }
}
