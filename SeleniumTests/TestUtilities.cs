using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.Support.UI;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Data.OleDb;
using System.IO;
using Excel;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Globalization;

namespace SeleniumTests
{
    // [TestFixture("firefox")]
    [TestFixture("chrome")]
    // [TestFixture("internet explorer")]
    public class TestUtilities
    {

        public IWebDriver driver;
        public StringBuilder verificationErrors;
        public string baseURL;
        private string browser;

        // SQL Connection String
        public SqlConnection SQLConnection;
        private string ConnectionString;
        public string WebServer;
        public string WebServerUsername;
        public string WebServerPasswd;
        private string DBServer;
        private string DBUsername;
        private string DBPassword;
        public string Database;

        // Constructor for Test object
        public TestUtilities(string browser)
        {
            this.browser = browser;

            string[] lines = System.IO.File.ReadAllLines(@"./ConfigFIle.config");

            baseURL = lines[1].Split('=')[1];
            DBServer = lines[3].Split('=')[1].Trim();
            DBUsername = lines[5].Split('=')[1].Trim();
            DBPassword = lines[7].Split('=')[1].Trim();
            Database = lines[9].Split('=')[1].Trim();

            WebServer = lines[11].Split('=')[1].Trim();
            WebServerUsername = lines[13].Split('=')[1].Trim();
            WebServerPasswd = lines[15].Split('=')[1].Trim();

            //Console.WriteLine("baseURL = " + baseURL);
            //Console.WriteLine("DBServer = " + DBServer);
            //Console.WriteLine("DBUsername = " + DBUsername);
            //Console.WriteLine("DBPassword = " + DBPassword);
            //Console.WriteLine("Database = " + Database);
            //Console.WriteLine("Web Server = " + WebServer);
            //Console.WriteLine("Web Server Username = " + WebServerUsername);
            //Console.WriteLine("Web Server Password = " + WebServerPasswd);
        }

        public static void ClickLinkByHref(IWebDriver driver, String href)
        {

            List<IWebElement> links = new List<IWebElement>();
            links = driver.FindElements(By.TagName("a")).ToList();

            foreach (var link in links)
            {
                // Console.WriteLine("Link" + link.GetAttribute("href"));
                if (link.GetAttribute("href").Equals(href))
                {
                    link.Click();
                    break;
                }
            }
        }

        public static bool IsElementPresent(IWebElement parent, By by)
        {
            try
            {
                parent.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public IWebDriver OpenMobileSite()
        {
            IWebDriver mobile_driver;

            switch (this.browser)
            {
                case "chrome":
                    mobile_driver = new ChromeDriver();
                    break;
                case "internet explorer":
                    mobile_driver = new InternetExplorerDriver();
                    break;
                default:
                    mobile_driver = new FirefoxDriver();
                    break;
            }

            // Open URL
            mobile_driver.Navigate().GoToUrl(baseURL + "/SubDomain");

            return (mobile_driver);
        }

        public void CloseMobileSite(IWebDriver mobile_driver)
        {

            mobile_driver.FindElement(By.CssSelector("li.account-menu > a > span[type=\"button\"] > img[alt=\"arrow\"]")).Click();
            TestUtilities.ClickLinkByHref(mobile_driver, baseURL + "/SubDomainMobile/Account/Logout");

            try
            {
                mobile_driver.Quit();
            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
        }


        public string ie_mod_string(string the_string)
        {
            if (this.browser.Equals("internet explorer"))
            {
                the_string = the_string.Replace("\r\n", " ");
            }
            return (the_string);
        }

        public void SetupTest(string browser)
        {
            switch (this.browser)
            {
                case "chrome":
                    this.driver = new ChromeDriver();
                    break;
                case "internet explorer":
                    this.driver = new InternetExplorerDriver();
                    break;
                default:
                    this.driver = new FirefoxDriver();
                    break;
            }
            verificationErrors = new StringBuilder();

            // Open URL
            // driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(baseURL + "/SubDomainv2");

            // Login
            IWebElement element = driver.FindElement(By.Name("UserName"));
            element.Clear();
            element.SendKeys("SubDomainadmin");
            element = driver.FindElement(By.Name("Password"));
            element.Clear();
            element.SendKeys("admin");
            element.Submit();

            // WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            // wait.Until((d) => { return d.Title.ToLower().StartsWith("SubDomain Workflow"); });

            //#region Building the connection string
            ConnectionString = "Data Source=" + DBServer + ";";
            ConnectionString += "User ID=" + DBUsername + ";";
            ConnectionString += "Password=" + DBPassword + ";";
            ConnectionString += "Initial Catalog=" + Database;
            //#endregion
            SQLConnection = new SqlConnection();

            try
            {
                SQLConnection.ConnectionString = ConnectionString;
                SQLConnection.Open();
                // You can get the server version
                // SQLConnection.ServerVersion
            }
            catch (Exception Ex)
            {
                // Try to close the connection
                if (SQLConnection != null)
                    SQLConnection.Dispose();
                // Create a (useful) error message
                string ErrorMessage = "An error occurred while trying to connect to the server.";
                ErrorMessage += Environment.NewLine;
                ErrorMessage += Environment.NewLine;
                ErrorMessage += Ex.Message;
                // Show error message (this = the parent Form object)
                Console.WriteLine(ErrorMessage + "Connection error.");
                // Stop here
                return;
            }

        }


        public void TearDownTest()
        {

            IWebDriver driver = this.driver;

            driver.Navigate().GoToUrl(baseURL + "/Account/Logout");
            //driver.FindElement(By.CssSelector("li.account-menu > a > span[type=\"button\"] > img[alt=\"arrow\"]")).Click();
            //TestUtilities.ClickLinkByHref(driver, baseURL + "/SubDomainV2/Account/Logout");


            try
            {
                driver.Quit();
                SQLConnection.Close();
                SQLConnection.Dispose();

            }
            catch (Exception)
            {
            }
        }

        public void LogoutUser()
        {
            driver.Navigate().Refresh();
            Thread.Sleep(2000);
            driver.Navigate().GoToUrl(baseURL + "/SubDomainv2/Account/Logout");
            Thread.Sleep(1000);
        }

        public ICollection<IWebElement> FindDocSetting(IWebElement baseTable, string attribute, string attr_value)
        {
            IWebElement search_box = driver.FindElement(By.Id("Name"));
            search_box.SendKeys(attr_value + Keys.Return);
            Thread.Sleep(5000);

            // Get the rows in the table on the page
            baseTable = driver.FindElement(By.Id("gridData"));
            IReadOnlyCollection<IWebElement> rows = (baseTable.FindElements(By.TagName("tr")));

            int num_settings_displayed = rows.Count() - 1;

            ICollection<IWebElement> return_list = new List<IWebElement>();

            foreach (IWebElement row in rows)
            {

                // Ignore the header row
                if (row.GetAttribute("class").Equals("jqgfirstrow"))
                {
                    continue;
                }

                List<IWebElement> columns = row.FindElements(By.TagName("td")).ToList();

                // Traverse each column. We want to do multiple tests as we traverse the rows
                foreach (IWebElement column in columns)
                {
                    if (column.GetAttribute("aria-describedby").Equals(attribute)
                                && column.GetAttribute("title").Equals(attr_value))
                    {
                        return_list.Add(row);
                    }
                }
            }
            return (return_list);
        }

        public ICollection<IWebElement> SelectDocSetting(IWebElement baseTable, string attribute, string attr_value)
        {

            IWebElement search_box = driver.FindElement(By.Id("Name"));
            search_box.SendKeys(attr_value + Keys.Return);
            Thread.Sleep(5000);

            // Get the rows in the table on the page
            baseTable = driver.FindElement(By.Id("gridData"));
            IReadOnlyCollection<IWebElement> rows = (baseTable.FindElements(By.TagName("tr")));

            int num_settings_displayed = rows.Count() - 1;

            ICollection<IWebElement> return_list = new List<IWebElement>();

            foreach (IWebElement row in rows)
            {

                // Ignore the header row
                if (row.GetAttribute("class").Equals("jqgfirstrow"))
                {
                    continue;
                }

                IWebElement checkbox = null;
                bool element_found = false;

                List<IWebElement> columns = row.FindElements(By.TagName("td")).ToList();

                // Traverse each column. We want to do multiple tests as we traverse the rows
                foreach (IWebElement column in columns)
                {
                    if (column.GetAttribute("aria-describedby").Equals("gridData_cb"))
                    {
                        checkbox = column;
                    }

                    if (column.GetAttribute("aria-describedby").Equals(attribute)
                                && column.GetAttribute("title").Equals(attr_value))
                    {

                        element_found = true;

                        return_list.Add(row);
                    }
                }
                if ((checkbox != null) && element_found)
                {
                    checkbox.FindElement(By.XPath(".//input")).Click();
                }
            }

            // driver.FindElement(By.Id("cb_gridData")).Click();
            return (return_list);
        }

        public void DeleteDocSetting(string attribute, string attr_value)
        {
            WebDriverWait wait;
            IWebElement baseTable;

            // Wait for next screen
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until((d) => { return d.FindElement((By.Id("gridData"))); });

            // Get the rows in the table on the page
            baseTable = driver.FindElement(By.Id("gridData"));
            ICollection<IWebElement> selected_doc_settings = SelectDocSetting(baseTable, attribute, attr_value);

            if (selected_doc_settings.Count == 0)
                return;

            // Wait for next screen
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until((d) => { return d.FindElement((By.Id("gridData"))); });

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until((d) => { return d.FindElement((By.Id("aDeleteDocuments"))).Displayed; });

            driver.FindElement(By.Id("aDeleteDocuments")).Click();

            // Need to wait for page to reload after previous doc setting deletion
            Thread.Sleep(1000);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until((d) => { return d.FindElement((By.CssSelector("#divDeleteDocumentDialog > div.er-button-block.er-clear > #btnOk"))); });

            driver.FindElement(By.CssSelector("#divDeleteDocumentDialog > div.er-button-block.er-clear > #btnOk")).Click();
        }

        public void CreateGroup(string group_name, List<string> rights_to_add, List<string> rights_to_remove)
        {
            driver.Navigate().GoToUrl(baseURL + "/SubDomainV2/UserGroup/");
            driver.FindElement(By.XPath("//input[@value='New']")).Click();
            driver.FindElement(By.Id("GroupName")).SendKeys(group_name);

            IWebElement AvailableGroupRightTable = driver.FindElement(By.Id("AvailableGroupRightID"));
            IReadOnlyCollection<IWebElement> AvailableGroupRights = (AvailableGroupRightTable.FindElements(By.TagName("option"))).ToList();
            int num_rights_displayed_table = AvailableGroupRights.Count() - 1;

            try
            {
                Assert.IsFalse((rights_to_add != null) && (rights_to_remove != null));
            }
            catch
            {
                Console.WriteLine("You can only add rights or remove rights. Cannot do both");
                throw;
            }

            foreach (IWebElement AvailableRight in AvailableGroupRights)
            {
                // Check all rights and provide permission
                if (((rights_to_add != null) && rights_to_add.Contains(AvailableRight.Text))
                    || ((rights_to_remove != null) && !rights_to_remove.Contains(AvailableRight.Text)))
                {
                    AvailableRight.Click();
                    driver.FindElement(By.Id("MoveRight")).Click();
                }
            }
            driver.FindElement(By.CssSelector("input[value='Save']")).Click();
        }

        public List<string> allGroupRights = new List<string>() {
                                                "Admin - settings1",
                                                "Admin - settings1",
                                        };

        public IWebElement FindFirstMatchInGrid(string attribute, string attr_value)
        {
            WebDriverWait wait;
            IWebElement baseTable;

            // Wait for next screen
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until((d) => { return d.FindElement((By.Id("gridData"))); });

            baseTable = driver.FindElement((By.Id("gridData")));

            IReadOnlyCollection<IWebElement> rows = (baseTable.FindElements(By.TagName("tr")));

            int num_settings_displayed = rows.Count() - 1;

            bool element_found = false;
            IWebElement returned_row = null;

            foreach (IWebElement row in rows)
            {

                // Ignore the header row
                if (row.GetAttribute("class").Equals("jqgfirstrow"))
                {
                    continue;
                }

                List<IWebElement> columns = row.FindElements(By.TagName("td")).ToList();

                // Traverse each column. We want to do multiple tests as we traverse the rows
                foreach (IWebElement column in columns)
                {
                    if (column.GetAttribute("aria-describedby").Equals(attribute)
                                && column.GetAttribute("title").Equals(attr_value))
                    {
                        element_found = true;
                        returned_row = row;
                        break;
                    }
                }
                if (element_found)
                {
                    break;
                }
            }

            return returned_row;
        }


        public void DeleteGroup(string group_name)
        {
            WebDriverWait wait;

            driver.Navigate().GoToUrl(baseURL + "/SubDomainV2/UserGroup/");

            IWebElement row_match = FindFirstMatchInGrid("gridData_GroupName", group_name);

            if (row_match != null)
            {
                row_match.Click();

                // Wait for next screen
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until((d) => { return d.FindElement((By.XPath("//input[@value='Delete']"))); });

                driver.FindElement(By.XPath("//input[@value='Delete']")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.CssSelector("input[value='Yes']")).Click();
            }
        }


        public void CreateUser(string user_name, string passwd, string group_name)
        {

            driver.Navigate().Refresh();
            // : Go to UserManagement and create user
            driver.Navigate().GoToUrl(baseURL + "/SubDomainV2/UserManagement/");
            driver.FindElement(By.XPath("//input[@value='New']")).Click();
            // : For, we are using the same user for each case if user is created in any of the case then test case fails
            driver.FindElement(By.Id("UserName")).Clear();
            driver.FindElement(By.Id("UserName")).SendKeys(user_name);
            // : For, we are using the same user for each case if user is created in any of the case then test case fails
            // : Now, we are passing the password values
            driver.FindElement(By.Id("Password")).Clear();
            driver.FindElement(By.Id("Password")).SendKeys(passwd);
            driver.FindElement(By.Id("VerifyPassword")).Clear();
            driver.FindElement(By.Id("VerifyPassword")).SendKeys(passwd);
            // Assign group
            // Get all the groups in available groups table
            IWebElement AvailableGroupsTable = driver.FindElement(By.Id("AvailableGroupID"));
            IReadOnlyCollection<IWebElement> AvailableGroup = (AvailableGroupsTable.FindElements(By.TagName("option"))).ToList();
            foreach (IWebElement AssignGroup in AvailableGroup)
            {
                if (AssignGroup.Text.Equals(group_name))
                {
                    AssignGroup.Click();
                    driver.FindElement(By.Id("MoveRight")).Click();
                    break;
                }
            }
            driver.FindElement(By.Id("saveuser")).Click();
        }

        public void DeleteUser(string user_name)
        {
            WebDriverWait wait;

            driver.Navigate().GoToUrl(baseURL + "/SubDomainV2/UserManagement/");

            IWebElement row_match = FindFirstMatchInGrid("gridData_UserName", user_name);

            if (row_match != null)
            {
                row_match.Click();

                // Wait for next screen
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until((d) => { return d.FindElement((By.Id("deleteuser"))); });

                driver.FindElement(By.Id("deleteuser")).Click();

                driver.FindElement(By.CssSelector("input[value='Yes']")).Click();
            }
        }

        public void RestartServices()
        {
            string domain = WebServer.Substring(2);
            string userName = WebServerUsername;
            string password = WebServerPasswd;
            Thread.Sleep(3000);

            driver.Navigate().GoToUrl(baseURL + "/SubDomainV2/Settings/ManageService");

            Thread.Sleep(3000);
            driver.FindElement(By.Id("ServiceDomainName")).Clear();
            driver.FindElement(By.Id("ServiceDomainName")).SendKeys(domain);
            driver.FindElement(By.Id("ServiceUserName")).Clear();
            driver.FindElement(By.Id("ServiceUserName")).SendKeys(userName);
            driver.FindElement(By.Id("ServicePassword")).Clear();
            driver.FindElement(By.Id("ServicePassword")).SendKeys(password);
            driver.FindElement(By.Id("restartService")).Click();
            Thread.Sleep(20000);
        }

        [TestCase]
        public void GetBuildNumber()
        {

            System.IO.StreamWriter file = new System.IO.StreamWriter(@".\build.txt", true);

            SetupTest("chrome");
            string version = driver.FindElement(By.Id("version")).Text;
            string[] words = version.Split('.');
            string build = words.Last();
            file.Write(build);
            file.Flush();
            file.Close();

            try
            {
                driver.Quit();
                SQLConnection.Close();
                SQLConnection.Dispose();

            }
            catch (Exception)
            {
                // Ignore errors if unable to close the browser
            }
            Assert.AreEqual("", verificationErrors.ToString());
            Console.WriteLine(verificationErrors.ToString());
        }

        public void ParseAuditTrail(string path) // string[] words)
        {

            string[] words;

            int page_number = 0;
            //           while (true)
            {
                page_number++;
                words = ExtractTextFromPDF(path, page_number);

                int wi = 0;

                while (!(words[wi].Equals("Package") || words[wi].Equals("Document") || words[wi].Equals("Consent"))) wi++;

                if (words[wi].Equals("Consent"))
                    return;

                if (words[wi].Equals("Package"))
                {
                    FileStream fs = new FileStream(@"C:\Temp\package_info.txt", FileMode.Create);
                    StreamWriter csv = new StreamWriter(fs);

                    string csvData = "";

                    csvData += words[wi++] + words[wi++] + ",";
                    csvData += words[wi++] + words[wi++] + ",";
                    csvData += words[wi++] + ",";
                    csvData += words[wi++] + ",";
                    csvData += words[wi++] + words[wi++] + ",";
                    csvData += words[wi++] + "\n";

                    string current_word = "";
                    string package_number = String.Copy(words[wi]);
                    while (wi < words.Length)
                    {
                        csvData += words[wi] + ",";
                        Console.WriteLine("Package Number: " + words[wi]);
                        wi++;
                        do
                        {
                            current_word = words[wi++];
                            csvData += current_word + " ";
                            Console.WriteLine(current_word);

                        } while (!(words[wi].Equals("Import") || words[wi].Equals("Checked-Out") || words[wi].Equals("Verify")
                                    || words[wi].Equals("Signing")));
                        csvData += ",";
                        current_word = words[wi++];
                        csvData += current_word;
                        Console.WriteLine(current_word);
                        if (!current_word.Equals("Checked-Out"))
                        {
                            csvData += " " + words[wi++];
                        }
                        csvData += ",";
                        for (int x = 0; x < 4; x++)
                        {
                            csvData += words[wi] + " ";
                            Console.WriteLine(words[wi]);
                            wi++;
                        }
                        csvData += ",";
                        if (words[wi].Equals("No"))
                        {
                            csvData += words[wi];
                            Console.WriteLine(words[wi]);
                            wi++;
                        }
                        else
                        {
                            do
                            {
                                csvData += words[wi] + " ";
                                Console.WriteLine(words[wi]);
                                wi++;
                            } while (((wi + 1) < words.Length) && !package_number.Equals(words[wi + 1]));
                        }
                        csvData += ",";
                        csvData += words[wi] + "\n";
                        Console.WriteLine(words[wi]);
                        wi++;

                    }
                    // string output = filepath + filename + ".csv"; // define your own filepath & filename
                    // StreamWriter csv = new StreamWriter(@targetFile, false);
                    csv.Write(csvData);
                    csv.Flush();
                    csv.Close();
                }
            }


        }


        public string[] ExtractTextFromPDF(string path_to_pdf, int page_no)
        {

            using (PdfReader reader = new PdfReader(path_to_pdf))
            {
                PdfReaderContentParser parser = new PdfReaderContentParser(reader);

                SimpleTextExtractionStrategy strategy;

                List<string> word_list = new List<string>();
                string[] words = null;


                //for (int i = 1; i <= 1 /* reader.NumberOfPages */; i++)
                //{
                strategy = parser.ProcessContent(page_no, new SimpleTextExtractionStrategy());
                string[] lines = strategy.GetResultantText().Split('\n');

                foreach (string line in lines)
                {
                    Console.WriteLine("Line: " + line);
                    string[] next_words = line.Split(' ');

                    List<string> next_string = next_words.ToList();
                    word_list.AddRange(next_string);

                }

                words = word_list.ToArray();
                return (words);
                //}

            }
        }



        public static void ConvertExcelToCSV(string sourceFile, string worksheetName, string targetFile)
        {


            FileStream stream = File.Open(sourceFile, FileMode.Open, FileAccess.Read);

            // Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            // Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            // DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();

            // Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();

            if (result == null) throw new ArgumentNullException("result");

            // Console.WriteLine(result.Tables[0].TableName); // to get first sheet name (table name)
            // DataTable dt = result.Tables[0];

            string csvData = "";
            int row_no = 0;

            // Index used to access worksheets. If worksheet is not specified, default is first table 
            int req_ind = 0;

            for (int ind = 0; ind < result.Tables.Count; ind++)
            {
                if (result.Tables[ind].TableName.Equals(worksheetName))
                {
                    req_ind = ind;
                    break;
                }
            }

            while (row_no < result.Tables[req_ind].Rows.Count) // ind is the index of table
            // (sheet name) which you want to convert to csv
            {
                for (int i = 0; i < result.Tables[req_ind].Columns.Count; i++)
                {
                    csvData += result.Tables[req_ind].Rows[row_no][i].ToString();
                    if (i != (result.Tables[req_ind].Columns.Count - 1))
                    {
                        csvData += ",";
                    }
                }
                row_no++;
                csvData += "\n";
            }

            // string output = filepath + filename + ".csv"; // define your own filepath & filename
            StreamWriter csv = new StreamWriter(@targetFile, false);
            csv.Write(csvData);
            csv.Close();
        }
    }
}
