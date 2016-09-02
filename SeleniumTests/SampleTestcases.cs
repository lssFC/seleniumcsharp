using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// Requires reference to WebDriver.Support.dll
using OpenQA.Selenium.Support.UI;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Text.RegularExpressions;

namespace SeleniumTests
{
    [TestFixture("chrome")]
    //[TestFixture("internet explorer")]
    public class SampleTestcases
    {
        private IWebDriver driver;
        private StringBuilder verificationErrors;
        private string baseURL;
        private SqlConnection SQLConnection;
        private string Database;
        private TestUtilities testutil;
        private string browser;

        // Constructor for Test object
        public SampleTestcases(string browser)
        {
            this.browser = browser;
        }

        [OneTimeSetUp]
        public void SetupSampleTests()
        {
            testutil = new TestUtilities(this.browser);
            testutil.SetupTest(this.browser);
            driver = testutil.driver;
            verificationErrors = testutil.verificationErrors;
            baseURL = testutil.baseURL;
            SQLConnection = testutil.SQLConnection;
            Database = testutil.Database;
        }

        [OneTimeTearDown]
        public void TearDownSampleTests()
        {

            testutil.TearDownTest();
        }

        [SetUp]
        public void TestSetup()
        {
            driver.Navigate().GoToUrl(testutil.baseURL + "/SubDomain/Account/Logout");
            Thread.Sleep(3000);

            // Now logout and delete user and group
            IWebElement element = driver.FindElement(By.Name("UserName"));
            element.Clear();
            element.SendKeys("SubDomainadmin");
            element = driver.FindElement(By.Name("Password"));
            element.Clear();
            element.SendKeys("admin");
            element.Submit();
            Thread.Sleep(3000);
            verificationErrors.Clear();
        }


        [TearDown]
        public void TestTearDown()
        {
            Console.WriteLine(verificationErrors.ToString());
            Assert.AreEqual("", verificationErrors.ToString());
        }


        [TestCase]
        public void VerifyDataInDB()
        {

            try
            {
                // Create the query to get all the data for the profile under test from the SQL Database
                string PackagesQuery = "select * from " + (Database + ".dbo.T_Package");

                // Create a SqlDataAdapter to get the results as DataTable
                SqlDataAdapter PackageDataAdapter = new SqlDataAdapter(PackagesQuery, SQLConnection);
                // Create a new DataTable
                DataTable PackagesTable = new DataTable();
                // Fill the DataTable with the result of the SQL statement
                PackageDataAdapter.Fill(PackagesTable);

                //Make sure we have at least one row
                if (PackagesTable.Rows.Count < 4)
                {
                    verificationErrors.Append("Not enough rows in grid in database");
                }

                // Compare the dates on the first two rows. Are they acending or descending?
                bool change_to_descending = false;
                DateTime row0_date = (DateTime)PackagesTable.Rows[0]["P_LastModifiedDate"];
                DateTime row1_date = (DateTime)PackagesTable.Rows[1]["P_LastModifiedDate"];
                if (DateTime.Compare(row0_date, row1_date) <= 0)
                {
                    change_to_descending = true;
                }

                int num_rows = PackagesTable.Rows.Count;
                int day = 0;

                // Change the Modified dates on the packages
                foreach (DataRow drRow in PackagesTable.Rows)
                {
                    // If we are changing the dates to descending order, go back in time with each successive package
                    if (change_to_descending)
                    {
                        drRow["P_LastModifiedDate"] = DateTime.Now.AddDays(-(day++));
                    }
                    else
                    {
                        drRow["P_LastModifiedDate"] = DateTime.Now.AddDays(-(num_rows--));
                    }
                }

                // Update the table with the new dates
                PackageDataAdapter.UpdateCommand = new SqlCommandBuilder(PackageDataAdapter).GetUpdateCommand();
                PackageDataAdapter.Update(PackagesTable);

                PackageDataAdapter.Dispose();

                // Refresh the page so that it reads the database again
                driver.Navigate().Refresh();

                // Get the rows in the table on the page
                IWebElement baseTable = driver.FindElement(By.Id("gridData"));
                IReadOnlyCollection<IWebElement> rows = (baseTable.FindElements(By.TagName("tr")));

                int num_packages_displayed = rows.Count() - 1;

                // Initial time is now
                DateTime last_dt = DateTime.Now;

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
                        // Compare the Modified dates. The displayed plackages should have dates in descending order
                        if (column.GetAttribute("aria-describedby").Equals("gridData_P_LastModifiedDate"))
                        {
                            // Getting just the text will not give values that are not displayed. Hence this technique of getting the title attribute
#if DEBUG
                            Console.WriteLine(column.GetAttribute("title"));
#endif
                            DateTime dt = Convert.ToDateTime(column.GetAttribute("title"));
                            int dt_compare = DateTime.Compare(dt, last_dt);

                            try
                            {
                                Assert.LessOrEqual(dt_compare, 0);
                            }
                            catch (AssertionException e)
                            {
                                Console.WriteLine("Last Date:" + last_dt.ToString() + "Current Date:" + dt.ToString());
                                verificationErrors.Append("Packages are not displayed with the most recent displayed first" + e.ToString());
                            }
                            last_dt = dt;
                        }

                        // In the Signers column make sure the signer names are alphabetically ordered.
                        // To make sure signers are sorted my Last Name and then First Name we will change the names to have 
                        // Last Name first
                        if (column.GetAttribute("aria-describedby").Equals("gridData_Signers"))
                        {
                            string[] signers = column.GetAttribute("title").Split(',');
                            string[] last_name_first_signer = new string[signers.Length];
                            string[] sorted_name_array = new string[signers.Length];
#if DEBUG
                            Console.WriteLine("List before sort");
#endif
                            for (int i = 0; i < signers.Length; i++)
                            {
                                signers[i] = signers[i].TrimStart(' ');
                                string signer = signers[i];
                                // Console.WriteLine(signer);
                                // We don't care how Unknown Signers are ordered. Make it a space 
                                if (signer.Equals("Unknown Signer"))
                                    signer = " ";
                                // Financial Institution is treated as a Last Name
                                if (signer.Equals("Financial Institution"))
                                {
                                    last_name_first_signer[i] = signer;
                                }
                                else
                                {
                                    // Bring Last name to the front of the name
                                    int last_name_index = signer.LastIndexOf(" ");
                                    // One word name
                                    if (last_name_index < 0)
                                    {
                                        last_name_first_signer[i] = signer;
                                    }
                                    else
                                    {
                                        string other_names = signer.Substring(0, last_name_index);
                                        last_name_first_signer[i] = signer.Substring((last_name_index + 1), signer.Length - (last_name_index + 1)) + " " + other_names;
                                    }
                                }

                                sorted_name_array[i] = last_name_first_signer[i];
#if DEBUG
                                Console.WriteLine(last_name_first_signer[i]);
#endif
                            }

                            // Sort the name array. We will compare the order of the diaplayed names to this sorted array
                            Array.Sort(sorted_name_array);
#if DEBUG
                            Console.WriteLine("List after sort");
                            foreach (string signer in sorted_name_array)
                            {
                                Console.WriteLine(signer);
                            }
#endif
                            // Compare to sorted array
                            try
                            {
                                Assert.IsTrue(sorted_name_array.SequenceEqual(last_name_first_signer));
                            }
                            catch (AssertionException e)
                            {
                                verificationErrors.Append("Signer names not in alphabetical order of last name");
                                Console.WriteLine(e.ToString());
                            }
                        }
                    }
                }

                // Create the query to get all the data for the profile under test from the SQL Database
                string SQLStatement = "select * from " + (Database + ".dbo.T_UserPreferenceDefaultConfiguration");

                // +" where EP_ProfileName = '" + ProfileName + "'";

                // Create a SqlDataAdapter to get the results as DataTable
                SqlDataAdapter SQLDataAdapter = new SqlDataAdapter(SQLStatement, SQLConnection);
                // Create a new DataTable
                DataTable dtResult = new DataTable();
                // Fill the DataTable with the result of the SQL statement
                SQLDataAdapter.Fill(dtResult);

                //Make sure we have at least one row
                if (dtResult.Rows.Count != 1)
                {
                    verificationErrors.Append("Did not find rows in grid info in database");
                }

                // Get the user prefernce setting for rows to display on the grid
                int max_rows_in_grid = (int)dtResult.Rows[0]["UPDC_NumberOfRowsPerGrid"];

                SQLDataAdapter.Dispose();

                // Compare to actual rows displayed.
                try
                {
                    Assert.LessOrEqual(num_packages_displayed, max_rows_in_grid);
                }
                catch (AssertionException)
                {
                    verificationErrors.Append("Number of packages displayed exceeds max number set in System settings");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Something went wrong!!!" + e.ToString());
                verificationErrors.Append("Something went wrong!!!" + e.ToString());
            }
        }


       

        [TestCase(TestName ="Demo Testcase Name")]
        [Retry(2)]
        public void TestcaseTwo()
        {

            try
            {
                

            }
            catch (Exception e)
            {
                Console.WriteLine("Something went wrong!!!" + e.ToString());
                verificationErrors.Append("Something went wrong!!!" + e.ToString());
            }
        }
    }
}

