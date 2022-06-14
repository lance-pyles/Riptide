using System;
using System.IO;
using System.Windows.Forms;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using ClosedXML.Excel;
using System.Threading;
using System.Collections;
using System.Threading.Tasks;
using GemBox.Spreadsheet;
using System.Reflection;
using System.Xml.Linq;
using System.Xml.XPath;

using static SeleniumConsoleFramework.KnownFolders;
using SeleniumConsoleFramework;

namespace Riptide
{
    public partial class fRiptide : Form
    {
        static readonly CultureInfo ci = CultureInfo.CurrentCulture;

        public fRiptide()
        {
            InitializeComponent();
        }

        public static async Task UpdateDriver()
        {
            Console.WriteLine("Installing ChromeDriver");       

            await ChromeDriverInstaller.Install();

            Console.WriteLine("ChromeDriver installed");
        }

        public static async Task Rip()
        {
            ChromeOptions opt = new();

            DirectoryInfo diRiptide = new(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Riptide"));
            DirectoryInfo diCSVs = null;
            DirectoryInfo diPDFs = null;
            DirectoryInfo diXLSXs = null;

            Logger logger = new($"{Path.Combine(diRiptide.FullName, "results.txt")}");

            if (diRiptide.Exists)
            {
                diRiptide.Delete(true);
            }

            diRiptide.Create();
            diCSVs = diRiptide.CreateSubdirectory("CSVs");
            diPDFs = diRiptide.CreateSubdirectory("PDFs");
            diXLSXs = diRiptide.CreateSubdirectory("XLSXs");

            logger.Add("Rip started.");

            opt.AddArguments("disable-extensions");
            opt.AddArguments("--start-maximized");

            await UpdateDriver();

            string targetPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string projectDirectory = Directory.GetParent(targetPath).Parent.FullName;
            string configFile = Path.Combine(projectDirectory, "riptide.config");
            
            XElement credentials = XElement.Parse(File.ReadAllText(configFile));
            
            XElement xusername = credentials.XPathSelectElement("//username");
            XElement xpassword = credentials.XPathSelectElement("//password");

            WebDriver driver = new ChromeDriver(targetPath, opt);

            // This will open up the URL
            driver.Navigate().GoToUrl(new Uri("https://members.nwmls.com"));

            #region "Login Page"

            IWebElement username_field = driver.FindElement(By.Id("clareity"));
            IWebElement password_field = driver.FindElement(By.Id("security"));

            string username = xusername.Value;
            string password = xpassword.Value;

            username_field.SendKeys(username);

            password_field.Click();
            password_field.SendKeys(password);

            driver.FindElement(By.Id("loginbtn")).Click();

            #endregion

            IWebElement selectCity = null;
            IWebElement selectCounty = null;

            SelectElement seCity = null;
            SelectElement seCounty = null;

            int iCityCount = 0;

            string sCity = string.Empty;

            driver.Navigate().GoToUrl(new Uri("https://www.matrix.nwmls.com/Matrix/"));

            //click news dialog
            IWebElement weNews = Custom.FindId(logger, driver, "NewsDetailDismissNew");

            if (weNews != null && weNews.Displayed)
            {
                weNews.Click();
            }
            
            Custom.FindText(logger, driver, "Working As").Click();
            Custom.FindPartialLinkText(logger, driver, "Aubie Pouncey").Click();

            NavigateToSearchPage(driver);

            selectCounty = driver.FindElement(By.Id("Fm15_Ctrl36_LB"));
            seCounty = new SelectElement(selectCounty);
            seCounty.SelectByText("Snohomish");

            selectCity = driver.FindElement(By.Id("Fm15_Ctrl31_LB"));
            seCity = new SelectElement(selectCity);

            iCityCount = seCity.Options.Count;

            for (int i = 0; i < iCityCount; i++)
            {
                selectCounty = driver.FindElement(By.Id("Fm15_Ctrl36_LB"));
                seCounty = new SelectElement(selectCounty);
                seCounty.SelectByText("Snohomish");

                selectCity = driver.FindElement(By.Id("Fm15_Ctrl31_LB"));
                seCity = new SelectElement(selectCity);
                seCity.SelectByIndex(i);

                sCity = seCity.Options[i].Text;

                foreach (IWebElement x in driver.FindElements(By.Name("Fm15_Ctrl3412_LB")))
                {
                    //active
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Active")
                    {
                        x.Click();
                    }

                    //contingent
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Contingent")
                    {
                        x.Click();
                    }

                    //pending
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Pending BU Requested")
                    {
                        x.Click();
                    }
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Pending Feasibility")
                    {
                        x.Click();
                    }
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Pending Inspection")
                    {
                        x.Click();
                    }
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Pending Short Sale")
                    {
                        x.Click();
                    }
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Pending")
                    {
                        x.Click();
                    }

                    //sold
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Sold")
                    {
                        x.Click();
                    }

                    //expired
                    if (x.GetAttribute("data-mtx-track") == "Multi Status - Expired")
                    {
                        x.Click();
                    }
                }

                driver.FindElement(By.Id("m_ucSearchButtons_m_lbSearch")).Click();

                #region "Result Page"

                IWebElement weSearchSummary = Custom.FindId(logger, driver, "m_tblSqlSearchInfo");
                IWebElement weNothingFound = Custom.FindId(logger, driver, "m_lblNoResultsText");

                // Create a FileInfo  
                FileInfo fi_download = null;
                FileInfo fi_city = new($"{Path.Combine(diCSVs.FullName, sCity)}.csv");

                int iSummaryCount = 0;

                string sSummary = string.Empty;
                string sCheckedCount = string.Empty;
                string sDateSearch = DateTime.Now.Date.ToShortDateString();
                string sDateEndSearch = DateTime.Now.AddDays(-180).Date.ToShortDateString();
                string sSearch = $@"Status is one of 'Active', 'Contingent'
Status is 'Pending BU Requested'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Pending Feasibility'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Pending Inspection'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Pending Short Sale'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Pending'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Sold'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Status is 'Expired'
Contractual Date is {sDateSearch} to {sDateEndSearch}
Sale Type is 'MLS'
State Or Province is 'Washington'
County is 'Snohomish'
City is '{sCity}'
Status is not 'Incomplete'
Ordered by Status, Area, Current Price";
                string sSearchActual = weSearchSummary.Text;

                if (sSearchActual.Contains(sSearch) == false)
                {
                    MessageBox.Show($"Search setup incorrectly for {sCity}.");
                }

                if (weNothingFound == null)
                {
                    sSummary = driver.FindElement(By.Id("m_lblPagingSummary")).Text;
                    sCheckedCount = driver.FindElement(By.Id("m_lblCheckedCount")).Text;

                    sSummary = sSummary.Substring(sSummary.IndexOf("of ", StringComparison.OrdinalIgnoreCase), sSummary.Length - sSummary.IndexOf("of ", StringComparison.OrdinalIgnoreCase)).Substring(3);

                    if (sSummary == "5000+")
                    {
                        iSummaryCount = -1;
                    }
                    else
                    {
                        iSummaryCount = Convert.ToInt32(sSummary, ci);
                    }
                }

                if (iSummaryCount == 0)
                {
                    // Log
                    logger.Add($"No result(s) found in {sCity}.  Results could not be exported. Time: {DateTime.UtcNow}.{Environment.NewLine}");

                    NavigateToSearchPage(driver);

                    continue;
                }

                if (iSummaryCount > 2000)
                {
                    MessageBox.Show($"Too many results found in {sCity}.");
                    NavigateToSearchPage(driver);

                    continue;
                }

                //You have too many items checked for this action (max = 2000)
                //Too many items to check (max 5000)

                // Log
                logger.Add($"{iSummaryCount} result(s) found in {sCity}. Time: {DateTime.UtcNow}.{Environment.NewLine}");

                //select All
                Custom.FindId(logger, driver, "m_lnkCheckAllLink").Click();
                sCheckedCount = Custom.FindId(logger, driver, "m_lblCheckedCount").Text.Replace("Checked ", "");
                int iCheckedCount = Convert.ToInt32(sCheckedCount, ci);
                if (iCheckedCount != iSummaryCount)
                {
                    MessageBox.Show("Not all items checked.");
                }

                //click Actions
                Custom.FindId(logger, driver, "m_liActionsTab").Click();

                //Export
                Custom.FindId(logger, driver, "m_lbExport").Click();

                Custom.FindId(logger, driver, "m_btnExport").Click();

                #endregion

                // Source file to be renamed  
                string sourceFile = Path.Combine(KnownFolders.GetPath(KnownFolder.Downloads), "Rental(testa) - Market Snapshot -wBuilder.csv");

                //download delay
                for (int x = 0; x < 20; x++)
                {
                    fi_download = new FileInfo(sourceFile);

                    if (fi_download.Exists)
                    {
                        // Move file with a new name. Hence renamed.  
                        break;
                    }

                    Thread.Sleep(2000);
                };

                // Move file with a new name. Hence renamed.  
                if (fi_city.Exists)
                {
                    fi_city.Delete();
                }

                fi_download.MoveTo(Path.Combine(diCSVs.FullName, fi_city.Name));

                if (!fi_download.Exists)
                {
                    MessageBox.Show($"Download failed for {sCity}.");

                    // Log
                    logger.Add($"Download failed for {sCity}.");

                }

                NavigateToSearchPage(driver);

            }

            driver.Close();
            driver.Quit();

            Parallel.ForEach(diCSVs.GetFiles("*.csv"), i =>
            {
                AnalyzeData(new FileInfo(i.FullName), logger, diXLSXs, diPDFs);
            });

            logger.Add("Rip Completed.");

            MessageBox.Show("Run completed.");
        }

        private void bRun_Click(object sender, EventArgs e)
        {
            Rip();
        }
                
        private static void AnalyzeData(FileInfo fi, Logger logger, DirectoryInfo di_XLSXs, DirectoryInfo di_PDFs)
        {
            using (DataTable dtHomes = Custom.ConvertCSVtoDataTable(fi.FullName))
            {
                DataTable dtCloned = dtHomes.Clone();

                List<(decimal startRange, decimal endRange)> lPriceRanges = new()
                {
                    (0, 249999),
                    (250000, 299999),
                    (300000, 349999),
                    (350000, 399999),
                    (400000, 449999),
                    (450000, 499999),
                    (500000, 549999),
                    (550000, 599999),
                    (600000, 649999),
                    (650000, 699999),
                    (700000, 749999),
                    (750000, 799999),
                    (800000, 849999),
                    (850000, 899999),
                    (900000, 949999),
                    (950000, 999999),
                    (1000000, 1149000),
                    (1150000, 1299000),
                    (1300000, -1)
                };

                CleanupColum(dtHomes, dtCloned, "Current Price", typeof(decimal));
                CleanupColum(dtHomes, dtCloned, "Listing Number", typeof(int));
                CleanupColum(dtHomes, dtCloned, "Status", typeof(string));
                CleanupColum(dtHomes, dtCloned, "Status Change Date", typeof(DateTime));
                CleanupColum(dtHomes, dtCloned, "Original Price", typeof(decimal));
                CleanupColum(dtHomes, dtCloned, "Selling Price", typeof(decimal));
                CleanupColum(dtHomes, dtCloned, "Listing Price", typeof(decimal));
                CleanupColum(dtHomes, dtCloned, "Listing Date", typeof(DateTime));

                foreach ((decimal startRange, decimal endRange) in lPriceRanges)
                {
                    foreach (DataRow row in dtHomes.Rows)
                    {
                        string cp = row[dtCloned.Columns.IndexOf("CurrentPrice")].ToString().Replace("\"", "");
                        row[dtCloned.Columns.IndexOf("CurrentPrice")] = DBNull.Value;
                        if (!String.IsNullOrEmpty(cp))
                        {
                            row[dtCloned.Columns.IndexOf("CurrentPrice")] = Convert.ToDecimal(cp, ci);
                        }

                        row[dtCloned.Columns.IndexOf("ListingNumber")] = Convert.ToInt32(row[dtCloned.Columns.IndexOf("ListingNumber")].ToString().Replace("\"", ""), ci);
                        row[dtCloned.Columns.IndexOf("Status")] = Convert.ToString(row[dtCloned.Columns.IndexOf("Status")].ToString().Replace("\"", ""), ci);
                        row[dtCloned.Columns.IndexOf("StatusChangeDate")] = Convert.ToString(row[dtCloned.Columns.IndexOf("StatusChangeDate")].ToString().Replace("\"", ""), ci);

                        string op = row[dtCloned.Columns.IndexOf("OriginalPrice")].ToString().Replace("\"", "");
                        row[dtCloned.Columns.IndexOf("OriginalPrice")] = DBNull.Value;
                        if (!String.IsNullOrEmpty(op))
                        {
                            row[dtCloned.Columns.IndexOf("OriginalPrice")] = Convert.ToDecimal(op, ci);
                        }

                        string sp = row[dtCloned.Columns.IndexOf("SellingPrice")].ToString().Replace("\"", "");
                        row[dtCloned.Columns.IndexOf("SellingPrice")] = DBNull.Value;
                        if (!String.IsNullOrEmpty(sp))
                        {
                            row[dtCloned.Columns.IndexOf("SellingPrice")] = Convert.ToDecimal(sp, ci);
                        }

                        string lp = row[dtCloned.Columns.IndexOf("ListingPrice")].ToString().Replace("\"", "");
                        row[dtCloned.Columns.IndexOf("ListingPrice")] = DBNull.Value;
                        if (!String.IsNullOrEmpty(lp))
                        {
                            row[dtCloned.Columns.IndexOf("ListingPrice")] = Convert.ToDecimal(lp, ci);
                        }

                        row[dtCloned.Columns.IndexOf("ListingDate")] = Convert.ToDateTime(row[dtCloned.Columns.IndexOf("ListingDate")].ToString().Replace("\"", ""), ci);

                        dtCloned.ImportRow(row);
                    }
                                        
                    using (DataTable results = new(Path.GetFileNameWithoutExtension(fi.FullName)))
                    {
                            results.Columns.Add("Price Range", typeof(string));
                            results.Columns.Add("Active Listings", typeof(int));
                            results.Columns.Add("Pending Listings", typeof(int));
                            results.Columns.Add("Pending Ratio", typeof(decimal));
                            results.Columns["Pending Ratio"].ExtendedProperties.Add("Format", "P");
                            
                            results.Columns.Add("Months Of Inventory", typeof(decimal));
                            results.Columns.Add("Expired Listings", typeof(int));
                            results.Columns.Add("Closed Listings", typeof(int));
                            results.Columns.Add("Avg Orignal List Price SOLD", typeof(int));
                            results.Columns["Avg Orignal List Price SOLD"].ExtendedProperties.Add("Format", "C0");

                            results.Columns.Add("Avg Final List Price SOLD", typeof(int));
                            results.Columns["Avg Final List Price SOLD"].ExtendedProperties.Add("Format", "C0");

                            results.Columns.Add("Avg Sale Price SOLD", typeof(int));
                            results.Columns["Avg Sale Price SOLD"].ExtendedProperties.Add("Format", "C0");

                            results.Columns.Add("List To Sales Ratio", typeof(decimal));
                            results.Columns["List To Sales Ratio"].ExtendedProperties.Add("Format", "P");

                            results.Columns.Add("Average Days On Market SOLD", typeof(int));
                            results.Columns.Add("Average Days On Market ACTIVE", typeof(int));

                            foreach ((int startRange, int endRange) x in lPriceRanges)
                            {
                                DataRow row = results.NewRow();

                                if (x.endRange == -1)
                                {
                                    row["Price Range"] = $"{x.startRange:C0}+";
                                }
                                else
                                {
                                    row["Price Range"] = $"{x.startRange:C0} - {x.endRange:C0}";
                                }

                                row["Active Listings"] = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Active")
                                .Select(_ => new { key1 = _.Field<int>("ListingNumber") }).Distinct().Count();

                                row["Pending Listings"] = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Pending")
                                .Select(_ => new { key1 = _.Field<int>("ListingNumber") }).Distinct().Count();

                                if (Convert.ToInt16(row["Active Listings"], ci) == 0)
                                {
                                    row["Pending Ratio"] = 0;
                                }
                                else
                                {
                                    row["Pending Ratio"] = (Convert.ToInt16(row["Pending Listings"], ci) / Convert.ToInt16(row["Active Listings"], ci));
                                }

                                row["Months Of Inventory"] = "0";

                                row["Expired Listings"] = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Expired" && _.Field<DateTime>("StatusChangeDate").AddMonths(6).Date >= DateTime.Now.Date)
                                .Select(_ => new { key1 = _.Field<int>("ListingNumber") }).Distinct().Count();

                                row["Closed Listings"] = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Sold" && _.Field<DateTime>("StatusChangeDate").AddMonths(6).Date >= DateTime.Now.Date)
                                .Select(_ => new { key1 = _.Field<int>("ListingNumber") }).Distinct().Count();

                                var t = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Sold");

                                if (!t.Any())
                                {
                                    row["Avg Orignal List Price SOLD"] = DBNull.Value;
                                }
                                else
                                {
                                    row["Avg Orignal List Price SOLD"] = Convert.ToInt32(t.Average(_ => _.Field<decimal?>("OriginalPrice")));
                                }

                                var s = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Sold");

                                if (!s.Any())
                                {
                                    row["Avg Final List Price SOLD"] = DBNull.Value;
                                }
                                else
                                {
                                    row["Avg Final List Price SOLD"] = Convert.ToInt32(s.Average(_ => _.Field<decimal?>("ListingPrice")));
                                }

                                var r = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Sold");

                                if (!r.Any())
                                {
                                    row["Avg Sale Price SOLD"] = DBNull.Value;
                                }
                                else
                                {
                                    row["Avg Sale Price SOLD"] = Convert.ToInt32(r.Average(_ => _.Field<decimal>("SellingPrice")));
                                }

                                if (row["Avg Sale Price SOLD"] == DBNull.Value)
                                {
                                    row["List To Sales Ratio"] = DBNull.Value;
                                }
                                else
                                {
                                    if (Convert.ToInt32(row["Avg Sale Price SOLD"], ci) == 0)
                                    {
                                        row["List To Sales Ratio"] = 0.ToString("P", ci);
                                    }
                                    else
                                    {
                                        row["List To Sales Ratio"] = Decimal.Round(Convert.ToDecimal((Convert.ToDecimal(row["Avg Sale Price SOLD"], ci) / Convert.ToDecimal(row["Avg Final List Price SOLD"], ci))), 0);                                        
                                    }
                                }

                                var q = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Sold");

                                if (!q.Any())
                                {
                                    row["Average Days On Market SOLD"] = DBNull.Value;
                                }
                                else
                                {
                                    row["Average Days On Market SOLD"] = Decimal.ToInt32(Math.Truncate(Convert.ToDecimal(q.Average(_ => DateTime.Now.Date.Subtract(_.Field<DateTime>("ListingDate")).Days))));                                    
                                }

                                var p = dtCloned.AsEnumerable()
                                .Where(_ => _.Field<decimal?>("CurrentPrice") >= x.startRange && (_.Field<decimal?>("CurrentPrice") <= x.endRange || _.Field<decimal?>("CurrentPrice") == -1) && _.Field<string>("Status") == "Active");

                                if (!p.Any())
                                {
                                    row["Average Days On Market ACTIVE"] = DBNull.Value;
                                }
                                else
                                {
                                    row["Average Days On Market ACTIVE"] = Decimal.ToInt32(Math.Truncate(Convert.ToDecimal(p.Average(_ => DateTime.Now.Date.Subtract(_.Field<DateTime>("ListingDate")).Days))));                                    
                                }

                                //Active listings in 6 months / Sold listings in 6 months
                                if (Convert.ToInt16(row["Closed Listings"], ci) != 0)
                                {
                                    row["Months Of Inventory"] = Decimal.Round(Convert.ToDecimal(row["Active Listings"], ci) / Convert.ToDecimal(row["Closed Listings"], ci), 2);
                                }

                                results.Rows.Add(row);

                            };

                            Custom.WriteDataToExcel(results, di_XLSXs, di_PDFs, true);
                        }
                    
                    // Log
                    logger.Add($"Excel file created for {Path.GetFileNameWithoutExtension(fi.FullName)}.");
                }
            }
        }

        private static void NavigateToSearchPage(WebDriver driver)
        {
            driver.Url = "https://www.matrix.nwmls.com/Matrix/Search/SingleFamily/Quick";

            driver.Navigate();
        }

        private static void CleanupColum(DataTable dtOriginal, DataTable dtClone, string name, Type t)
        {
            dtOriginal.Columns[dtClone.Columns.IndexOf("\"" + name + "\"")].ColumnName = name.Replace(" ", "");
            dtClone.Columns[dtClone.Columns.IndexOf("\"" + name + "\"")].ColumnName = name.Replace(" ", "");
            dtClone.Columns[dtClone.Columns.IndexOf(name.Replace(" ", ""))].DataType = t;
        }
    }

    public class Logger
    {
        static readonly CultureInfo ci = CultureInfo.CurrentCulture;

        private readonly Queue logs = new();

        private readonly string path = string.Empty;

        public Logger(string path)
        {
            this.path = path;
        }

        public void Add(string t)
        {
            this.logs.Enqueue("[" + CurrentTime() + "] " + t);
            this.SaveNow();
        }

        private void SaveNow()
        {
            if (this.logs.Count > 0)
            {
                // Get from queue
                string err = Convert.ToString(this.logs.Dequeue(), ci);
                // Save to logs
                SaveToFile(err, this.path);
            }
        }

        public bool SaveToFile(string text, string path)
        {
            try
            {
                // string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                // text = text + Environment.NewLine;
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(text);
                    sw.Close();
                }
            }
            catch
            {
                // return to queue
                this.logs.Enqueue(text + "[SAVE_ERROR]");
                return false;
            }
            return true;
        }

        public static string CurrentTime()
        {
            DateTime d = DateTime.UtcNow.ToLocalTime();
            return d.ToString("yyyy-MM-dd hh:mm:ss", ci);
        }
    }

    public static class Custom
    {
        static readonly CultureInfo ci = CultureInfo.CurrentCulture;

        public static IWebElement FindId(Logger logger, WebDriver driver, string search)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }

            if (driver == null)
            {
                throw new ArgumentNullException(nameof(driver));
            }

            int attempts = 0;

            do
            {
                try
                {
                    attempts += 1;

                    return driver.FindElement(By.Id(search));
                }
                catch (ElementNotSelectableException)
                {
                    Thread.Sleep(100);
                }
                catch (NoSuchElementException)
                {
                    Thread.Sleep(100);
                }
                finally
                {

                }
            } while (attempts < 30);

            // Log
            logger.Add($"Unable to find by id of '{search}'.");

            return null;
        }

        public static IWebElement FindText(Logger logger, WebDriver driver, string search)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }

            if (driver == null)
            {
                throw new ArgumentNullException(nameof(driver));
            }

            foreach (IWebElement element in driver.FindElements(By.XPath("//*")))
            {
                if (element.Text == search)
                {
                    return element;
                }
            }

            // Log
            logger.Add($"Unable to find by text of '{search}'.");

            return null;
        }

        public static IWebElement FindPartialLinkText(Logger logger, WebDriver driver, string search)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }

            if (driver == null)
            {
                throw new ArgumentNullException(nameof(driver));
            }

            int attempts = 0;

            do
            {
                try
                {
                    attempts += 1;

                    return driver.FindElement(By.PartialLinkText(search));
                }
                catch (ElementNotSelectableException)
                {
                    Thread.Sleep(100);
                }
                catch (NoSuchElementException)
                {
                    Thread.Sleep(100);
                }
                finally
                {

                }
            } while (attempts < 30);

            // Log
            logger.Add($"Unable to find by partial link text of '{search}'.");

            return null;
        }

        public static IWebElement FindCssSelector(Logger logger, WebDriver driver, string search)
        {
            if (logger == null)
            {
                throw new ArgumentNullException(nameof(logger));
            }

            if (driver == null)
            {
                throw new ArgumentNullException(nameof(driver));
            }

            int attempts = 0;

            do
            {
                try
                {
                    attempts += 1;

                    return driver.FindElement(By.CssSelector(search));
                }
                catch (ElementNotSelectableException)
                {                    
                    Thread.Sleep(100);
                }
                catch (NoSuchElementException)
                {
                    Thread.Sleep(100);
                }
                finally
                {

                }
            } while (attempts < 30);

            // Log
            logger.Add($"Unable to find by css selector of '{search}'.");

            return null;
        }

        public static DataTable ConvertCSVtoDataTable(string filePath)
        {
            using (StreamReader sr = new(filePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                using (DataTable dt = new())
                {
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        dt.Rows.Add(dr);
                    }
                    return dt;
                }
            }
        }

        private static string DayOfMonthSuffix(int day)
        {
            if (day == 1 || day == 21 || day == 31) { return "st"; };
            if (day == 2 || day == 22) { return "nd"; };
            if (day == 3 || day == 23) { return "rd"; };
            return "th";
        }

        public static void WriteDataToExcel(DataTable dt, DirectoryInfo diXLSXs, DirectoryInfo diPDFs, bool WriteToPDF)
        {
            if (dt == null)
            {
                throw new ArgumentNullException(nameof(dt));
            }

            if (diXLSXs == null)
            {
                throw new ArgumentNullException(nameof(diXLSXs));
            }

            if (diPDFs == null)
            {
                throw new ArgumentNullException(nameof(diPDFs));
            }

            string sDate = DateTime.Now.ToString("MMMM d", ci) + DayOfMonthSuffix(DateTime.Now.Day);
            string sDateEnd = DateTime.Now.AddDays(-180).ToString("MMMM d", ci) + DayOfMonthSuffix(DateTime.Now.Day);

            XLColor blackColor = XLColor.Black;            

            //Name of file
            string fileName = $"{Path.Combine(diXLSXs.FullName, dt.TableName)}.xlsx";

            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.Worksheets.Add(dt.TableName);

                foreach (DataColumn column in dt.Columns)
                {
                    IXLCell c1 = ws.Cell(4, dt.Columns.IndexOf(column) + 2);
                    c1.Value = column.ColumnName.Replace(" ", Environment.NewLine);
                    c1.Style.Fill.SetBackgroundColor(XLColor.DarkGray);
                    c1.Style.Font.SetFontColor(XLColor.White);

                    c1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    c1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                
                foreach (DataRow row in dt.Rows)
                {
                    foreach (DataColumn data_column in dt.Columns)
                    {
                        IXLCell c = ws.Cell(dt.Rows.IndexOf(row) + 5, dt.Columns.IndexOf(data_column) + 2);
                        c.Value = dt.Rows[dt.Rows.IndexOf(row)][dt.Columns[dt.Columns.IndexOf(data_column)]];

                        c.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        if (!String.IsNullOrEmpty(Convert.ToString(data_column.ExtendedProperties["Format"], ci)) && !string.IsNullOrEmpty(Convert.ToString(c.Value, ci)))
                        {
                            c.Value = Convert.ToDecimal(c.Value, ci).ToString(Convert.ToString(data_column.ExtendedProperties["Format"], ci), ci);
                        }

                        if (dt.Rows.IndexOf(row) % 2 != 0)
                        {
                            c.Style.Fill.SetBackgroundColor(XLColor.LightGray);
                        }
                    }
                }

                IXLCell firstCell = ws.FirstCellUsed();
                IXLCell lastCell = ws.LastCellUsed();

                foreach (IXLRow row in ws.Rows())
                {
                    foreach (IXLColumn data_column in ws.Columns())
                    {
                        IXLCell c = ws.Cell(row.RowNumber(), data_column.ColumnNumber());

                        if (string.IsNullOrEmpty(c.Value.ToString())
                            && c.Address.RowNumber >= ws.FirstCellUsed().Address.RowNumber
                            && c.Address.RowNumber <= ws.LastCellUsed().Address.RowNumber
                            && c.Address.ColumnNumber >= ws.FirstCellUsed().Address.ColumnNumber
                            && c.Address.ColumnNumber <= ws.LastCellUsed().Address.ColumnNumber)
                        {
                            c.Value = "'-";
                        }
                    }
                }

                //UPGRADE: use with after upgraded to C# 9.0
                IXLRange range = ws.Range(firstCell.Address, lastCell.Address);

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                range.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                range.Style.Border.TopBorder = XLBorderStyleValues.Thin;                              

                IXLRange borderTop = ws.Range(1, 1, 3, lastCell.Address.ColumnNumber + 1);
                borderTop.Style.Fill.SetBackgroundColor(blackColor);

                IXLRange borderTopTitle = ws.Range(2, 3, 2, lastCell.Address.ColumnNumber - 1);
                borderTopTitle.Merge();
                borderTopTitle.Style.Font.FontColor = XLColor.White;
                borderTopTitle.Style.Font.FontSize = 16;
                borderTopTitle.Style.Font.SetFontName("Candara");
                borderTopTitle.Style.Font.Bold = true;
                borderTopTitle.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                borderTopTitle.Value = $"Total Market Overview for {dt.TableName}, WA {sDateEnd} - {sDate}";
                
                IXLRange borderBottom = ws.Range(lastCell.Address.RowNumber + 1, 1, lastCell.Address.RowNumber + 5, lastCell.Address.ColumnNumber + 1);
                borderBottom.Style.Fill.SetBackgroundColor(blackColor);

                IXLRange borderLeft = ws.Range(1, 1, lastCell.Address.RowNumber + 1, 1);
                borderLeft.Style.Fill.SetBackgroundColor(blackColor);

                IXLRange borderRight = ws.Range(1, lastCell.Address.ColumnNumber + 1, lastCell.Address.RowNumber + 1, lastCell.Address.ColumnNumber + 1);
                borderRight.Style.Fill.SetBackgroundColor(blackColor);

                IXLRange nameCell = ws.Range(lastCell.Address.RowNumber + 2, lastCell.Address.ColumnNumber - 3, lastCell.Address.RowNumber + 2, lastCell.Address.ColumnNumber);
                nameCell.Merge();
                nameCell.Value = "Just Listed NW -KW Everett";
                nameCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                nameCell.Style.Font.Bold = true;
                nameCell.Style.Font.FontColor = XLColor.White;

                IXLRange addressCell = ws.Range(lastCell.Address.RowNumber + 3, lastCell.Address.ColumnNumber - 3, lastCell.Address.RowNumber + 3, lastCell.Address.ColumnNumber);
                addressCell.Merge();
                addressCell.Value = "1000 SE Everett Mall Way Ste 201";
                addressCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                addressCell.Style.Font.Bold = true;
                addressCell.Style.Font.FontColor = XLColor.White;

                IXLRange cityCell = ws.Range(lastCell.Address.RowNumber + 4, lastCell.Address.ColumnNumber - 3, lastCell.Address.RowNumber + 4, lastCell.Address.ColumnNumber);
                cityCell.Value = "Everett, WA 98208";
                cityCell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                cityCell.Style.Font.Bold = true;
                cityCell.Style.Font.FontColor = XLColor.White;
                cityCell.Merge();

                IXLCell CellForPDF = ws.Cell(lastCell.Address.RowNumber + 6, lastCell.Address.ColumnNumber + 2);
                CellForPDF.Value = " ";

                ws.Columns().AdjustToContents();
                ws.Column(firstCell.Address.ColumnNumber).Width = 12.14;

                //priceRangeCells
                ws.Range(firstCell.Address.RowNumber + 1, firstCell.Address.ColumnNumber, lastCell.Address.RowNumber, firstCell.Address.ColumnNumber).Style.Alignment.WrapText = true;
                ws.Rows().AdjustToContents();
                ws.Rows(6, lastCell.Address.RowNumber - 1).Height = 30.75;

                string targetPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string projectDirectory = Directory.GetParent(targetPath).Parent.FullName;
                string imageFile = Path.Combine(projectDirectory, "justListed.png");
                 
                var image = ws.AddPicture(imageFile)
                    .MoveTo(ws.Cell(lastCell.Address.RowNumber + 1, lastCell.Address.ColumnNumber - 6))
                    .Scale(.25);

                wb.SaveAs(fileName);

                if (WriteToPDF)
                {
                    ConvertExcelToPDF(new FileInfo(fileName), diPDFs);
                }
            }
        }

        private static void ConvertExcelToPDF(FileInfo fi, DirectoryInfo diPDFs)
        {
            if (fi == null)
            {
                throw new ArgumentNullException(nameof(fi));
            }

            if (diPDFs == null)
            {
                throw new ArgumentNullException(nameof(diPDFs));
            }

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            ExcelFile workbook = ExcelFile.Load(fi.FullName);
            
            ExcelWorksheet worksheet = workbook.Worksheets.ActiveWorksheet;

            worksheet.PrintOptions.FitWorksheetWidthToPages = 1;
            worksheet.PrintOptions.FitWorksheetHeightToPages = 1;
            worksheet.PrintOptions.HorizontalCentered = true;
            worksheet.PrintOptions.Portrait = false;

            workbook.Save(Path.Combine(diPDFs.FullName, $"{ Path.GetFileNameWithoutExtension(fi.FullName)}.pdf"));
        }
    }
}
