using System;
using System.IO;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.SDK.Web;
using AdVantage_AppModel;                       //Add a Reference to the seperate Application Model Project
using Excel = Microsoft.Office.Interop.Excel;   //Add a Reference to the EXCEL namespace

//  An Additionally reference has been added, right-click 'References' under solution explorer.  Then select
// 'Assemblies'->'Framework'->'Microsoft.CSharp'.  This is needed for the EXCEL access



namespace AdVantage_On_Line_Banking_Unit_Tests
{
    [TestFixture]
    public class LeanFtTest : UnitTestClassBase
    {
        // Set up global variables for the instance of the browser and the Application Model

        private IBrowser AdVantBrowser;
        private VantModel MyVantModel;

        [TestFixtureSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture

            // Optional to turn on snapshots on every step.  It can be set here this way
            // or by making modifications to the App.config file
            Reporter.SnapshotCaptureLevel = HP.LFT.Report.CaptureLevel.All;

            // Initially create an instance of the browser and navigate to the AdVantage home page
            AdVantBrowser = BrowserFactory.Launch(BrowserType.InternetExplorer);
            AdVantBrowser.Navigate("http://alm-aob:47001/advantage/");
            //AdVantBrowser.Navigate("http://15.126.221.115:47001/advantage/");

            // Create a new instance of the Application Model

            MyVantModel = new VantModel(AdVantBrowser);

        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
            // Login to the AdVantage using the criteria supplied below.

            MyVantModel.AdvantageOnlineBankingPage.Username.SetValue("jojo");
            MyVantModel.AdvantageOnlineBankingPage.Password.SetSecure("55b53bda2f699367c92559ae7afe46feee7e11975800086ca742b0a12fd9");
            MyVantModel.AdvantageOnlineBankingPage.Login.Click();



        }
        // Declare an array to hold the various screens we are going to validate are available

        static string[] strMenuItems = new string[] { "Accounts", "Bill Pay", "Money Transfer", "Brokerage", "Create Account", "Credit Cards", "Order Checkbooks" };

        // Declate the location of the datafile for transfering amounts

        static string strDataSheet = @"C:\Users\hpswadm\Desktop\Demo Information\LeanFT\Visual Studio\AdVantage On-Line Banking Unit Tests\DataSheets\MoneyTransfer.xlsx";

        // First test is to validate that a screen is available after the build.  Use the global array to set
        // each of the screens to be validate.

        [Test, TestCaseSource("strMenuItems")]
        public void TestScreenExists(string strScreenName)
        {

            // Click on the link in the left hand menu based on the passed in screen name

            AdVantBrowser.Describe<ILink>(new LinkDescription
            {
                TagName = @"A",
                InnerText = @strScreenName
            }).Click();

            // Set a variable to hold the object for the screen banner.  Note, the banner is in uppercase, therefore
            // the screenname must be converted to uppercase.

            var objScreenBanner = AdVantBrowser.Describe<IWebElement>(new WebElementDescription
            {
                ClassName = @"center-name-text",
                TagName = @"SPAN",
                InnerText = @strScreenName.ToUpper()
            });

            // If the object exists on the screen, the test has passed.  Otherwise we have a failure.

            if (objScreenBanner.Exists())
            {
                Reporter.ReportEvent("Availablity", strScreenName + " screen is Available", HP.LFT.Report.Status.Passed, AdVantBrowser.GetSnapshot());
            }
            else
            {
                Reporter.ReportEvent("Availablity", strScreenName + " screen is NOT Available", HP.LFT.Report.Status.Failed, AdVantBrowser.GetSnapshot());
                Assert.Fail();
            }

        }
        [Test]

        // Test to transfer money from one account to another.

        public void TransferMoney()
        {
            // Set up acccess to various EXCEL assets

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range DataRange;

            // Declare the strings to hold the from account, to account and the amount to be transferred

            string strFromAccount, strToAccount, strAmount;

            // Define the Row Count, the columns in the spreadsheet that hold the 'From Account', 'To Account' and 'Amount' respectively.

            int rCnt = 0;
            int FrmAccountCol = 1;
            int ToAccountCol = 2;
            int AmountCol = 3;

            // Open a new instance EXCEL.  Then open the data sheet that's defined the the global variable strDataSheet

            xlApp = new Excel.Application();

            xlWorkBook = xlApp.Workbooks.Open(strDataSheet, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            // Get the first work sheet

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Find the 'used range' of the excel sheet

            DataRange = xlWorkSheet.UsedRange;

            // Starting at row 2 (row 1 is the header information) loop through each row that has data

            for (rCnt = 2; rCnt <= DataRange.Rows.Count; rCnt++)
            {
                // First click on the 'Money Transfer' link in the left menu

                AdVantBrowser.Describe<ILink>(new LinkDescription
                {
                    TagName = @"A",
                    InnerText = @"Money Transfer"
                }).Click();

                // Read each column text for the current loop item

                strFromAccount = (string)(DataRange.Cells[rCnt, FrmAccountCol] as Excel.Range).Text;
                strToAccount = (string)(DataRange.Cells[rCnt, ToAccountCol] as Excel.Range).Text;
                strAmount = (string)(DataRange.Cells[rCnt, AmountCol] as Excel.Range).Text;

                // As the data in the EXCEL sheet does not contain ALL the text in the selection boxes 'From Account'
                // and 'To account' (i.e. the selection box items contain the Account type, the account number AND
                // the amount currently in that account [which obviously changes]) we have to look through
                // each of the items in the selection box and see if it 'contains' the data in the EXCEL sheet 
                // (which is only the Account Type and the Account Number).  Therefore, first loop through all
                // of the available selections in the list.


                foreach (IListItem MyItems in MyVantModel.AdvantageOnlineBankingPage.FromAccount.Items)
                {
                    // If the current list item contains the text from the spreadsheet - Then reset the
                    // string holding the account number to be the full text and then select it from the list

                    if (MyItems.Text.Contains(strFromAccount))
                    {
                        strFromAccount = MyItems.Text;
                        MyVantModel.AdvantageOnlineBankingPage.FromAccount.Select(strFromAccount);

                    }
                    else if (MyItems.Text.Contains(strToAccount))
                    {

                        strToAccount = MyItems.Text;
                        MyVantModel.AdvantageOnlineBankingPage.ToAccount.Select(strToAccount);
                    }
                }
                // Click the 'Next' button.

                MyVantModel.AdvantageOnlineBankingPage.NextButton.Click();

                // Set the amount value, the transfer date to todays date (converted appropriately) then
                // click the 'Next' button.

                MyVantModel.AdvantageOnlineBankingPage.Amount.SetValue(strAmount);
                MyVantModel.AdvantageOnlineBankingPage.TransferDate.SetValue(DateTime.Now.ToString("MM/dd/yyyy"));
                MyVantModel.AdvantageOnlineBankingPage.NextButton.Click();

                // Click Ok to complete the transfer

                MyVantModel.AdvantageOnlineBankingPage.OKButton.Click();

                // Validate the message that the money has been transfer.  Note, we have to find the amount via
                // a regular expression (See Application Model) as the inner text contains the ACTUAL account.

                if (MyVantModel.AmountTransfered.InnerText.Contains(strAmount))
                {
                    Reporter.ReportEvent("MoneyTransfer", "Money Transfered Successfully", HP.LFT.Report.Status.Passed, AdVantBrowser.GetSnapshot());

                }
                else
                {
                    Reporter.ReportEvent("MoneyTransfer", "Money Not Transfered", HP.LFT.Report.Status.Failed, AdVantBrowser.GetSnapshot());
                    Assert.Fail();

                }

            }


        }


        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            // Click the Logout

            if (MyVantModel.AdvantageOnlineBankingPage.Logout.Exists())
            {
                MyVantModel.AdvantageOnlineBankingPage.Logout.Click();
            }
            else
            {
                // if the screen does not have a 'logout' link, we need to report it.

                Reporter.ReportEvent("Missing Link", "The Following URL is missing the Logout Link: " + AdVantBrowser.URL, HP.LFT.Report.Status.Warning, AdVantBrowser.GetSnapshot());

                AdVantBrowser.Navigate("http://alm-aob:47001/advantage/");
                //AdVantBrowser.Navigate("http://15.126.221.115:47001/advantage/");
                MyVantModel.AdvantageOnlineBankingPage.Logout.Click();
            }


        }


        [TestFixtureTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
            // Close the browser

            AdVantBrowser.Close();

            // The following are optional lines if you wish the report to open
            // automatically at the end of the execution
            Reporter.GenerateReport();
            IBrowser browser = BrowserFactory.Launch(BrowserType.Chrome);
            browser.Navigate(@"file:///" + Directory.GetCurrentDirectory() + @"\RunResults\runresults.html");
        }
    }
}
