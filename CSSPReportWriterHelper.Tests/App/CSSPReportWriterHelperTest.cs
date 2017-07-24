using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CSSPReportWriterHelper.Tests.SetupInfo;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using CSSPReportWriterHelperDLL.Services;

namespace CSSPReportWriterHelper.Tests.App
{
    /// <summary>
    /// Summary description for BaseServiceTest
    /// </summary>
    [TestClass]
    public class CSSPReportWriterHelperTest : SetupData
    {
        #region Variables
        private TestContext testContextInstance;
        private SetupData setupData;

        #endregion Variables

        #region Properties
        public CSSPReportWriterHelper csspReportWriterHelper { get; set; }
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion
        #endregion Properties

        #region Constructors
        public CSSPReportWriterHelperTest()
        {
            setupData = new SetupData();
        }
        #endregion Constructors

        #region Testing Functions public
        [TestMethod]
        public void BaseService_Constructors_Test()
        {
            foreach (CultureInfo culture in setupData.cultureListGood)
            {
                SetupTest(culture);

                if (csspReportWriterHelper.StartWebAddressCSSP.Contains("localhost"))
                {
                    Assert.AreEqual("http://localhost:11562/", csspReportWriterHelper.StartWebAddressCSSP);
                }
                else
                {
                    Assert.AreEqual("http://wmon01dtchlebl2/csspwebtools/", csspReportWriterHelper.StartWebAddressCSSP);
                }
                Assert.AreEqual("", csspReportWriterHelper.reportBaseService.LastHref);
                Assert.AreEqual("", csspReportWriterHelper.reportBaseService.LastCSSPTVText);
                Assert.IsFalse(csspReportWriterHelper.WebIsVisible);
                Assert.IsNotNull(csspReportWriterHelper.reportBaseService);
            }
        }
        [TestMethod]
        public void BaseService_GetID_TVText_Test()
        {
            foreach (CultureInfo culture in setupData.cultureListGood)
            {
                SetupTest(culture);

                string url = @"http://localhost:11562/en-CA/#!View/All locations|||1|||30000000000000000000000000000000";
                csspReportWriterHelper.GetID_TVText(url);
                Assert.AreEqual(url, csspReportWriterHelper.reportBaseService.LastHref);
                Assert.AreEqual("All locations", csspReportWriterHelper.reportBaseService.LastCSSPTVText);
            }
        }
        #endregion Functions public

        #region Functions
        public void SetupTest(CultureInfo culture)
        {
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;

            csspReportWriterHelper = new CSSPReportWriterHelper();
        }
        private void SetupShim()
        {
            //shimClimateSiteService = new ShimClimateSiteService(climateSiteService);
        }
        #endregion Functions




    }
}
