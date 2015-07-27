using System;
using System.Linq;
using NUnit.Framework;
using HP.LFT.Report;
using HP.LFT.UnitTesting;

namespace AdVantage_On_Line_Banking_Unit_Tests
{
    [TestFixture]
    public abstract class UnitTestClassBase : UnitTestBase
    {
        [TestFixtureSetUp]
        public void GlobalSetup()
        {
            TestSuiteSetup();
        }

        [TestFixtureTearDown]
        public void GlobalTearDown()
        {
            TestSuiteTearDown();
            Reporter.GenerateReport();
        }

        [SetUp]
        public void BasicSetUp()
        {
            TestSetUp();
        }

        [TearDown]
        public void BasicTearDown()
        {
            TestTearDown();
        }

        protected override string GetClassName()
        {
            return TestContext.CurrentContext.Test.FullName;
        }

        protected override string GetTestName()
        {
            return TestContext.CurrentContext.Test.Name;
        }

        protected override Status GetFrameworkTestResult()
        {
            switch (TestContext.CurrentContext.Result.State)
            {
                case TestState.Failure:
                case TestState.Error:
                    return Status.Failed;
                case TestState.Cancelled:
                    return Status.Warning;
                case TestState.Success:
                    return Status.Passed;
                default:
                    return Status.Passed;
            }
        }
    }
}
