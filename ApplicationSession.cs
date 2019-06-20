//Based off of Paint3DTest https://github.com/Microsoft/WinAppDriver/tree/master/Samples/C%23/Paint3DTest

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using System;

namespace Paint3DTest
{
    public class ApplicationSession
    {
        // Note: append /wd/hub to the URL if you're directing the test at Appium
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723/wd/hub";
        private const string Paint3DAppId = @"Microsoft.MSPaint_8wekyb3d8bbwe!Microsoft.MSPaint";

        protected static WindowsDriver<WindowsElement> session;

        public static void Setup(TestContext context)
        {
            // Launch Paint 3D application if it is not yet launched
            if (session == null)
            {
                // Create a new session to launch Paint 3D application
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                //appCapabilities.SetCapability("app", Paint3DAppId);
                appCapabilities.SetCapability("deviceName", "WindowsPC");
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);

                //// Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
                //session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);

                //// Maximize Paint 3D window to ensure all controls being displayed
                //session.Manage().Window.Maximize();
            }
        }

        public static void TearDown()
        {
            // Close the application and delete the session
            if (session != null)
            {
                ClosePaint3D();
                session.Quit();
                session = null;
            }
        }

        [TestInitialize]
        public void CreateNewPaint3DProject()
        {
            //try
            //{
            //    session.FindElementByAccessibilityId("WelcomeScreenNewButton").Click();
            //    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
            //}
            //catch
            //{
            //    // Create a new Paint 3D project by pressing Ctrl + N
            //    session.SwitchTo().Window(session.CurrentWindowHandle);
            //    session.Keyboard.SendKeys(Keys.Control + "n" + Keys.Control);
            //    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
            //    DismissSaveConfirmDialog();
            //}
        }

        private static void DismissSaveConfirmDialog()
        {
            try
            {
                WindowsElement closeSaveConfirmDialog = session.FindElementByAccessibilityId("CloseSaveConfirmDialog");
                closeSaveConfirmDialog.FindElementByAccessibilityId("SecondaryBtnG3").Click();
            }
            catch { }
        }

        private static void ClosePaint3D()
        {
            try
            {
                session.Close();
                string currentHandle = session.CurrentWindowHandle; // This should throw if the window is closed successfully

                // When the Paint 3D window remains open because of save confirmation dialog, attempt to close modal dialog
                DismissSaveConfirmDialog();
            }
            catch { }
        }
    }
}
