using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium.Interactions;
using DocumentFormat.OpenXml.Office.Excel;

namespace RegisterUserNew
{
    public class RegisterUser
    {
       
        static void Main(String[] args)
        {
            //Open Chrome Driver
            IWebDriver driver = new ChromeDriver(@"C:\Users\27828\source\repos\");
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //Path to your excel file
            string path = "C:\\DATA\\RegisterNewUser.xlsx";
            FileInfo fileInfo = new FileInfo(path);

            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            //Get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            //Initialise variables
            int i = 2;
            int j = 1;

            //Open Environment
            string URl = worksheet.Cells[i, j].Value.ToString();
            driver.Navigate().GoToUrl((string)URl);
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

            //Create an Account
            IWebElement RegisterLink = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[3]/div/div[1]/a"));
            RegisterLink.Click();

            // loop through the worksheet rows and columns
            while ((i <= rows) && (i < 3))
            {
                //Enter the Email Address
                string UserName = worksheet.Cells[i, 2].Value.ToString();
                IWebElement USerName = driver.FindElement(By.Id("Input_Username"));
                USerName.SendKeys((string)UserName);

                //Enter the Email Address
                string EmailAddr = worksheet.Cells[i, 3].Value.ToString();
                IWebElement EmailAddress = driver.FindElement(By.Id("Input_Email"));
                EmailAddress.SendKeys((string)EmailAddr);

                //Enter the Password
                string PassW = worksheet.Cells[i, 4].Value.ToString();
                IWebElement PassWOrd = driver.FindElement(By.Id("Input_Password"));
                PassWOrd.SendKeys((string)PassW);

                //Confirm Password
                string ConfPassW = worksheet.Cells[i, 5].Value.ToString();
                IWebElement ConfirmPassWOrd = driver.FindElement(By.Id("Input_ConfirmPassword"));
                ConfirmPassWOrd.SendKeys((string)ConfPassW);

                //Click on the 'I have read and agree to the terms and conditions' Checkbox
                IWebElement CheckBoxTermsConditions = driver.FindElement(By.Id("TermsConditions"));
                CheckBoxTermsConditions.Click();
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                //Click on the Register button
                Actions action1 = new Actions(driver);
                action1.MoveToElement(driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div[2]/form/div[11]/div/button"))).Build().Perform();
                IWebElement RegButton = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div[2]/form/div[11]/div/button"));
                RegButton.Click();

                //Get Message to determine if Registration was successful or not
                string Message = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[1]/h2")).Text;

                if (Message == "REGISTRATION CONFIRMATION SUCCESSFUL")
                {

                    i = i + 1;

                    if (i > rows)
                    {
                        break;
                    }

                }
                
                while (i > 2)
                {
                    IWebElement ClickOnLoginLink = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[3]/div/div[1]/a"));
                    ClickOnLoginLink.Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                    //Create an Account
                    IWebElement RegisterLink2 = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[3]/div/div[1]/a"));
                    RegisterLink2.Click();

                    //Enter the Email Address
                    string UserName2 = worksheet.Cells[i, 2].Value.ToString();
                    IWebElement USerName2 = driver.FindElement(By.Id("Input_Username"));
                    USerName2.SendKeys((string)UserName2);

                    //Enter the Email Address
                    string EmailAddr2 = worksheet.Cells[i, 3].Value.ToString();
                    IWebElement EmailAddress2 = driver.FindElement(By.Id("Input_Email"));
                    EmailAddress2.SendKeys((string)EmailAddr2);

                    //Enter the Password
                    string PassW2 = worksheet.Cells[i, 4].Value.ToString();
                    IWebElement PassWOrd2 = driver.FindElement(By.Id("Input_Password"));
                    PassWOrd2.SendKeys((string)PassW2);

                    //Confirm Password
                    string ConfPassW2 = worksheet.Cells[i, 5].Value.ToString();
                    IWebElement ConfirmPassWOrd2 = driver.FindElement(By.Id("Input_ConfirmPassword"));
                    ConfirmPassWOrd2.SendKeys((string)ConfPassW2);

                    //Click on the 'I have read and agree to the terms and conditions' Checkbox
                    IWebElement CheckBoxTermsConditions2 = driver.FindElement(By.Id("TermsConditions"));
                    CheckBoxTermsConditions2.Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                    //Click on the Register button
                    Actions action2 = new Actions(driver);
                    action2.MoveToElement(driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div[2]/form/div[11]/div/button"))).Build().Perform();
                    IWebElement RegButton2 = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div[2]/form/div[11]/div/button"));
                    RegButton2.Click();

                    //Get Message to determine if Registration was successful or not
                    string Message2 = driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/div/div[1]/h2")).Text;

                    if (Message2 == "REGISTRATION CONFIRMATION SUCCESSFUL")
                    {

                        i = i + 1;

                        if (i > rows)
                        {
                            break;
                        }

                    }
                }
            }

            //Quit Browser
            driver.Quit();
            package.Dispose();
        }

    }
}
