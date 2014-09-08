using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ServiceNowWebDriverExtensions
{
    namespace ExtensionMethods
    {
        public static class WebDriverExtensions
        {
            /// <summary>
            /// Lookup a group value in a reference field.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired reference field within the current document.</param>
            /// <param name="sendKeys">A property value of the current reference field.</param>
            public static void LookupGroup(this IWebDriver driver, By by, string sendKeys)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                IWebElement lookup = wait.Until(ExpectedConditions.ElementIsVisible(by));
                lookup.Click();

                System.Threading.Thread.Sleep(10000);

                string currentWindowHandle = driver.CurrentWindowHandle;
                string searchWindow = driver.WindowHandles.First(o => o != currentWindowHandle);
                driver.SwitchTo().Window(searchWindow);                               

                IWebElement item = wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(sendKeys)));

                item.Click();

                driver.SwitchTo().Window(currentWindowHandle);
                driver.SwitchTo().Frame("gsft_main");
            }

            /// <summary>
            /// Lookup a reference field value.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired reference field within the current document.</param>
            /// <param name="sendKeys">A property value of the current reference field.</param>
            public static void Lookup(this IWebDriver driver, By by, string sendKeys)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                IWebElement lookup = wait.Until(ExpectedConditions.ElementIsVisible(by));
                lookup.Click();

                string currentWindowHandle = driver.CurrentWindowHandle;
                string searchWindow = driver.WindowHandles.First(o => o != currentWindowHandle);
                driver.SwitchTo().Window(searchWindow);
                
                IWebElement searchFor = wait.Until(ExpectedConditions.ElementExists(By.ClassName("list_search_text")));
                searchFor.Click();
                searchFor.Clear();
                searchFor.SendKeys(sendKeys);
                
                var selects = driver.FindElements(By.ClassName("list_search_select"));
                IWebElement search = selects.First(o => o.GetAttribute("id").Contains("_select"));
                driver.SelectOption(search, "Name"); //"for text");

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[title='Go']"))).Click();

                IWebElement item = wait.Until(ExpectedConditions.ElementExists(By.LinkText(sendKeys)));

                item.Click();

                driver.SwitchTo().Window(currentWindowHandle);
                driver.SwitchTo().Frame("gsft_main");
            }

            /// <summary>
            /// Lookup a reference field value.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired reference field within the current document.</param>
            /// <param name="what">A property of the current reference field.</param>
            /// <param name="sendKeys">A property value of the current reference field.</param>
            public static void Lookup(this IWebDriver driver, By by, string what, string sendKeys)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                
                wait.Until(ExpectedConditions.ElementIsVisible(by)).Click();

                string currentWindowHandle = driver.CurrentWindowHandle;
                string searchWindow = driver.WindowHandles.First(o => o != currentWindowHandle);
                driver.SwitchTo().Window(searchWindow);

                Search(driver, what, sendKeys);
                System.Threading.Thread.Sleep(1000);
                wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText(sendKeys))).Click();

                driver.SwitchTo().Window(currentWindowHandle);
                driver.SwitchTo().Frame("gsft_main");
            }

            /// <summary>
            /// Search a list by a specified column and value pair.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="what">A property of the current reference field.</param>
            /// <param name="sendKeys">A property value of the current reference field.</param>
            public static void Search(this IWebDriver driver, string what, string sendKeys)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                IWebElement searchFor = wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("list_search_text")));
                searchFor.Click();
                searchFor.Clear();
                searchFor.SendKeys(sendKeys);

                var selects = driver.FindElements(By.ClassName("list_search_select"));
                IWebElement search = selects.First(o => o.GetAttribute("id").Contains("_select"));
                driver.SelectOption(search, what); //"for text");

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("img[title='Go']"))).Click();
            }

            /// <summary>
            /// Select a specific value in a dropdown list.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired list field within the current document.</param>
            /// <param name="value">A desired value of the current list field.</param>
            public static void SelectOption(this IWebDriver driver, By by, string value)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                IWebElement select = wait.Until(ExpectedConditions.ElementIsVisible(by));
                List<IWebElement> options = select.FindElements(By.TagName("option")).ToList();
                options.First(o => o.Text == value).Click();
            }

            /// <summary>
            /// Select a specific value in a dropdown list.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="select">A list field within the current document.</param>
            /// <param name="value">A desired value of the current list field.</param>
            public static void SelectOption(this IWebDriver driver, IWebElement select, string value)
            {
                List<IWebElement> options = select.FindElements(By.TagName("option")).ToList();
                options.First(o => o.Text == value).Click();
            }

            /// <summary>
            /// Retrieve the currently selected value in a dropdown list.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired list field within the current document.</param>
            /// <returns></returns>
            public static string GetSelectedOption(this IWebDriver driver, By by)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                IWebElement item = wait.Until(ExpectedConditions.ElementIsVisible(by));
                return (new SelectElement(item).SelectedOption.Text);
            }

            /// <summary>
            /// Right click an item in within the current document.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired item within the current document.</param>
            public static void RightClick(this IWebDriver driver, By by)
            {
                RightClick(driver, driver.FindElement(by));
            }

            /// <summary>
            /// Right click an item in within the current document.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="oWE">An item within the current document.</param>
            public static void RightClick(this IWebDriver driver, IWebElement oWE)
            {
                Actions oAction = new Actions(driver);
                oAction.MoveToElement(oWE);
                oAction.ContextClick(oWE).Build().Perform();
            }

            /// <summary>
            /// Scroll a specified element into view.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="element"></param>
            public static void ScrollTo(this IWebDriver driver, IWebElement element)
            {
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", element);
            }

            /// <summary>
            /// Detect whether an alert has occurred.
            /// </summary>
            /// <param name="driver"></param>
            /// <returns></returns>
            public static Boolean isAlertPresent(this IWebDriver driver)
            {
                try
                {
                    driver.SwitchTo().Alert();
                    return true;
                }   // try 
                catch (NoAlertPresentException Ex)
                {
                    return false;
                }   // catch 
            }   // isAlertPresent()

            /// <summary>
            /// Reload a form by right-clicking on its title bar.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="titleBar">A mechanism by which to find the desired form title bar within the current document.</param>
            /// <param name="btnReload">A mechanism by which to find the desired reload option within the current choice list.</param>
            public static void ReloadForm(this IWebDriver driver, By titleBar, By btnReload)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                driver.RightClick(titleBar);
                wait.Until(ExpectedConditions.ElementIsVisible(btnReload)).Click();
            }

            /// <summary>
            /// Collapse all applications in the left navigation menu.
            /// </summary>
            /// <param name="driver"></param>
            public static void CollapseAllAplications(this IWebDriver driver)
            {
                driver.SwitchTo().Frame("gsft_main");

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                // Add content »
                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#top_link_add > a:nth-child(1)")));

                driver.SwitchTo().DefaultContent();

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("a[title=\"Collapse All Applications\"] > img"))).Click();

                driver.SwitchTo().Frame("gsft_nav");
            }

            /// <summary>
            /// Log into a ServiceNow instance.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="uid">User ID</param>
            /// <param name="password">Password</param>
            /// <param name="url">A proxy URL that will lead to a ServiceNow instance.</param>
            public static void Login(this IWebDriver driver, string uid, string password, string url)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                driver.Navigate().GoToUrl(url);
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("signin"))).Click();
                driver.SwitchTo().DefaultContent();

                SignIn(driver, uid, password);
            }

            /// <summary>
            /// Log into a ServiceNow instance.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="uid">User ID</param>
            /// <param name="password">Password</param>
            /// <param name="url">URL to ServiceNow instance.</param>
            public static void LoginDirect(this IWebDriver driver, string uid, string password, string url)
            {
                driver.Navigate().GoToUrl(url);

                SignIn(driver, uid, password);
            }

            private static void SignIn(IWebDriver driver, string uid, string password)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("USERNAME"))).SendKeys(uid);

                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("PASSWORD"))).SendKeys(password);

                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("signin"))).Click();
                
                driver.SwitchTo().DefaultContent();
            }

            /// <summary>
            /// Log out of a ServiceNow instance.
            /// </summary>
            /// <param name="driver"></param>
            public static void Logout(this IWebDriver driver)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                driver.SwitchTo().DefaultContent();

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#gsft_logout > button"))).Click();

                driver.SwitchTo().DefaultContent();

                SignOut(driver);
            }

            /// <summary>
            /// Log out of a ServiceNow Portal instance.
            /// </summary>
            /// <param name="driver"></param>
            public static void PortalLogout(this IWebDriver driver)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                driver.SwitchTo().DefaultContent();

                wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText("Logout"))).Click();

                driver.SwitchTo().DefaultContent();

                SignOut(driver);
            }

            private static void SignOut(IWebDriver driver)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                IWebElement signIn = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("signin")));
                signIn.Click();
                driver.SwitchTo().DefaultContent();

                wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText("Log Out"))).Click();

                driver.SwitchTo().DefaultContent();
            }

            /// <summary>
            /// Click on a desired element within a document (with reference to Internet Explorer).
            /// </summary>
            /// <param name="element"></param>
            /// <param name="browser">Current webdriver instance.</param>
            public static void ForceClick(this IWebElement element, IWebDriver browser)
            {
                if (typeof(InternetExplorerDriver) == browser.GetType())
                {
                    element.SendKeys(Keys.Enter);
                }
                else
                {
                    element.Click();
                }
            }

            /// <summary>
            /// Click on a desired element within a document (with reference to Internet Explorer).
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="by">A mechanism by which to find the desired item within the current document.</param>
            public static void ForceClick(this IWebDriver driver, By by)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(by));

                if (typeof(InternetExplorerDriver) == driver.GetType())
                {
                    element.SendKeys(Keys.Enter);
                }
                else
                {
                    element.Click();
                }
            }

            /// <summary>
            /// Filter a listview of a table.
            /// </summary>
            /// <param name="driver"></param>
            /// <param name="field">Table Field on which to filter.</param>
            /// <param name="oper">Operation</param>
            /// <param name="value">Desired value of the current table field on which to filter.</param>
            public static void FilterList(this IWebDriver driver, string field, string oper, string value)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("list_filter_toggle"))).Click();

                driver.SelectOption(By.CssSelector("#field > select"), field);

                driver.SelectOption(By.CssSelector("#oper > select"), oper);

                wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("#value > input:nth-child(2)")))
                    .SendKeys(value);

                wait.Until(ExpectedConditions.ElementIsVisible(By.LinkText("Run"))).Click();
            }
        }
    }
}
