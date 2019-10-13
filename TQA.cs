using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace VismaTQA
{
    [TestClass]
    public class Visma
    {

        [TestMethod]
        public void TestQA()
        {
            //Validation messages
            String validation = "Šis lauks ir obligāts.";
            String validation_cb = "Mums ir nepieciešama Jūsu piekrišana Šis lauks ir obligāts.";
            String validation_correct_email = "Ievadiet derîgu epasta adresi.";

            String url = "https://www.visma.lv/";

            using (ChromeDriver driver = new ChromeDriver())
            {
                IJavaScriptExecutor js = driver as IJavaScriptExecutor;

                // Maximize the browser
                driver.Manage().Window.Maximize();

                // 1

                // Go to the "Visma" homepage
                driver.Navigate().GoToUrl(url);

                // 2
                
                // Go to the presentation request page
                driver.FindElement(By.ClassName("cta")).Click();

                // Waiting for VismaCookieConsentText
                Thread.Sleep(7000);

                // Check if VismaCookieConsentText appears
                if (driver.FindElement(By.ClassName("modal__action")).Displayed)
                {
                    driver.FindElement(By.ClassName("modal__action")).Click();
                }

                Thread.Sleep(1000);

                //3

                //Scroll down to Submit button
                js.ExecuteScript("window.scrollBy(0,250);");
                Thread.Sleep(1000);

                //Checking validation for required fields
                driver.FindElement(By.ClassName("ctavisma")).Click();
                Thread.Sleep(1000);

                String validation_name = driver.FindElement(By.XPath("//*[@id='__field_']/div[2]/span")).Text;
                String validation_lastname = driver.FindElement(By.XPath("//*[@id='__field_']/div[4]/span")).Text;
                String validation_organization = driver.FindElement(By.XPath("//*[@id='__field_']/div[6]/span")).Text;
                String validation_phone = driver.FindElement(By.XPath("//*[@id='__field_']/div[8]/span")).Text;
                String validation_email = driver.FindElement(By.XPath("//*[@id='__field_']/div[10]/span")).Text;
                String validation_checkbox = driver.FindElement(By.XPath("//*[@id='311b75c7-650b-4977-8a81-3524032e8081_terms_container']/div/div[1]/div[1]/span")).Text;

                Assert.AreEqual(validation_name, validation);
                Assert.AreEqual(validation_lastname, validation);
                Assert.AreEqual(validation_organization, validation);
                Assert.AreEqual(validation_phone, validation);
                Assert.AreEqual(validation_email, validation);
                Assert.AreEqual(validation_checkbox, validation_cb);

                //Fill-in the form and press checkbox
                driver.FindElement(By.Name("__field_608010")).SendKeys("Name");
                driver.FindElement(By.Name("__field_608011")).SendKeys("Last Name");
                driver.FindElement(By.Name("__field_608012")).SendKeys("Organization");
                driver.FindElement(By.Name("__field_608014")).SendKeys("Phone"); // No validation on letters
                driver.FindElement(By.Name("__field_608013")).SendKeys("email");
                driver.FindElement(By.ClassName("terms-checkbox-container")).Click();

                js.ExecuteScript("window.scrollBy(0,300);");
                Thread.Sleep(1000);
                driver.FindElement(By.ClassName("ctavisma")).Click();
                Thread.Sleep(1000);

                //Checking validation for incorrect email
                String correct_email = driver.FindElement(By.XPath("//*[@id='__field_']/div[10]/span")).Text;
                Assert.AreEqual(correct_email, validation_correct_email);
                Thread.Sleep(1000);

                // 4

                // Go back to the "Visma" homepage
                driver.Navigate().GoToUrl(url);
                //Scroll down to blogs
                js.ExecuteScript("window.scrollBy(0,2500);");
                Thread.Sleep(1000);
                //Check that blog links opens new tab with different blog records 
                for (int i = 1; i < 4; i++)
                {
                    //Wait until element is visible
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                    //Get name of blog link
                    String blog = driver.FindElement(By.XPath("/html/body/div[1]/main/section[4]/div/ul/li[" + i + "]/a")).GetAttribute("title");
                    //Click to blog
                    driver.FindElement(By.XPath("/html/body/div[1]/main/section[4]/div/ul/li[" + i + "]/a")).Click();
                    IList<String> tabs = driver.WindowHandles;
                    Thread.Sleep(3000);
                    driver.SwitchTo().Window(tabs[1]);
                    Thread.Sleep(1000);
                    //Get title of blog
                    String check_blog = driver.FindElement(By.TagName("h1")).Text;
                    //Compare the blog link with title
                    Assert.AreEqual(blog, check_blog);
                    driver.Close();
                    Thread.Sleep(1000);
                    driver.SwitchTo().Window(tabs[0]);
                    Thread.Sleep(3000);
                    // Check if VismaCookieConsentText appears
                    if (driver.FindElement(By.ClassName("modal__action")).Displayed)
                    {
                        driver.FindElement(By.ClassName("modal__action")).Click();
                    }
                    Thread.Sleep(3000);
                }

                // 5

                //Scroll down to social network links
                js.ExecuteScript("window.scrollBy(0,3150);");
                Thread.Sleep(1000);
                //Ccheck that all social network links works and opens correctly
                for (int i = 1; i < 4; i++)
                {
                    //Get name of social network
                    String socialNetwork = driver.FindElement(By.XPath("/html/body/div[1]/footer/div[1]/p[" + i + "]/a")).Text;
                    //Click to social network link
                    driver.FindElement(By.XPath("/html/body/div[1]/footer/div[1]/p[" + i + "]/a")).Click();
                    IList<String> tabs = driver.WindowHandles;
                    Thread.Sleep(1000);
                    driver.SwitchTo().Window(tabs[1]);
                    Thread.Sleep(1000);
                    //Get url of social network
                    String currentURL = driver.Url;
                    //Check that url contains the social network name
                    Assert.IsTrue(currentURL.Contains(socialNetwork.ToLower()), currentURL + " doesn't contains " + socialNetwork);
                    driver.Close();
                    Thread.Sleep(1000);
                    driver.SwitchTo().Window(tabs[0]);
                    Thread.Sleep(3000);
                    // Check if VismaCookieConsentText appears
                    if (driver.FindElement(By.ClassName("modal__action")).Displayed)
                    {
                        driver.FindElement(By.ClassName("modal__action")).Click();
                    }
                    Thread.Sleep(3000);
                }

                // 6

                //Check that pages goes to visma.co.uk site after changing the language
                driver.FindElement(By.ClassName("choices")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/div[1]/footer/div[6]/div/div[2]/div/div[10]")).Click();
                Thread.Sleep(1000);
                String ukURL = driver.Url;
                Assert.AreEqual(ukURL, "https://www.visma.co.uk/");
                String language = driver.FindElement(By.ClassName("choices__item")).GetAttribute("data-value");
                Assert.IsTrue(language.Contains("uk"), language + " doesn't contains uk");

                driver.Close();
                driver.Quit();

            }
        }
    }
}
