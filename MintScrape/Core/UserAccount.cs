using System;
using System.Configuration;
using MintScrape.Interfaces;
using OpenQA.Selenium.Remote;

namespace MintScrape.Core {
    public class UserAccount : IUserAccount {
        private readonly RemoteWebDriver _driver;
        private readonly string _loginButtonId = ConfigurationManager.AppSettings["LoginButtonId"];
        private readonly string _loginPage = ConfigurationManager.AppSettings["LoginPage"];
        private readonly string _logoutLinkId = ConfigurationManager.AppSettings["LogoutLinkId"];
        private readonly string _password = ConfigurationManager.AppSettings["Password"];
        private readonly string _passwordFieldId = ConfigurationManager.AppSettings["PasswordFieldId"];
        private readonly string _userName = ConfigurationManager.AppSettings["UserName"];
        private readonly string _userNameFieldId = ConfigurationManager.AppSettings["UserNameFieldId"];

        /// <summary>
        ///     Constructor.
        /// </summary>
        /// <param name="driver"></param>
        public UserAccount(RemoteWebDriver driver) {
            _driver = driver;
        }

        public string Password { get; set; }
        public string UserName { get; set; }

        /// <summary>
        ///     Login to Mint.
        /// </summary>
        public void Login() {
            _driver.Navigate().GoToUrl(_loginPage);
            _driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5));
            var userNameField = _driver.FindElementById(_userNameFieldId);
            var userPasswordField = _driver.FindElementById(_passwordFieldId);
            var loginButton = _driver.FindElementById(_loginButtonId);
            userNameField.SendKeys(_userName);
            userPasswordField.SendKeys(_password);
            loginButton.Click();
        }

        /// <summary>
        ///     Log out of Mint.
        /// </summary>
        public void Logout() {
            var logoutLink = _driver.FindElementById(_logoutLinkId);
            logoutLink.Click();
        }
    }
}