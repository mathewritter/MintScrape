using MintScrape.Core;
using MintScrape.Interfaces;
using Ninject.Modules;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;

namespace MintScrape.Infrastructure {
    public class IocModule : NinjectModule {
        public sealed override void Load() {
            AddBindings();
        }

        private void AddBindings() {
            Kernel.Bind<IAccount>().To<Account>();
            Kernel.Bind<IHtmlParserUtility>().To<HtmlParserUtility>();
            Kernel.Bind<IUserAccount>().To<UserAccount>();
            Kernel.Bind<IExcelMapper>().To<ExcelMapper>();
            Kernel.Bind<RemoteWebDriver>().To<ChromeDriver>();
        }
    }
}