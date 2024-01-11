// See https://aka.ms/new-console-template for more information
using Debank_Data;
using OpenQA.Selenium.Chrome;
using static Debank_Data.MainClass;


Console.WriteLine(GetChromeDriverPath());
ChromeOptions options = new ChromeOptions();
options.AddArgument("--headless");
ChromeDriver driver = new ChromeDriver(GetChromeDriverPath(), options);
List<Wallets> wallets = new List<Wallets>();
List<string> WalletAddresses;


var str = $"".Split("\r\n").ToList();


WalletAddresses = str;
Excute(ref driver, WalletAddresses, wallets);



