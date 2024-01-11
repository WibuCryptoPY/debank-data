using ClosedXML.Excel;
using HtmlAgilityPack;
using OpenQA.Selenium.Chrome;

namespace Debank_Data
{
    static internal class MainClass
    {
        public static void AddDataToVar()
        {

        }
        public static string GetChromeDriverPath()
        {
            string rootPath = System.IO.Directory.GetCurrentDirectory();
            string chromeDriverPath = rootPath + "\\chromedriver.exe";
            return chromeDriverPath;
        }
        public static Dictionary<string, string> GetDataChainsValue(string html)
        {
            var dataChainValues = new Dictionary<string, string>();

            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@data-chain]");
            var nodesTotalValue = htmlDoc.DocumentNode.SelectNodes("//span[contains(@class, 'AssetsOnChain_usdValue__I1B7X')]");

            if (nodes != null)
            {
                for (int i = 0; i < nodes.Count; i++)
                {
                    if (nodes[i].Attributes["data-chain"] == null || nodesTotalValue[i] == null)
                    {
                        continue;
                    }
                    string dataChainValue = nodes[i].Attributes["data-chain"]?.Value;
                    Console.WriteLine(dataChainValue);
                    string usdValue = nodesTotalValue[i].InnerText;
                    Console.WriteLine(usdValue);
                    if (!string.IsNullOrEmpty(dataChainValue) && !string.IsNullOrEmpty(usdValue))
                    {
                        dataChainValues.Add(dataChainValue, usdValue);
                    }
                }
            }

            return dataChainValues;
        }
        public static List<Token> GetTokenData(string html)
        {
            var TokenData = new List<Token>();
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            var nodes = htmlDoc.DocumentNode.Descendants("div")
                         .Where(n => n.GetAttributeValue("class", "").Contains("db-table-row"))
                         .ToList();
            if (nodes == null || nodes.Count < 0 || nodes.Where(s => s.ChildNodes.Count < 0).ToList().Count < 0)
            {
                return TokenData;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                var childNodeName = nodes[i].ChildNodes[0].SelectNodes("//a[contains(@class, 'TokenWallet_detailLink__goYJR')]")[i].InnerText;
                var childNodePrice = nodes[i].ChildNodes[1].InnerText;
                var childNodeAmount = nodes[i].ChildNodes[2].InnerText;
                var childNodeUSDValue = nodes[i].ChildNodes[3].InnerText;
                Console.WriteLine(childNodeName + "\t" + childNodePrice + "\t" + childNodeAmount + "\t" + childNodeUSDValue);
                TokenData.Add(new Token
                {
                    Name = childNodeName,
                    Price = childNodePrice,
                    Amount = childNodeAmount,
                    USDValue = childNodeUSDValue
                });
            }
            // Console.WriteLine(childNodeName);
            return TokenData;
        }
        public static List<ProjectPortofilo> GetProjectPortofilos(string html)
        {
            var ProjectPortofilos = new List<ProjectPortofilo>();
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='Project_portfolioProject__LxOOt']");
            if (nodes == null)
            {
                return ProjectPortofilos;
            }
            foreach (var node in nodes)
            {
                var nameAndValue = node.ChildNodes
                    .Where(s => s.SelectSingleNode("//span[@class='ProjectTitle_protocolLink__4Yqn3']").InnerText != null)
                    .FirstOrDefault()
                    .InnerText;
                if(nameAndValue == null)
                {
                    continue;
                }
                var name = nameAndValue.Split("$")[0];
                var value = "$" + nameAndValue.Split("$")[1];

                var bookmark = node.ChildNodes[1].ChildNodes[0].ChildNodes[0].InnerText;
                
                List<string> requiredClasses = "Panel_container__Vltd1 Panel_newContainer__gMO0h".Split(' ').ToList();
                HtmlNode panelNodes = node.Descendants()
                            .Where(node1 => requiredClasses.All(requiredClass =>
                                         node1.GetAttributeValue("class", "").Split(' ').Contains(requiredClass)))
                            .ToList()[0];
                var listOfDictionaries = GetDataFromPanel(panelNodes);
                ProjectPortofilos.Add(new ProjectPortofilo
                {
                    Name = name,
                    TotalValue = value,
                    Bookmark = bookmark,
                    Items = listOfDictionaries
                });



            }
            return ProjectPortofilos;
        }
        public static List<Dictionary<string, string>> GetDataFromPanel(HtmlNode node)
        {
            var data = new List<Dictionary<string, string>>();
            var dataNodes = node.ChildNodes;
            foreach (var dataNode in dataNodes)
            {
                var dt = new Dictionary<string, string>();
                List<string> classes = new List<string>();
                if (String.IsNullOrEmpty(dataNode.InnerText) || dataNode.OuterHtml.Contains("More_container__a3e03") || dataNode.OuterHtml.Contains("flex_flexRow__y0UR2 More_line__EHJmK"))
                {
                    continue;
                }
                foreach (var childNode in dataNode.ChildNodes[0].ChildNodes)
                {
                    Console.WriteLine(childNode.InnerText + "\t");
                    dt.Add(childNode.InnerText, "");
                    classes.Add(childNode.InnerText);
                }

                for (int i = 0; i < dataNode.ChildNodes[1].ChildNodes[0].ChildNodes.Count; i++)
                {
                    Console.WriteLine(dataNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].InnerText + "\t");
                    dt[classes[i]] = dataNode.ChildNodes[1].ChildNodes[0].ChildNodes[i].InnerText;
                }
                Console.WriteLine("--------------------------------------------------");
                data.Add(dt);

            }
            return data;
        }
        public static void Exportxcel(ref int startRow, List<WalletTokensByChains> wallets, string sheetName)
        {
            string excelPath = "FormattedExcel.xlsx";

            using (var workbook = new XLWorkbook(excelPath))
            {
                while (workbook.Worksheets.Count > 0)
                {
                    workbook.Worksheets.Delete(workbook.Worksheets.First().Name);
                }


                if (workbook.Worksheets.Any(s => s.Name == sheetName.Substring(0, sheetName.Length - 30)))
                {
                    workbook.Worksheets.Delete(sheetName.Substring(0, sheetName.Length - 30));
                }
                var worksheet = workbook.AddWorksheet(sheetName.Substring(0, sheetName.Length - 30));
                worksheet.Columns().AdjustToContents(); // Chỉnh độ rộng cột tự động
                                                        // worksheet.Cells().Style.Alignment.SetWrapText(true); 
                                                        // Đặt tiêu đề cho các cột
                worksheet.Cell("A" + startRow.ToString()).Value = "Chain";
                worksheet.Cell("B" + startRow).Value = "Section";
                worksheet.Range($"C{startRow.ToString()}:F{startRow.ToString()}").Merge().Value = sheetName;
                foreach (WalletTokensByChains wallet in wallets)
                {
                    worksheet.Cell("C" + (startRow + 1).ToString()).Value = "Token";
                    worksheet.Cell("D" + (startRow + 1).ToString()).Value = "Price";
                    worksheet.Cell("E" + (startRow + 1).ToString()).Value = "Amount";
                    worksheet.Cell("F" + (startRow + 1).ToString()).Value = "Value";

                    startRow += 2;


                    // Format các ô tiêu đề

                    if (wallet.Tokens.Count > 0 && wallet.Tokens != null)
                    {
                        int count = 0;
                        for (int i = startRow; i < wallet.Tokens.Count + startRow; i++)
                        {
                            worksheet.Cell("C" + i.ToString()).Value = wallet.Tokens[count].Name;
                            worksheet.Cell("D" + i.ToString()).Value = wallet.Tokens[count].Price;
                            worksheet.Cell("E" + i.ToString()).Value = wallet.Tokens[count].Amount;
                            worksheet.Cell("F" + i.ToString()).Value = wallet.Tokens[count].USDValue;
                            count++;

                        }

                        startRow += wallet.Tokens.Count;

                        worksheet.Range($"B{startRow - wallet.Tokens.Count - 1}:B{startRow - 1}").Merge().Value = $"Wallet({wallet.WalletValue})";

                    }
                    //wallet

                    // Merge các hàng dựa trên format yêu cầu



                    //int countprj = 0;
                    for (int i = 0; i < wallet.projects.Count; i++)
                    {
                        // Giả sử rằng 'Bookmark' là giá trị cố định cho mỗi project
                        worksheet.Range($"C{startRow}:F{startRow}").Merge().Value = wallet.projects[i].Bookmark;

                        int projectStartRow = startRow + 1; // Lưu lại vị trí bắt đầu của project
                        startRow++;
                        foreach (var item in wallet.projects[i].Items)
                        {
                            foreach (var kvp in item)
                            {
                                worksheet.Cell($"C{startRow}").Value = kvp.Key;
                                worksheet.Cell($"D{startRow}").Value = kvp.Value;
                                startRow++;
                            }
                        }

                        // Merge cells cho tên project sau khi đã thêm tất cả các KeyValuePair
                        worksheet.Range($"B{projectStartRow}:B{startRow - 1}").Merge().Value = wallet.projects[i].Name + $"({wallet.projects[i].TotalValue}";
                    }

                    worksheet.Range($"A{startRow - wallet.projects.Sum(s => s.GetAmountOfKeyValue()) - 1 * wallet.projects.Count - wallet.Tokens.Count - 1}:A{startRow - 1}").Merge().Value = wallet.Chain.ToUpper() + $"({wallet.TotalValue})";

                    // Các ô sau sẽ có các giá trị tương tự dựa trên dữ liệu của bạn

                    // Lưu file
                }
                workbook.Save();

            }

            Console.WriteLine("Excel file created successfully.");
        }
        public static void Excute(ref ChromeDriver driver, List<string> str, List<Wallets> wallets)
        {
            foreach (string addr in str)
            {
                driver.Navigate().GoToUrl($"https://debank.com/profile/{addr}");
                Thread.Sleep(5000);
                string pageSource = driver.PageSource;
                Console.WriteLine(pageSource);
                Dictionary<string, string> dataChainValues = GetDataChainsValue(pageSource);
                foreach (KeyValuePair<string, string> kvp in dataChainValues)
                {
                    Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
                    driver.Navigate().GoToUrl($"https://debank.com/profile/{addr}?chain=" + kvp.Key);
                    Thread.Sleep(5000);
                    pageSource = driver.PageSource;
                    var s = GetTokenData(pageSource);
                    var projects = GetProjectPortofilos(pageSource);
                    var htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(pageSource);
                    var node = htmlDoc.DocumentNode.SelectNodes("//div[@class='Portfolio_defiItem__cVQM-']");
                    List<string> requiredClasses = "ProjectTitle_projectTitle__yC5VD TokenWallet_walletProjectTitle__6TZPs".Split(' ').ToList();
                    HtmlNode walletNodes = node.Descendants()
                            .Where(node1 => requiredClasses.All(requiredClass =>
                                         node1.GetAttributeValue("class", "").Split(' ').Contains(requiredClass)))
                            .ToList()[0];
                    var rawWalletValue = walletNodes.InnerText;
                    string nameWallet = rawWalletValue.Split("$")[0];
                    string valueWallet = "$" + rawWalletValue.Split("$")[1];
                    if (wallets.Where(w => w.Address == addr).FirstOrDefault() == null)
                    {
                        wallets.Add(new Wallets
                        {
                            Address = addr,
                            WalletTokensByChainsProp = new List<WalletTokensByChains>()
                        });
                    }
                    wallets.Where(w => w.Address == addr).FirstOrDefault().WalletTokensByChainsProp.Add(new WalletTokensByChains
                    {
                        Chain = kvp.Key,
                        TotalValue = kvp.Value,
                        Tokens = s,
                        projects = projects,
                        WalletValue = valueWallet
                    });
                }
            }

            foreach (var wallet in wallets)
            {
                if (wallet.WalletTokensByChainsProp == null)
                {
                    continue;
                }
                foreach (var walletTokensByChain in wallet.WalletTokensByChainsProp)
                {
                    Console.WriteLine(walletTokensByChain.ToString());
                }
            }

            foreach (var wallet in wallets)
            {
                int startRow = 1;
                if(wallet.WalletTokensByChainsProp != null)
                Exportxcel(ref startRow, wallet.WalletTokensByChainsProp, wallet.Address);
            }
        }
    }
}
