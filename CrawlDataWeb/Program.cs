using System;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace SeleniumTest
{
    public class Program
    {
        static void Main(string[] args)
        {
            var program = new Program();
            var listUrl = program.GetUrlFromExcel();
            if (listUrl.Count > 0)
            {
                foreach(var url in listUrl)
                {
                    program.CrawlData(url);
                }

            }
        }

        public void CrawlData(string? url)
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--headless");
            //options.AddArgument("--disable-gpu");
            //options.AddArgument("--disable-extensions");
            //options.AddArgument("--disable-popup-blocking");
            //options.AddArgument("--disable-infobars");
            IWebDriver driver = new ChromeDriver(options);
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(url);
            bool checkPage = true;
            var page = 0;
            while (page < 3)
            {
                try
                {
                    string currentUrl = driver.Url;
                    page++;

                    Thread.Sleep(4000);
                    if (page == 1)
                    {
                        var button1 = driver.FindElement(By.ClassName("PopupCloseControl__PopupCloseControlContainer-sc-1ge5cvc-0"));
                        if (button1 != null) button1.Click();
                        Thread.Sleep(1000);
                    }
                    var listProduct = driver.FindElements(By.ClassName("product"));
                    var listproduct = new List<ProductInfo>();
                    if (listProduct.Count > 0)
                    {
                        for (var i = 0; i < listProduct.Count; i++)
                        {
                            // Get product info
                            var productInfoDetail = driver.FindElement(By.CssSelector("#product-listing-container > form > ul > li:nth-child(" + (i + 1) + ") > article > div > figure > a")).GetAttribute("data-analytics-sent");
                            var productInfo = new ProductInfo();
                            productInfo = JsonConvert.DeserializeObject<ProductInfo>(productInfoDetail);
                            productInfo!.imageSrc = driver.FindElement(By.CssSelector("#product-listing-container > form > ul > li:nth-child(" + (i + 1) + ") > article > div > figure > a > div > span > img")).GetAttribute("src");
                            productInfo.pathDetail = driver.FindElement(By.CssSelector("#product-listing-container > form > ul > li:nth-child(" + (i + 1) + ") > article > div > figure > a")).GetAttribute("href");
                            listproduct.Add(productInfo);
                        }
                    };
                    for (var j = 0; j < listproduct.Count; j++)
                    {
                        // Get more info detail product
                        GetInfoDetailProduct(driver, listproduct[j]);
                    }
                    Thread.Sleep(2000);

                    driver.Navigate().GoToUrl(currentUrl);
                    Thread.Sleep(2000);
                    var buttonNext = FindElement(driver, "#product-listing-container > div.pagination > ul > li.pagination-item.pagination-item--next");
                    if (buttonNext != null)
                    {
                        buttonNext.Click(); // next page
                    }
                    else
                    {
                        checkPage = false; // If page is last page => check =  false (Stop Crawl)
                    };
                }
                catch (Exception ex)
                {
                    checkPage = false; // If page is last page => check =  false (Stop Crawl)
                }
            }
            driver.Quit();
        }

        public List<string?> GetUrlFromExcel()
        {
            var listUrl = new List<string>();
            string filePath = "E:/CrawlDataWeb/CrawlDataWeb/CrawlDataWeb/FolderExcel/example.xlsx";
            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                if (sheetData != null)
                {
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            if (cell.CellReference == "B2")
                            {
                                string linkString = GetCellValue(cell, workbookPart);
                                listUrl.Add(linkString);
                            }
                        }
                    }
                }
            }
            return listUrl;
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            string cellValue = "";

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(cell.InnerText);
                SharedStringItem sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index);
                cellValue = sharedStringItem.InnerText;
            }
            else
            {
                cellValue = cell.InnerText;
            }

            return cellValue;
        }

        public ProductInfo GetInfoDetailProduct(IWebDriver driver, ProductInfo? productDetail)
        {
            driver.Navigate().GoToUrl(productDetail.pathDetail);
            //var link = productDetail.pathDetail;
            //string html = driver.PageSource;

            //HtmlWeb doc = new HtmlWeb();
            //var htmlDoc = doc.Load(link);
            //HtmlNode someNode = htmlDoc.DocumentNode.SelectSingleNode("//h1[@class='productView-title']");
           // var text = someNode;
            var descriptionHeader = FindElement(driver, "body > div.body > div.container > div > div:nth-child(1) > div > div:nth-child(3) > section:nth-child(1) > div > p:nth-child(1) > span");
            productDetail.descriptionHeader = descriptionHeader != null ? descriptionHeader.Text : null;
            productDetail.descriptionDetail = new List<string>();
            var listDescriptionDetail = driver.FindElements(By.CssSelector("body > div.body > div.container > div > div:nth-child(1) > div > div:nth-child(3) > section:nth-child(1) > div > li"));
            if (listDescriptionDetail.Count != 0)
            {
                for (var j = 1; j <= listDescriptionDetail.Count; j++)
                {
                    var descriptionDetail = FindElement(driver, "body > div.body > div.container > div > div:nth-child(1) > div > div:nth-child(3) > section:nth-child(1) > div > li:nth-child(" + j + ")");
                    if (descriptionDetail != null)
                    {
                        productDetail.descriptionDetail.Add(descriptionDetail.Text);
                    }
                }
            }
            else
            {
                listDescriptionDetail = driver.FindElements(By.CssSelector("body > div.body > div.container > div > div:nth-child(1) > div > div:nth-child(3) > section:nth-child(1) > div > ul > li"));
                if (listDescriptionDetail.Count != 0)
                {
                    for (var j = 1; j <= listDescriptionDetail.Count; j++)
                    {
                        var descriptionDetail = FindElement(driver, "body > div.body > div.container > div > div:nth-child(1) > div > div:nth-child(3) > section:nth-child(1) > div > ul > li:nth-child(" + j + ") > span > strong");
                        if (descriptionDetail != null)
                        {
                            productDetail.descriptionDetail.Add(descriptionDetail.Text);
                        }
                    }
                }
            }
            Console.Write(JsonConvert.SerializeObject(productDetail).ToString());
            return productDetail;
        }

        public IWebElement? FindElement(IWebDriver driver, string? cssSelecter)
        {
            IWebElement elemement = null;
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
            try
            {
                return elemement = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(cssSelecter)));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return elemement;
            }
        }

        public class ProductInfo
        {
            public int product_id { get; set; }
            public string? name { get; set; }
            public string? category { get; set; }
            public string? brand { get; set; }
            public decimal? price { get; set; }
            public string? currency { get; set; }
            public string? imageSrc { get; set; }
            public string? pathDetail { get; set; }
            public string? descriptionHeader { get; set; }
            public List<string>? descriptionDetail { get; set; }
        }
    }
}