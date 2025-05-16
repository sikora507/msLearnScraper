using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace MsLearnScraper;

using OpenQA.Selenium.Chrome;
using HtmlAgilityPack;
using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private readonly IConfiguration _configuration;

    public Worker(ILogger<Worker> logger, IConfiguration configuration)
    {
        _logger = logger;
        _configuration = configuration;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var baseUrl = _configuration["BaseUrl"]!;
        _logger.LogInformation("Worker started at: {time}", DateTimeOffset.Now);
        _logger.LogInformation("Base URL: {baseUrl}", baseUrl);

        try
        {
            // Configure Chrome options
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--headless");
            chromeOptions.AddArgument("--disable-gpu");
            chromeOptions.AddArgument("--no-sandbox");

            // Initialize Chrome Driver
            using var driver = new ChromeDriver(chromeOptions);
            // Navigate to the URL
            await driver.Navigate().GoToUrlAsync(baseUrl);

            // Wait for the specific element to be loaded
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(d => d.FindElement(By.CssSelector("ul.tree.table-of-contents")));

            // Get the Blazor menu section
            var blazorMenuSection = GetBlazorMenuSection(driver, baseUrl);

            // Expand all tree items in the menu
            await ExpandAllTreeItems(driver, wait, blazorMenuSection);

            // Get the fully expanded menu element
            var expandedBlazorMenu = GetBlazorMenuSection(driver, baseUrl);

            // Get the outerHTML of the fully expanded menu
            var expandedMenuHtml = expandedBlazorMenu.GetAttribute("outerHTML");

            // Print the expanded menu HTML to the console
            _logger.LogInformation("Fully expanded menu structure:");
            _logger.LogInformation(expandedMenuHtml);

            // You can also save the menu structure to a file if needed
            await File.WriteAllTextAsync("expanded-menu.html", expandedMenuHtml);
            _logger.LogInformation("Expanded menu saved to expanded-menu.html");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error occurred while fetching or parsing HTML");
        }
    }

    private IWebElement GetBlazorMenuSection(ChromeDriver driver, string baseUrl)
    {
        // Find all ul elements in the table of contents
        var ulElements = driver.FindElements(By.CssSelector("ul.tree.table-of-contents ul"));

        // Find the first ul where the first li has an a child with href starting with baseUrl
        foreach (var ul in ulElements)
        {
            try
            {
                var firstLi = ul.FindElement(By.CssSelector("li"));
                var anchor = firstLi.FindElement(By.CssSelector("a"));
                var href = anchor.GetAttribute("href");

                if (href.StartsWith(baseUrl))
                {
                    return ul;
                }
            }
            catch (NoSuchElementException)
            {
                // Element not found, continue to next ul
                continue;
            }
        }

        _logger.LogError("No matching table of contents found");
        throw new InvalidOperationException("No matching table of contents found");
    }

    private async Task ExpandAllTreeItems(ChromeDriver driver, WebDriverWait wait, IWebElement parentElement)
    {
        // Find all non-expanded tree items
        var nonExpandedItems = parentElement.FindElements(
            By.CssSelector("li.tree-item:not(.is-expanded)"));

        _logger.LogInformation("Found {count} non-expanded items", nonExpandedItems.Count);

        foreach (var item in nonExpandedItems)
        {
            try
            {
                // Click to expand
                _logger.LogInformation("Expanding an item");
                item.Click();

                // Wait for the item to be expanded
                wait.Until(d =>
                {
                    var refreshedItem = driver.FindElement(By.Id(item.GetAttribute("id")));
                    return refreshedItem.GetAttribute("class").Contains("is-expanded");
                });

                // Get the refreshed element
                var refreshedItem = driver.FindElement(By.Id(item.GetAttribute("id")));

                // Wait for child ul to be loaded
                wait.Until(d => refreshedItem.FindElements(By.CssSelector("ul")).Count > 0);

                // Recursively expand the newly loaded items
                var childUl = refreshedItem.FindElement(By.CssSelector("ul"));
                await ExpandAllTreeItems(driver, wait, childUl);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error expanding tree item");
            }

            // Small delay to prevent overwhelming the browser
            await Task.Delay(200);
        }
    }
}