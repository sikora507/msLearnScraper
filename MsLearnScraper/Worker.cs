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
        var outputBasePath = _configuration["OutputBasePath"]!;
        _logger.LogInformation("Worker started at: {time}", DateTimeOffset.Now);
        _logger.LogInformation("Base URL: {baseUrl}", baseUrl);
        _logger.LogInformation("Output Base Path: {outputBasePath}", outputBasePath);

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

            // Clear all attributes in ul and li elements, leaving only href attributes in a tags
            var cleanedHtml = CleanHtmlAttributes(expandedMenuHtml!);

            // Create the index.html file in the output directory
            var pagesDirectory = Path.Combine(outputBasePath, "pages");
            Directory.CreateDirectory(pagesDirectory); // Creates if doesn't exist

            await File.WriteAllTextAsync(Path.Combine(outputBasePath, "index.html"), cleanedHtml);
            _logger.LogInformation("Expanded menu saved to {path}", Path.Combine(outputBasePath, "index.html"));

            // Extract and download all pages
            await DownloadAllPages(driver, wait, cleanedHtml, pagesDirectory, baseUrl);
            
            // Update the index.html links to point to local files
            await UpdateIndexLinks(Path.Combine(outputBasePath, "index.html"), baseUrl);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error occurred while fetching or parsing HTML");
        }
    }

    private async Task UpdateIndexLinks(string indexFilePath, string baseUrl)
    {
        try
        {
            _logger.LogInformation("Updating links in index.html to point to local files");
        
            var indexHtml = await File.ReadAllTextAsync(indexFilePath);
            var doc = new HtmlDocument();
            doc.LoadHtml(indexHtml);
        
            var links = doc.DocumentNode.SelectNodes("//a[@href]");
            if (links == null || links.Count == 0)
            {
                _logger.LogWarning("No links found in the index HTML");
                return;
            }
        
            foreach (var link in links)
            {
                var href = link.GetAttributeValue("href", "");
                if (string.IsNullOrEmpty(href) || !href.StartsWith(baseUrl))
                {
                    continue;
                }
            
                // Generate the same filename as used in DownloadAllPages
                var fileName = GenerateFileName(href, link.InnerText.Trim());
            
                // Update the href to point to the local file
                link.SetAttributeValue("href", $"pages/{fileName}");
            }
        
            await File.WriteAllTextAsync(indexFilePath, doc.DocumentNode.OuterHtml);
            _logger.LogInformation("Successfully updated links in index.html");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error updating links in index.html: {error}", ex.Message);
        }
    }
    
    private async Task DownloadAllPages(ChromeDriver driver, WebDriverWait wait, string menuHtml, string pagesDirectory,
        string baseUrl)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(menuHtml);

        var links = doc.DocumentNode.SelectNodes("//a[@href]");
        if (links == null || links.Count == 0)
        {
            _logger.LogWarning("No links found in the menu HTML");
            return;
        }

        _logger.LogInformation("Found {count} links to download", links.Count);

        foreach (var link in links)
        {
            var href = link.GetAttributeValue("href", "");
            if (string.IsNullOrEmpty(href) || !href.StartsWith(baseUrl))
            {
                continue;
            }

            try
            {
                // Generate a meaningful filename from the URL and link text
                var fileName = GenerateFileName(href, link.InnerText.Trim());
                var filePath = Path.Combine(pagesDirectory, fileName);

                _logger.LogInformation("Downloading {url} to {filePath}", href, filePath);

                // Navigate to the page
                await driver.Navigate().GoToUrlAsync(href);

                // Wait for the main content to load
                wait.Until(d => d.FindElement(By.CssSelector("main")));

                // Get the main content
                var mainContent = GetVisibleContent(driver);

                // Clean up the content according to requirements
                var cleanedContent = CleanPageContent(mainContent);

                // Save the content to a file
                await File.WriteAllTextAsync(filePath, cleanedContent);

                _logger.LogInformation("Successfully saved {url} to {filePath}", href, filePath);

                // Small delay to avoid overwhelming the server
                await Task.Delay(300);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading page {url}", href);
            }
        }

        _logger.LogInformation("Finished downloading all pages");
    }

    private string CleanPageContent(string htmlContent)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        // 1. Remove specific elements by ID
        RemoveElementById(doc, "article-header");
        RemoveElementById(doc, "article-metadata");
        RemoveElementById(doc, "assertive-live-region");
        RemoveElementById(doc, "polite-live-region");
        RemoveElementById(doc, "ms--additional-resources-mobile");
        RemoveElementById(doc, "ms--inline-notifications");

        // 3. Remove feedback section
        RemoveElementsByCssSelector(doc, "section.feedback-section");

        // 4. Remove all button elements - use tag name instead of selector
        RemoveElementsByCssSelector(doc, "button");
        RemoveElementsByCssSelector(doc, "form");

        // 5. Remove all attributes from all elements (except we won't keep href here)
        RemoveAllAttributes(doc);
        
        // 6. Add "Back to menu" link at the end of the content
        AddBackToMenuLink(doc);

        return doc.DocumentNode.OuterHtml;
    }
    
    private void AddBackToMenuLink(HtmlDocument doc)
    {
        // Create a new div element with a "Back to menu" link
        var backToMenuDiv = doc.CreateElement("div");
    
        var backToMenuLink = doc.CreateElement("a");
        backToMenuLink.SetAttributeValue("href", "../index.html");
        backToMenuLink.InnerHtml = "Back to menu";
    
        backToMenuDiv.AppendChild(backToMenuLink);
    
        // Append the div to the main content
        var mainElement = doc.DocumentNode.SelectSingleNode("//main");
        if (mainElement != null)
        {
            mainElement.AppendChild(backToMenuDiv);
        }
        else
        {
            // If there's no main element, add it to the body or to the document
            var body = doc.DocumentNode.SelectSingleNode("//body");
            if (body != null)
            {
                body.AppendChild(backToMenuDiv);
            }
            else
            {
                doc.DocumentNode.AppendChild(backToMenuDiv);
            }
        }
    }

    private void RemoveElementById(HtmlDocument doc, string id)
    {
        var element = doc.DocumentNode.SelectSingleNode($"//*[@id='{id}']");
        element?.Remove();
    }

    private void RemoveElementsByCssSelector(HtmlDocument doc, string cssSelector)
    {
        var elements = doc.DocumentNode.QuerySelectorAll(cssSelector);
        if (elements != null)
        {
            foreach (var element in elements.ToList())
            {
                element.Remove();
            }
        }
    }

    private void RemoveAllAttributes(HtmlDocument doc)
    {
        var allElements = doc.DocumentNode.SelectNodes("//*");
        if (allElements != null)
        {
            foreach (var element in allElements)
            {
                // Remove all attributes for every element
                element.Attributes.RemoveAll();
            }
        }
    }

    private string GetVisibleContent(ChromeDriver driver)
    {
        // Execute JavaScript to get only the visible content
        string script = @"
        function isVisible(elem) {
            if (!elem) return false;
            if (window.getComputedStyle(elem).display === 'none') return false;
            if (window.getComputedStyle(elem).visibility === 'hidden') return false;
            if (elem.offsetWidth === 0 && elem.offsetHeight === 0) return false;
            return true;
        }

        function processNode(node) {
            if (node.nodeType === Node.ELEMENT_NODE) {
                // Skip if this element is not visible
                if (!isVisible(node)) {
                    return document.createDocumentFragment();
                }
                
                // Create a new element of the same type
                const newElem = document.createElement(node.tagName);
                
                // Copy attributes
                Array.from(node.attributes).forEach(attr => {
                    newElem.setAttribute(attr.name, attr.value);
                });
                
                // Recursively process child nodes
                Array.from(node.childNodes).forEach(child => {
                    const processedChild = processNode(child);
                    if (processedChild) {
                        newElem.appendChild(processedChild);
                    }
                });
                
                return newElem;
            } else if (node.nodeType === Node.TEXT_NODE) {
                return document.createTextNode(node.textContent);
            } else {
                return document.createDocumentFragment();
            }
        }

        const mainContent = document.querySelector('main');
        if (!mainContent) return '<div>Content not found</div>';
        
        const visibleContent = processNode(mainContent);
        return visibleContent.outerHTML;
    ";

        // Execute JavaScript and get the result
        var jsExecutor = (IJavaScriptExecutor)driver;
        var result = jsExecutor.ExecuteScript(script);

        return result?.ToString() ?? "<div>Failed to extract content</div>";
    }

    private string GenerateFileName(string url, string linkText)
    {
        // Clean the link text to create a meaningful filename
        var fileName = string.IsNullOrWhiteSpace(linkText) ? "page" : linkText;

        // Replace invalid characters with underscores
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(c, '_');
        }

        // Replace spaces with underscores and remove other potentially problematic characters
        fileName = fileName.Replace(' ', '_')
            .Replace(':', '_')
            .Replace('/', '_')
            .Replace('\\', '_')
            .Replace('?', '_')
            .Replace('&', '_')
            .Replace('%', '_');

        // Remove leading dots
        fileName = fileName.TrimStart('.');

        // Ensure filename isn't too long
        if (fileName.Length > 50)
        {
            fileName = fileName.Substring(0, 50);
        }

        // Extract the last part of the URL path for uniqueness
        var uri = new Uri(url);
        var pathParts = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        var lastPathPart = pathParts.Length > 0 ? pathParts[pathParts.Length - 1] : "index";

        // Combine filename with the last path part for uniqueness
        fileName = $"{fileName}_{lastPathPart}.html";

        return fileName;
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

    private string CleanHtmlAttributes(string html)
    {
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var baseUrl = _configuration["BaseUrl"]!;

        // First, remove li elements with direct children a elements that don't start with baseUrl
        var liElements = doc.DocumentNode.SelectNodes("//li[a]");
        if (liElements != null)
        {
            foreach (var li in liElements.ToList()) // ToList to avoid collection modification issues
            {
                var anchor = li.SelectSingleNode("./a");
                if (anchor != null)
                {
                    var href = anchor.GetAttributeValue("href", "");
                    if (!href.StartsWith(baseUrl))
                    {
                        li.Remove();
                    }
                }
            }
        }

        var allElements = doc.DocumentNode.SelectNodes("//*");
        if (allElements != null)
        {
            foreach (var element in allElements)
            {
                // For 'a' elements, create a new attribute collection containing only 'href'
                if (element.Name.ToLowerInvariant() == "a")
                {
                    var hrefAttribute = element.Attributes["href"];
                    element.Attributes.RemoveAll();
                    if (hrefAttribute != null)
                    {
                        var urlToReplace = hrefAttribute.Value.Contains("?")
                            ? hrefAttribute.Value + "&pivots=cli"
                            : hrefAttribute.Value + "?pivots=cli";

                        element.Attributes.Add("href", urlToReplace);
                    }
                }
                // For other elements, simply remove all attributes
                else
                {
                    element.Attributes.RemoveAll();
                }
            }
        }

        return doc.DocumentNode.OuterHtml;
    }
}