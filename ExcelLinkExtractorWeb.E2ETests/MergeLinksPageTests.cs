using System.Text.RegularExpressions;
using Microsoft.Playwright;
using NUnit.Framework;

namespace ExcelLinkExtractorWeb.E2ETests;

[TestFixture]
public class MergeLinksPageTests : SheetLinkPageTest
{
    [Test]
    public async Task MergePage_ShouldLoadSuccessfully()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Check page title
        await Expect(Page).ToHaveTitleAsync(new Regex("SheetLink"));

        // Check main heading
        var heading = Page.Locator("h1");
        await Expect(heading).ToContainTextAsync("Link Merger");
    }

    [Test]
    public async Task MergePage_ShouldShowDescription()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Check description text - match actual content
        var description = Page.Locator("text=Combine Title and URL into clickable hyperlinks");
        await Expect(description).ToBeVisibleAsync();
    }

    [Test]
    public async Task DownloadMergeTemplateButton_ShouldBeVisible()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Check download template button
        var downloadButton = Page.Locator("button:has-text('Download Sample')");
        await Expect(downloadButton).ToBeVisibleAsync();
    }

    [Test]
    public async Task NavigationBetweenPages_ShouldWork()
    {
        // Start on home page
        await Page.GotoAsync(BaseUrl);
        await WaitForHomeInteractiveAsync();

        // Wait for h1 to be visible (this is always visible, not affected by isInteractive)
        var homeHeading = Page.Locator("h1").Filter(new() { HasText = "SheetLink" });
        await Expect(homeHeading).ToBeVisibleAsync(new() { Timeout = 10000 });

        // Navigate to Merge page
        var mergeLink = Page.Locator(".navbar-nav-mobile").GetByRole(AriaRole.Link, new() { Name = "Merge" });
        await mergeLink.ClickAsync();
        await WaitForMergeInteractiveAsync();

        // Wait for merge page heading
        var mergeHeading = Page.Locator("h1").Filter(new() { HasText = "Link Merger" });
        await Expect(mergeHeading).ToBeVisibleAsync(new() { Timeout = 10000 });

        // Navigate back to Extract page
        var extractLink = Page.Locator(".navbar-nav-mobile").GetByRole(AriaRole.Link, new() { Name = "Extract" });
        await extractLink.ClickAsync();
        await WaitForHomeInteractiveAsync();

        // Check we're back on home page
        await Expect(homeHeading).ToBeVisibleAsync(new() { Timeout = 10000 });
    }

    [Test]
    public async Task MergePage_ShouldShowFileInput()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Check file input exists - wait for it to appear
        var fileInput = Page.Locator("#mergeFileInput");
        await Expect(fileInput).ToBeVisibleAsync();
    }

    [Test]
    public async Task MergePage_ShouldShowMergeButton()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Now the merge button should be visible
        var mergeButton = Page.Locator("button:has-text('Upload & Merge')");
        await Expect(mergeButton).ToBeVisibleAsync();
    }

    [Test]
    public async Task MergePage_ErrorMessage_ShouldShowForInvalidFile()
    {
        await Page.GotoAsync($"{BaseUrl}/merge");
        await WaitForMergeInteractiveAsync();

        // Create a temporary text file (not Excel)
        var tempFile = Path.GetTempFileName();
        await File.WriteAllTextAsync(tempFile, "This is not an Excel file");

        try
        {
            // Upload the invalid file
            var fileInput = Page.Locator("input[type='file']");
            await fileInput.SetInputFilesAsync(tempFile);

            // Submit
            var mergeButton = Page.Locator("button:has-text('Upload & Merge')");
            await mergeButton.ClickAsync();

            // Wait for error message
            await Page.WaitForTimeoutAsync(3000);

            // Check for error alert
            var errorAlert = Page.Locator(".alert-danger");
            await Expect(errorAlert).ToBeVisibleAsync();
        }
        finally
        {
            File.Delete(tempFile);
        }
    }
}
