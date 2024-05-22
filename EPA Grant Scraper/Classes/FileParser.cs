using ClosedXML.Excel;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace EPA_Grant_Scraper.Classes
{
    public static partial class FileParser
    {
        public static async void ProcessFile(string filePath)
        {
            Console.WriteLine("\nChecking given file path...");

            if (!IsValidExcelFile(filePath))
            {
                ConsoleColorChanger.WriteMessageInRedWithWhiteBackground("" +
                    "The file given is not able to be parsed.\n" +
                    "Make sure the file path you gave is correct.\n" +
                    "\nRestarting the application.");

                RestartApplication();
            }

            Console.WriteLine("File path appears to be good!");


            string downloadPath = UserPrompts.GetPathOfDownloadLocation();

            List<string> downloadDirectories = CreateDirectories(downloadPath);
            string parentDirPath = downloadDirectories[0];
            string selectedDirPath = downloadDirectories[1];
            string notSelectedDirPath = downloadDirectories[2];

            using (XLWorkbook wb = new(filePath))
            {
                await ParseExcelFile(wb, selectedDirPath, notSelectedDirPath, parentDirPath);
            }
        }

        private static List<string> CreateDirectories(string downloadPath)
        {
            Console.WriteLine("\nCreating directories...");

            string year = DateTime.Now.Year.ToString();
            string parentDirName = $"{year}_EPAGrantSelections";
            string selectedDirName = "Selected";
            string notSelectedDirName = "Not-selected";

            string parentDirPath = Path.Combine(downloadPath, parentDirName);
            string selectedDirPath = Path.Combine(parentDirPath, selectedDirName);
            string notSelectedDirPath = Path.Combine(parentDirPath, notSelectedDirName);

            Directory.CreateDirectory(parentDirPath);
            Directory.CreateDirectory(selectedDirPath);
            Directory.CreateDirectory(notSelectedDirPath);

            Console.WriteLine("\nDirectories created!");

            return [parentDirPath, selectedDirPath, notSelectedDirPath];
        }

        private async static Task ParseExcelFile(XLWorkbook wb, string selectedPath, string notSelectedPath, string parentDirPath)
        {
            Console.WriteLine("\nProcessing Excel file...");

            try
            {
                // Assuming data is in the first worksheet
                IXLWorksheet worksheet = wb.Worksheet(1);

                // Loop through each row in the worksheet, starting from the second row
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    EPAGranteeFile granteeFile = new()
                    {
                        Region = row.Cell(1).GetString(),
                        State = row.Cell(2).GetString(),
                        ApplicantName = row.Cell(3).GetString(),
                        TypeOfApplication = row.Cell(4).GetString(),
                        FundingStatus = row.Cell(5).GetString().Equals("Selected", StringComparison.OrdinalIgnoreCase),
                        BipartisanInfrastructureLawFunds = row.Cell(6).GetString().Equals("Yes", StringComparison.OrdinalIgnoreCase)
                            ? true
                            : (row.Cell(6).GetString().Equals("No", StringComparison.OrdinalIgnoreCase)
                                ? false
                                : null)
                    };

                    // Column D (1-based index)
                    IXLCell urlCell = row.Cell(4);

                    if (urlCell.HasHyperlink)
                    {
                        granteeFile.FilePath = urlCell.GetHyperlink().ExternalAddress.ToString();

                        if (granteeFile.FundingStatus == true)
                        {
                            await DownloadFileAsync(granteeFile.FilePath, selectedPath, granteeFile, parentDirPath);
                        }
                        else
                        {
                            await DownloadFileAsync(granteeFile.FilePath, notSelectedPath, granteeFile, parentDirPath);
                        }
                    }

                    await DelayDownload();
                }
            }
            catch (Exception ex)
            {
                ConsoleColorChanger.WriteMessageInRedWithWhiteBackground($"\nError processing the Excel file: {ex.Message}\n");
            }

            Console.WriteLine("\nProcessing of Excel file completed!");
        }

        /// <summary>
        /// Wait for a random delay between 1 and 5 seconds.
        /// Simple method for avoiding detection as a bot while mass downloading files.
        /// </summary>
        private static async Task DelayDownload()
        {
            Random random = new();
            int delay = random.Next(1000, 2000);
            await Task.Delay(delay);
        }

        private static async Task DownloadFileAsync(string url, string downloadFolderPath, EPAGranteeFile epaGranteeFile, string parentDirPath)
        {
            string year = DateTime.Now.Year.ToString();
            string fileName = $"{year}_{SanitizeString(epaGranteeFile.State)}_{SanitizeString(epaGranteeFile.ApplicantName)}.pdf";

            try
            {
                using (HttpClient client = new())
                {
                    // Set headers to mimic a browser request
                    client.DefaultRequestHeaders.Add("authority", "www.epa.gov");
                    client.DefaultRequestHeaders.Add("method", "GET");
                    client.DefaultRequestHeaders.Add("path", $"{url.Replace("https://www.epa.gov", "")}");
                    client.DefaultRequestHeaders.Add("scheme", "https");
                    client.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8");
                    client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br, zstd");
                    client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.9");
                    client.DefaultRequestHeaders.Add("Priority", "u=0, i");
                    client.DefaultRequestHeaders.Add("Referer", "https://www.epa.gov/brownfields");
                    client.DefaultRequestHeaders.Add("Sec-Ch-Ua", "\"Brave\"; v = \"125\", \"Chromium\"; v = \"125\", \"Not.A/Brand\"; v = \"24\"");
                    client.DefaultRequestHeaders.Add("Sec-Ch-Ua-Mobile", "?0");
                    client.DefaultRequestHeaders.Add("Sec-Ch-Ua-Platform", "Windows");
                    client.DefaultRequestHeaders.Add("Sec-Fetch-Dest", "document");
                    client.DefaultRequestHeaders.Add("Sec-Fetch-Mode", "navigate");
                    client.DefaultRequestHeaders.Add("Sec-Fetch-Site", "same-origin");
                    client.DefaultRequestHeaders.Add("Sec-Fetch-User", "?1");
                    client.DefaultRequestHeaders.Add("Sec-Gpc", "1");
                    client.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36");

                    HttpResponseMessage response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string filePath = Path.Combine(downloadFolderPath, fileName);

                    using (FileStream fileStream = new(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }

                    ConsoleColorChanger.WriteMessageInGreen($"Downloaded: {fileName}");
                }
            }
            catch (Exception ex)
            {
                ConsoleColorChanger.WriteMessageInRedWithWhiteBackground($"\nError downloading:\n\tURL => {url},\n\tFileName => {fileName}\n\tError Message: {ex.Message}\n");
                FileDownloadLogger fileDownloadLogger = new (parentDirPath);
                fileDownloadLogger.LogFailedDownload(url, fileName);
            }
        }

        /// <summary>
        /// <list type="number">
        /// <item>Remove non-alphanumeric characters except spaces</item>
        /// <item>Reduce multiple spaces to one</item>
        /// <item>Trim leading and trailing whitespace</item>
        /// <item>Replace spaces with hyphens</item>
        /// </list>
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private static string SanitizeString(string input)
        {
            string sanitized = NonAlphanumericOrWhitespace().Replace(input, "");
            sanitized = ConsecutiveWhitespace().Replace(sanitized, " ");
            sanitized = sanitized.Trim();
            return sanitized.Replace(" ", "-");
        }

        private static bool IsValidExcelFile(string path)
        {
            Console.WriteLine("\nVerifying file contents...");

            try
            {
                using (XLWorkbook workbook = new(path))
                {
                    // Assuming the data is in the first worksheet
                    IXLWorksheet worksheet = workbook.Worksheet(1);
                    string[] expectedHeaders =
                    [
                        "Region", "State", "Applicant Name", "Type of Application", "Funding Status", "Bipartisan Infrastructure Law Funds"
                    ];

                    // Read the first row (headers)
                    var row = worksheet.Row(1);

                    // Check each expected header
                    for (int i = 0; i < expectedHeaders.Length; i++)
                    {
                        // Cells are 1-based
                        string cellValue = row.Cell(i + 1).GetString().Trim();

                        if (!cellValue.Equals(expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                        {
                            return false;
                        }
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                ConsoleColorChanger.WriteMessageInRedWithWhiteBackground($"Error reading the Excel file: {ex.Message}");
                return false;
            }
        }

        private static void RestartApplication()
        {
            string fileName = Environment.ProcessPath;
            Process.Start(fileName);
            Environment.Exit(0);
        }

        [GeneratedRegex(@"[^a-zA-Z0-9\s]")]
        private static partial Regex NonAlphanumericOrWhitespace();

        [GeneratedRegex(@"\s+")]
        private static partial Regex ConsecutiveWhitespace();
    }
}
