namespace EPA_Grant_Scraper.Classes
{
    /// <summary>
    /// Given a path on where the log file should reside, it will create or add to a text file failed downloads.
    /// </summary>
    /// <param name="logFilePath"></param>
    public class FileDownloadLogger(string logFilePath)
    {
        private readonly string logFilePath = logFilePath;

        public void LogFailedDownload(string url, string fileName)
        {
            string logEntry = $"{url}\t{fileName}";

            try
            {
                using (StreamWriter writer = new (logFilePath, true))
                {
                    writer.WriteLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                ConsoleColorChanger.WriteMessageInRedWithWhiteBackground($"Failed to log the download attempt: {ex.Message}");
            }
        }
    }
}
