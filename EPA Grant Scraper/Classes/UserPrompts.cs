namespace EPA_Grant_Scraper.Classes
{
    public class UserPrompts
    {
        public static void WelcomeMessage()
        {
            ConsoleColorChanger.WriteMessageInCyan("" +
                "Welcome to Grant Buddy!\n" +
                "by Joshua Wren - 2024\n\n");
            Console.WriteLine("This application helps parse your\n" +
                "selected/non-selected grantees\n" +
                "file from the EPA into a folder\n" +
                "containing two folders. One folder\n" +
                "(Selected) will contain all of the\n" +
                "files of selected grant winners.\n" +
                "The other (Non-selected) will\n" +
                "contain the files of the\n" +
                "non-selected grantees.\n");
        }

        public static void WorkingTaskMessage()
        {
            ConsoleColorChanger.WriteMessageInRedWithWhiteBackground("" +
                "\nGrant Buddy is working on its task." +
                "\nIf you press any key, the application will close... even before the program finishes.\n");
        }

        public static string GetPathOfEPAFile()
        {
            string path = string.Empty;

            while (string.IsNullOrWhiteSpace(path))
            {
                path = ValidatePathString(PromptUserForFilePath());
            }

            return path;
        }

        public static string GetPathOfDownloadLocation()
        {
            string directory = string.Empty;

            while (string.IsNullOrWhiteSpace(directory))
            {
                directory = ValidateDirectoryString(PromptUserForDownloadPath());
            }

            return directory;
        }

        private static string PromptUserForDownloadPath()
        {
            ConsoleColorChanger.WriteMessageInYellow("" +
                "\nPlease provide the path for where you want to create the" +
                "\ndirectory containing selected and non-selected grantee files." +
                "\nIt should look something like:");
            ConsoleColorChanger.WriteMessageInCyan("\nC:\\Users\\username\\Desktop\\FolderOfGrantStuff");

            return Console.ReadLine().Trim();
        }

        private static string PromptUserForFilePath()
        {
            ConsoleColorChanger.WriteMessageInYellow("" +
                "\nPlease provide the filepath for where the EPA file is located." +
                "\nIt should look something like:");
            ConsoleColorChanger.WriteMessageInCyan("\nC:\\Users\\username\\Desktop\\All winning Grants - 2024.xlsx");

            return Console.ReadLine().Trim();
        }

        /// <summary>
        /// Checks if a string is null or whitespace or if it is not a filepath.
        /// </summary>
        /// <param name="path"></param>
        /// <returns>If valid, a valid filepath; Otherwise, <see cref="string.Empty"/>.</returns>
        private static string ValidatePathString(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                ConsoleColorChanger.WriteMessageInYellow("\nThat path was invalid. Please try again.\n");
                path = string.Empty;
            }

            return path;
        }

        /// <summary>
        /// Checks if a string is null or whitespace or if it is not a directory.
        /// </summary>
        /// <param name="directory"></param>
        /// <returns>If valid, a valid filepath; Otherwise, <see cref="string.Empty"/>.</returns>
        private static string ValidateDirectoryString(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                ConsoleColorChanger.WriteMessageInYellow("\nThat directory was invalid. Please try again.\n");
                directory = string.Empty;
            }

            return directory;
        }
    }
}
