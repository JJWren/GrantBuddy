using EPA_Grant_Scraper.Classes;

UserPrompts.WelcomeMessage();
string path = UserPrompts.GetPathOfEPAFile();
FileParser.ProcessFile(path);
UserPrompts.WorkingTaskMessage();
Console.ReadLine();