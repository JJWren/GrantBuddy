namespace EPA_Grant_Scraper.Classes
{
    public class EPAGranteeFile
    {
        public string Region {  get; set; }
        public string State { get; set; }
        public string ApplicantName { get; set; }
        public string TypeOfApplication { get; set; }
        public string FilePath { get; set; }

        /// <summary>
        /// Selected vs Not-Selected.
        /// </summary>
        public bool FundingStatus { get; set; }

        /// <summary>
        /// Yes if <c>true</c>.
        /// </summary>
        public bool? BipartisanInfrastructureLawFunds { get; set; }
    }
}
