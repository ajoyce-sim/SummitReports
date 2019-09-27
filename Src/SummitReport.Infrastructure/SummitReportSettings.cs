
namespace SummitReports.Objects
{
	public class SummitReportSettings
    {
        /// <summary></summary>
		public string ConnectionString { get; set; } = "";
        /// <summary></summary>
		public string SchemaName { get; set; } = "";
        /// <summary></summary>
		public string Version { get; set; } = "";
		/// <summary></summary>
		public bool VerboseMessages { get; set; } = false;
        private static SummitReportSettings instance;
        
		private SummitReportSettings()
        {
			//_configuration = configuration;
        }
        public static SummitReportSettings Instance
        {
            get
            {

                if (instance == null)
                {
                    instance = new SummitReportSettings();
				}
                return instance;
            }
        }
    }
}
