namespace WebScraper
{
    public class DexTableUrlOptions
    {
        public string Path { get; set; }
        public string TokenVarName { get; set; }
    }

    public class TokenToScrape
    {
        public string Hash { get; set; }
        public int RowsAmount { get; set; }
    }

    public class DexTableOptions
    {
        public DexTableUrlOptions Url { get; set; }
    }

    public class AppSettingsOptions
    {
        public OutputOptions Output { get; set; }
        public BlockchainsOptions Blockchains { get; set; }
    }

    public class BlockchainsOptions
    {
        public BlockchainExplorerWebsiteOptions Etherscan { get; set; }
    }

    public class BlockchainExplorerWebsiteOptions
    {
        public string DomainName { get; set; }
        public ApiOptions Api { get; set; }
        public DexTableOptions DexTable { get; set; }
        public int SleepTimeBetweenScrapesInMs { get; set; }
        public TokenToScrape[] TokensToScrape { get; set; }
    }   

    public class ApiOptions
    {
        public string Path { get; set; }
        public int CallsPerSecond { get; set; }
        public int TryAgainDelayInMs { get; set; }
        public string ApiKey { get; set; }
    }

    public class OutputOptions
    {
        public string Path { get; set; }
        public string LogPath { get; set; }
        public int AppendPeriodInSeconds { get; set; }
        public string DateFormatForXlsx { get; set; }
        public int IndexFirstRowDexTable { get; set; }
    }
}
