using System;
using WebScraper.WebScrapers.EtherscanDex;

namespace WalletAnalyzer
{
    public class DexRowOutputDto
    {
        public string TxnHash { get; set; }
        public DateTime TxnDate { get; set; }
        public DexAction Action { get; set; }
        public string BuyerHash { get; set; }
        public string SellerHash { get; set; }
        public string LastSell { get; set; }
    }
}
