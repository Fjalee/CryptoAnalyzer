using System.Collections.Generic;

namespace WalletAnalyzer
{
    public class DexTableCsvOutputDto
    {
        public string TokenName { get; set; }
        public List<DexRowCsvOutputDto> Rows { get; set; }
    }
}
