using System.Collections.Generic;

namespace WalletAnalyzer
{
    public class DexTableOutputDto
    {
        public string TokenName { get; set; }
        public List<DexRowOutputDto> Rows { get; set; }
    }
}
