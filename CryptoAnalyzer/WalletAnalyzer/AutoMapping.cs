using AutoMapper;
using WebScraper.WebScrapers;

namespace WalletAnalyzer
{
    public class AutoMapping : Profile
    {
        public AutoMapping()
        {
            CreateMap<DexRow, DexRowOutputDto>();
            CreateMap<DexTable, DexTableOutputDto>();
        }
    }
}
