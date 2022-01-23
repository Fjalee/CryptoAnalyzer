using CsvHelper.Configuration;
using System.Globalization;

namespace WalletAnalyzer
{
    public class DexRowOutputDtoMap : ClassMap<DexRowOutputDto>
    {
        public DexRowOutputDtoMap()
        {
            //var numSpec = "0." + new string('#', 339);
            //var floatNumSpec = "F4";
            //var intNumSpec = "N4";
            AutoMap(CultureInfo.InvariantCulture);

            //Map(m => m.TotalInaccurateValue).Convert(x => x.Value.TotalInaccurateValue.ToString(intNumSpec));
            //Map(m => m.TotalValue).Convert(x => x.Value.TotalValue.ToString(floatNumSpec));
        }
    }
}
