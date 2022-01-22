namespace WalletAnalyzer
{
    public interface IDexOutput
    {
        public void DoOutput(string outputName, string tokenHash, DexTableCsvOutputDto table, string timeElapsed, int nmRows);
    }
}
