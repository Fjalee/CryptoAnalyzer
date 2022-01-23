namespace WalletAnalyzer
{
    public interface IDexOutput
    {
        public void DoOutput(string outputName, string tokenHash, DexTableOutputDto table, string timeElapsed, int nmRows);
    }
}
