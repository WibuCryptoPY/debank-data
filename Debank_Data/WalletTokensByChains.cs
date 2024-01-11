namespace Debank_Data
{
    struct Token
    {
        public string Name { get; set; }
        public string Price { get; set; }
        public string Amount { get; set; }
        public string USDValue { get; set; }
        public override string ToString()
        {
            return Name + "\t" + Price + "\t" + Amount + "\t" + USDValue;
        }
    }
    internal class WalletTokensByChains
    {
        public string Chain { get; set; }
        public string TotalValue { get; set; }
        public string? WalletValue { get; set; }
        public List<Token>? Tokens { get; set; }
        public List<ProjectPortofilo> projects { get; set; }
        public WalletTokensByChains()
        {
            Chain = "";
            TotalValue= "0";
            Tokens = new List<Token>();
            projects = new List<ProjectPortofilo>();
        }
        public override string ToString()
        {
            string result = Chain + "\t" + TotalValue + "\n";
            foreach (var token in Tokens)
            {
                result += token.ToString() + "\n";
            }
            foreach (var project in projects)
            {
                result += project.ToString() + "\n";
            }
            return result;
        }

    }
}
