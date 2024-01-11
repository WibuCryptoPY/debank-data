namespace Debank_Data
{
    internal class ProjectPortofilo
    {
        public string? Name { get; set; }
        public string? TotalValue { get; set; }
        public string? Bookmark { get; set; }
        public List<Dictionary<string, string>> Items { get; set; }
        public ProjectPortofilo()
        {

            Items = new List<Dictionary<string, string>>();
        }
        public override string ToString()
        {
            string result = Name + "\t" + TotalValue + "\t" + Bookmark + "\n";
            foreach (var item in Items)
            {
                foreach (var kvp in item)
                {
                    result += $"Key: {kvp.Key}, Value: {kvp.Value}" + "\n";
                }
            }
            return result;
        }
        public int GetAmountOfKeyValue()
        {
            int count = 0;
            foreach (var item in Items)
            {
                foreach (var kvp in item)
                {
                    count++;
                }
            }
            return count;
        }
    }
}
