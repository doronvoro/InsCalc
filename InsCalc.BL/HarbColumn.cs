public class HarbColumn
{
    public string Name { get; set; }         // The camelCase name of the column
    public int Index { get; set; }          // The zero-based index of the column
    public string DisplayName { get; set; } // The column's Hebrew name

    public HarbColumn(string name, int index, string displayName)
    {
        Name = name;
        Index = index;
        DisplayName = displayName;
    }
}
