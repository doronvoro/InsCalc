using System.Collections.Concurrent;

public class HarbColumnDictionary
{
    private readonly ConcurrentDictionary<int, HarbColumn> _columns;

    // Publicly expose the dictionary
    public IReadOnlyDictionary<int, HarbColumn> Columns => _columns;

    // Constructor for initialization
    public HarbColumnDictionary()
    {
        _columns = new ConcurrentDictionary<int, HarbColumn>();
        InitializeColumns();
    }

    // Initialize the dictionary with column information
    private void InitializeColumns()
    {
        var columns = new List<HarbColumn>
        {

            new(nameof(HarbRow.MainBranch), 0, "ענף ראשי"), 
            new(nameof(HarbRow.SubBranch), 1, "ענף (משני)"),
            new(nameof(HarbRow.ProductType), 2, "סוג מוצר"),
            new(nameof(HarbRow.Company), 3, "חברה"),
            new(nameof(HarbRow.InsurancePeriod), 4, "תקופת ביטוח"),
            new(nameof(HarbRow.PremiumInNis), 5, "פרמיה בש\"ח"),
            new(nameof(HarbRow.PremiumType), 6, "סוג פרמיה"),
            new(nameof(HarbRow.PolicyNumber), 7, "מספר פוליסה"),
            new(nameof(HarbRow.PlanClassification), 8, "סיווג תכנית")
        };

        foreach (var column in columns)
        {
            _columns.TryAdd(column.Index, column);
        }
    }

    public int GetIndexByName(string name)
    {
        for (int i = 0; i < _columns.Count; i++)
        {
           if( _columns[i].Name.ToLower() == name.ToLower())
            {
                return i;
            }
        }

        return -1;
      //  _columns.Where(c => c.Value.Name == name).Select(c => c.Key).
    }
}
