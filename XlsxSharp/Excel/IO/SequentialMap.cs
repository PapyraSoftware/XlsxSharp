using XlsxSharp.Utils;

namespace XlsxSharp.Excel.IO;

internal class SequentialMap<TKey, T>
    where TKey : struct
{
    /// <summary>
    /// The index is the one that is used to save the <typeparamref name="T"/> while value is an index to the <c>_fullMap</c>.
    /// </summary>
    private readonly Dictionary<int, TKey> _savedIdToActualId = new();

    private readonly IReadOnlyBiDictionary<TKey, T> _fullMap;

    /// <summary>
    /// A table that will be saved to file. Contains used and necessary entries along with
    /// the id under which the entry can be retrieved.
    /// </summary>
    private List<(int SaveId, T Actual)>? _saveTable;

    public SequentialMap(IReadOnlyBiDictionary<TKey, T> fullMap) => this._fullMap = fullMap;

    /// <summary>
    /// How many entries to save are in the map.
    /// </summary>
    public int Count => this._savedIdToActualId.Count;

    internal static SequentialMap<TKey, T> Create(
        HashSet<T> usedValues,
        IReadOnlyBiDictionary<TKey, T> allValuesMap,
        IReadOnlyDictionary<T, int>? firstValues = null,
        int usedStart = 0
    )
    {
        SequentialMap<TKey, T> map = new(allValuesMap);
        firstValues ??= new Dictionary<T, int>();
        foreach ((T firstValue, int savedId) in firstValues)
        {
            // A predefined number format's usual id can be occupied by a differently-spelled but
            // equivalent format code instead (e.g. a real file writing "$"#,##0.00 rather than the
            // bare $#,##0.00 this predefined entry uses) - nothing reserves that saved id then.
            if (!allValuesMap.ContainsValue(firstValue))
            {
                continue;
            }

            TKey actualId = allValuesMap[firstValue];
            map.Add(actualId, savedId);
        }

        // This is here basically for number formats. It ensures that user defined number
        // formats start at 164 and the 0-164 is reserved for predefined formats.
        // Number formats is the only table that can have gaps in the ids.
        int usedSaveId = Math.Max(map.Count, usedStart);
        foreach ((TKey actualId, T value) in allValuesMap)
        {
            if (firstValues.ContainsKey(value))
            {
                continue;
            }

            if (!usedValues.Contains(value))
            {
                continue;
            }

            map.Add(actualId, usedSaveId++);
        }

        map.Sort();
        return map;
    }

    public void Add(TKey actualId) =>
        this._savedIdToActualId.Add(this._savedIdToActualId.Count, actualId);

    private void Add(TKey actualId, int saveId) => this._savedIdToActualId.Add(saveId, actualId);

    public void Sort() =>
        this._saveTable = [
            .. this
                ._savedIdToActualId.Select(x => (x.Key, this._fullMap[x.Value]))
                .OrderBy(x => x.Item1),
        ];

    public IEnumerable<(int SaveId, T Actual)> GetActual() => this._saveTable!;

    public int GetSavedId(T item)
    {
        TKey actualId = this._fullMap[item];
        return this.GetSavedId(actualId);
    }

    public int GetSavedId(TKey actualId)
    {
        // TODO Styles: Use a better better internal structure
        foreach ((int mapSaveId, TKey mapActualId) in this._savedIdToActualId)
        {
            if (mapActualId.Equals(actualId))
            {
                return mapSaveId;
            }
        }

        throw new InvalidOperationException(
            $"Unable to find saveId for {actualId} of {typeof(T).Name}."
        );
    }
}
