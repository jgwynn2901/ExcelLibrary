namespace ExcelLibrary
{
    public class ColumnData<T>
    {
        private readonly T _data;

        /// <summary>Initializes a new instance of the <see cref="ColumnData{T}" /> class.</summary>
        /// <param name="value">The value.</param>
        public ColumnData(T value)
        {
            _data = value;
        }

        /// <summary>Initializes a new instance of the <see cref="ColumnData{T}" /> class.</summary>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        public ColumnData(string name, T value) : this(value)
        {
            ColumnName = name;
        }

        /// <summary>Gets the value.</summary>
        /// <returns>
        ///   T Value
        /// </returns>
        public T GetValue() 
        {
            return _data;
        }
        public string ColumnName { get; set; }

        

    }
}
