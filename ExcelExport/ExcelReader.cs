using ExcelDataReader;
using System.Data;

namespace ExcelExport
{
    internal class ExcelReader
    {
        private string _file;
        private bool _useHeaderRow = false;
        private bool _skipSecondRow = false;

        private DataSet? _dataSet;
        private Stream? _stream;
        private IExcelDataReader? _dataReader;

        private ExcelReader(string file)
        {
            _file = file;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public ExcelReader(string file, bool skipHeader, bool skipSecondRow = false) : this(file)
        {
            _useHeaderRow = skipHeader;
            _skipSecondRow = skipSecondRow;
        }

        public bool Open()
        {
            Clear();

            try
            {
                _stream = new FileStream(_file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                _dataReader = ExcelReaderFactory.CreateReader(_stream);
            }
            catch (Exception e)
            {
                _stream?.Dispose();
                _stream = null;
                _dataReader?.Dispose();
                _dataReader = null;

                Console.WriteLine(e.ToString());

                throw;
            }

            return true;
        }

        public void Clear()
        {
            _dataSet?.Clear();
        }

        public IEnumerable<string> GetNames()
        {
            if (_dataSet == null)
            {
                _dataSet = ToDataSet();
            }

            for (int i = 0; i < _dataSet?.Tables.Count; i++)
            {
                yield return _dataSet.Tables[i].TableName;
            }
        }

        private DataSet? ToDataSet()
        {
            if (_dataReader == null)
                return null;

            var conf = new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (tableReader) => new()
                {
                    UseHeaderRow = _useHeaderRow,
                    EmptyColumnNamePrefix = "EmptyColumn",
                    FilterRow = (rowReader) => rowReader.Depth > (_skipSecondRow ? 1 : 0)
                }
            };

            return _dataReader.AsDataSet(conf);
        }

        public DataTable? ToDataTable(int sheetIndex)
        {
            if (_dataSet == null)
            {
                _dataSet = ToDataSet();
                if (_dataSet?.Tables.Count == 0)
                    return null;
            }

            return _dataSet?.Tables[sheetIndex];
        }

        public DataTable? ToDataTable(string sheetName)
        {
            if (_dataSet == null)
            {
                _dataSet = ToDataSet();
                if (_dataSet?.Tables.Count == 0)
                    return null;
            }

            return _dataSet?.Tables[sheetName];
        }
    }
}
