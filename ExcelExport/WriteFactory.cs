using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Write
{
    public enum WriteTo { CSV, JSON }

    public class WriteFactory
    {
        public static IWrite? Create(WriteTo tp)
        {
            switch (tp)
            {
                case WriteTo.CSV: return new CSVWriter();
                case WriteTo.JSON: return new JsonWriter();
                default:
                    break;
            }

            return default;
        }
    }

    public class WriteBase
    {
        Encoding _encoding = Encoding.UTF8;
        public bool SetEncoding(Encoding ec) { _encoding = ec; return true; }
        public Encoding GetEncoding() { return _encoding; }

        public bool OverwriteToFile(string path, string file, string text)
        {
            var fullPath = Path.Combine(path, file);
            if (!TryDeleteFile(fullPath))
            {
                return false;
            }

            try
            {
                File.AppendAllText(fullPath, text, _encoding);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }

            return true;
        }

        private bool TryDeleteFile(string fullPath)
        {
            try
            {
                if (!File.Exists(fullPath))
                    return true;
                
                File.Delete(fullPath);
            }
            catch 
            {
                return false;
            }

            return true;
        }
    }

    public interface IWrite
    {
        public bool ToSave(string path, string file, DataTable dt);
    }

    public class JsonWriter : WriteBase, IWrite
    {
        private const string ext = ".json";

        public bool ToSave(string path, string file, DataTable dt)
        {
            file += "-" + dt.TableName;
            file = Path.ChangeExtension(file, ext);
            string json = JsonConvert.SerializeObject(dt/*, Formatting.Indented*/);

            return OverwriteToFile(path, file, json);
        }
    }

    public class CSVWriter : WriteBase, IWrite
    {
        private const string ext = ".csv";

        public bool ToSave(string path, string file, DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            string[] columnNames = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            var header = string.Join(",", columnNames);
            sb.AppendLine(header);

            List<string> valueLines = dt.AsEnumerable().Select(r => string.Join(",", r.ItemArray)).ToList();
            var rows = string.Join(Environment.NewLine, valueLines);
            sb.AppendLine(rows);

            file += "-" + dt.TableName;
            file = Path.ChangeExtension(file, ext);
            return OverwriteToFile(path, file, sb.ToString());
        }
    }
}
