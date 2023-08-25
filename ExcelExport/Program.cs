using ExcelExport.Write;
using System.Data;
using System.Diagnostics;

namespace ExcelExport
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = Directory.GetCurrentDirectory();
            var files = GetXlsxFiles(path);
            var to = WriteTo.CSV;

            Stopwatch sw = Stopwatch.StartNew();

            WriteFromDataTable(to, files, path);

            sw.Stop();
            Console.WriteLine(sw.Elapsed.ToString());
        }

        static string[] GetXlsxFiles(string targetDirectory)
        {
            if (!Directory.Exists(targetDirectory))
                return Array.Empty<string>();

            return Directory.GetFiles(targetDirectory, "*.xlsx");
        }

        static bool WriteFromDataTable(WriteTo to, string[] files, string path)
        {
            Parallel.ForEach(files, file =>
            {
                WriteFromDataTable(to, file, path);
            });

            //foreach (var xlsx in files)
            //    WriteFromDataTable(to, xlsx, path);

            return true;
        }

        static bool WriteFromDataTable(WriteTo to, string xlsx, string path)
        {
            var name = Path.GetFileNameWithoutExtension(xlsx); 
            if (null == name)
                return false;

            using var er = new ExcelReader(xlsx, true);
            if (!er.Open())
                return false;

            foreach (var sheet in er.GetNames())
            {
                var dt = er.ToDataTable(sheet);
                if (dt == null)
                    continue;

                if (!WriteFile(to, path, name, dt))
                    continue;
            }

            return true;
        }

        private static bool WriteFile(WriteTo to, string path, string file, DataTable dt)
        {
            if (to == WriteTo.JSON)
            {
                var jsonWriter = WriteFactory.Create(WriteTo.JSON);
                if (jsonWriter == null)
                    return false;

                if (!jsonWriter.ToSave(path, file, dt))
                {
                    Console.WriteLine("Fail - {0}", file);
                    return false;
                }
            }
            else if (to == WriteTo.CSV)
            {
                var csvWriter = WriteFactory.Create(WriteTo.CSV);
                if (csvWriter == null)
                    return false;

                if (!csvWriter.ToSave(path, file, dt))
                {
                    Console.WriteLine("Fail - {0}", file);
                    return false;
                }
            }

            return true;
        }
    }
}