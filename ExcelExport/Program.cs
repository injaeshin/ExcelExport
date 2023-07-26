using ExcelExport.Write;

namespace ExcelExport
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = Directory.GetCurrentDirectory();
            var files = GetXlsxFiles(path);
            
            foreach ( var file in files )
            {
                WriteFromDataTableAsync(file, path);
            }
        }

        static string[] GetXlsxFiles(string targetDirectory)
        {
            if (!Directory.Exists(targetDirectory))
                return Array.Empty<string>();

            return Directory.GetFiles(targetDirectory, "*.xlsx");
        }

        static bool WriteFromDataTableAsync(string file, string path)
        {   
            var er = new ExcelReader(file, true);
            if (!er.Open())
                return false;

            foreach (var name in er.GetNames())
            {
                var dt = er.ToDataTable(name);
                if (dt == null)
                    continue;

                var targetFile = Path.GetFileNameWithoutExtension(file);
                var jsonWriter = WriteFactory.Create(WriteTo.JSON);
                if (jsonWriter == null)
                    continue;

                if (!jsonWriter.ToSave(path, targetFile, dt))
                {
                    Console.WriteLine("Fail - {0}", targetFile);
                    continue;
                }

                var csvWriter = WriteFactory.Create(WriteTo.CSV);
                if (csvWriter == null)
                    continue;
                
                if (!csvWriter.ToSave(path, targetFile, dt))
                {
                    Console.WriteLine("Fail - {0}", targetFile);
                    continue;
                }
            }

            return true;
        }
    }
}