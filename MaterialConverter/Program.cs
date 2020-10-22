using OfficeOpenXml;
using System;
using System.IO;

namespace MaterialConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 && File.Exists(args[0]))
            {
                
                var GetDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string path = args[0];
                using (StreamReader sr = new StreamReader(path))
                {

                    var fi = new FileInfo(path);
                    using (var package = new ExcelPackage(fi))
                    {
                        var ws = package.Workbook.Worksheets[0];

                        ws.DeleteColumn(1);
                        ws.DeleteColumn(1);
                        ws.DeleteColumn(2);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);
                        ws.DeleteColumn(4);

                        if (File.Exists(path))
                        {
                            try { File.Delete(path); } catch { Console.Write("Error, cannot delete old files"); }
                        }
                        byte[] data = package.GetAsByteArray();
                        string path2 = GetDirectory + "\\final.xlsx";
                        File.WriteAllBytes(path2, data);



                    }
                }

            }
            else
            {
                Console.Write("Start this software with the excel file");
            }
        }
    }
}
