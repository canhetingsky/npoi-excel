using System.IO;

namespace npoi_excel
{
    class Logger
    {
        public static void AddLogToTXT(string logstring, string filePath)
        {
            if (!File.Exists(filePath))
            {
                FileStream stream = File.Create(filePath);
                stream.Close();
                stream.Dispose();
            }
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine(logstring);
            }
        }
    }
}
