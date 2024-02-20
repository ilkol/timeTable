using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelSharp
{
    internal class FileManager
    {
        private string rootPath;
        public FileManager(string path)
        {
            rootPath = path + @"\";
        }
        public string readFile(string file)
        {
            string filePath = rootPath + file;
            checkDir(filePath);
            string data = "";
            try
            {
                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(fileStream);
                data = reader.ReadToEnd();
                reader.Close();
                fileStream.Close();
            }
            catch (FileNotFoundException e)
            {
                throw e;
                //MessageBox.Show("Файл не найден!");
            }
            return data;
        }
        public void writeToFile(string file, string data)
        {
            string filePath = rootPath + file;
            checkDir(filePath);

            FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            StreamWriter writer = new StreamWriter(fileStream, Encoding.UTF8);
            writer.BaseStream.Seek(0, SeekOrigin.End);
            writer.WriteLine(data);
            writer.Flush();
            writer.Close();
            fileStream.Close();
        }
        private void checkDir(string filePath)
        {
            string path = filePath.Substring(0, filePath.LastIndexOf(@"\"));
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        public void writeListToFile(string filePath, List<string> list)
        {
            string data = "";
            foreach (string student in list)
            {
                data += student + Environment.NewLine;
            }
            writeToFile(filePath + @".data", data);
        }
    }
}
