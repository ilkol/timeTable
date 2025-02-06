using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
        public void checkDir(string filePath)
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
        public List<string> readDataToListFromFileMbNExist(string filePath)
        {
            checkDir(filePath);
            if (!File.Exists(filePath + ".data"))
            {
                File.Create(filePath + ".data");
            }
            return readDataToListFromFile(filePath);
        }
        public List<string> readDataToListFromFile(string filePath)
        {
            string data = "";
            List<string> studentList = new List<string>();
            try
            {
                data = readFile(filePath + @".data");
            }
            catch (Exception ex)
            {
                if (ex is FileNotFoundException)
                {
                    MessageBox.Show("Файл не был найден");
                }
                else
                {
                    MessageBox.Show(filePath + ".data" + Environment.NewLine + Environment.NewLine + ex.Message);

                }
                return studentList;
            }
            int pos = data.IndexOf(Environment.NewLine);
            string student;

            while (pos != -1)
            {
                student = data.Substring(0, pos);
                if (student.Length > 0)
                    studentList.Add(student);
                data = data.Substring(pos + 2);
                pos = data.IndexOf(Environment.NewLine);
            }
            if(data.Length > 0)
            {
                studentList.Add(data);
            }
            return studentList;
        }
    }
}
