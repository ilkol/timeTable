
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace excelSharp
{
    public partial class mainForm : Form
    {
        private ExcelApp app;
        private FileManager fileManager;

        private List<string> groupsList;
        private List<string> curStudentList;

        private TimeTable curTimeTable = null;
        public mainForm()
        {
            InitializeComponent();
            app = new ExcelApp();
            string rootPath = Directory.GetCurrentDirectory();
            fileManager = new FileManager(rootPath);

            if (File.Exists(rootPath + "groups.data"))
            {
                File.Create(rootPath + "groups.data");
            }
            groupsList = readDataToListFromFile("groups");

            //groupListBox.Enter += (s, e) => { groupListBox.Parent.Focus(); };

            groupsListBox.Items.AddRange(groupsList.ToArray());
            if (groupsListBox.Items.Count == 0)
                hideAllGroupItems();
                
            else
            {
                groupsListBox.SelectedIndex = 0;
            }
            //List<string> students = new List<string>
            //{
            //    "test",
            //    "ad",
            //    "asd"
            //};

            //writeStudentListToFile("ОБ-02.03.02", students);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            foreach (string group in groupsList)
            {
                MessageBox.Show(group + "123");
            }
            //MessageBox.Show(groupsListBox.SelectedIndex.ToString());

            //List<string> students = new List<string>();
            //List<string> studList = readStudentListFromFile("ОБ-02.03.02");
            //foreach(string student in studList)
            //{
            //    MessageBox.Show(student);
            //}
        }
        private void updateGroupListBox()
        {
            int groupIndex = groupsListBox.SelectedIndex;
            string groupName = groupsList[groupIndex];

            curStudentList = readStudentListFromFile(groupName);

            groupListBox.Text = "Список группы \"" + groupName + "\":" + Environment.NewLine;
            int counter = 1;
            foreach (string student in curStudentList)
            {
                groupListBox.Text += Environment.NewLine + (counter++) + ". " + student;
            }

        }
        private void writeStudentListToFile(string groupName, List<string> studentList)
        {
            fileManager.writeListToFile(@"\students\" + groupName, studentList);

        }

        private List<string> readStudentListFromFile(string groupName)
        {
            return readDataToListFromFile(@"\students\" + groupName);
        }
        private List<string> readDataToListFromFile(string filePath)
        {
            string data = "";
            List<string> studentList = new List<string>();
            try
            {
                data = fileManager.readFile(filePath + @".data");
            }
            catch (Exception ex)
            {
                if (ex is FileNotFoundException)
                {
                    MessageBox.Show("Файл не был найден");
                    //File.Create(filePath + @".data");
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
                if (student.Length > 1)
                    studentList.Add(student);
                data = data.Substring(pos + 2);
                pos = data.IndexOf(Environment.NewLine);
            }
            return studentList;
        }

        private void groupsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            updateGroupListBox();
        }

        private void addGroupButton_Click(object sender, EventArgs e)
        {
            

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = Directory.GetCurrentDirectory();
                openFileDialog.Filter = "excel (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    List<string> list = app.readStudentsFromExcel(filePath);
                    string groupName = list[0];
                    list.Remove(groupName);
                    addGroup(groupName, list);
                }
            }

        }

        private void addGroup(string groupName, List<string> students)
        {
            if(groupsList.IndexOf(groupName) != -1)
            {
                MessageBox.Show("Группа с названием \""+ groupName + "\" уже найдена" + Environment.NewLine 
                    + "Новая группа не может быть создана!", "Ошибка при добавлении группы", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            writeStudentListToFile(groupName, students);

            groupsList.Add(groupName);
            if(groupsList.Count == 1)
            {
                showAllGroupItems();
            }
            fileManager.writeListToFile("groups", groupsList);


            groupsListBox.Items.Clear();
            groupsListBox.Items.AddRange(groupsList.ToArray());

            groupsListBox.SelectedIndex = groupsList.IndexOf(groupName);
        }
        private void removeGroup(string groupName)
        {
            int index = groupsList.IndexOf(groupName);
            if (index == -1)
            {
                MessageBox.Show("Группа с названием \"" + groupName + "\" не найдена" + Environment.NewLine
                    + "Удаление невозможно", "Ошибка при удалении группы",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            groupsList.Remove(groupName);
            fileManager.writeListToFile("groups", groupsList);

            groupsListBox.Items.Clear();
            groupsListBox.Items.AddRange(groupsList.ToArray());
            if(groupsListBox.Items.Count > 0)
            {
                groupsListBox.SelectedIndex = 0;
            }
            else
            {
                hideAllGroupItems();
            }
        }
        private void hideAllGroupItems()
        {
            groupsListBox.Visible = false;
            groupListBox.Visible = false;
            removeGroupButton.Visible = false;
        }
        private void showAllGroupItems()
        {
            groupsListBox.Visible = true;
            groupListBox.Visible = true;
            removeGroupButton.Visible = true;
        }

        private void removeGroupButton_Click(object sender, EventArgs e)
        {
            string name = groupsList[groupsListBox.SelectedIndex];
            var resutl = MessageBox.Show("Вы уверены, что хотите удалить группу \""+name+"\"?",
                "Удаление группы", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(resutl == DialogResult.Yes)
                removeGroup(name);
        }

        private void createTable_Click(object sender, EventArgs e)
        {
            app.createTable(curStudentList);
        }

        private void timeTableButton_Click(object sender, EventArgs e)
        {
            curTimeTable = app.readTimeTable(Directory.GetCurrentDirectory() + @"\2_kurs_2_sem.xls");

            foreach(var group in curTimeTable.getGroups())
            {
                fileManager.writeListToFile(@"timtable\" + group.Key, group.Value.getTimetable().Numerator);
            }

        }

        private void writeTimetableButton_Click(object sender, EventArgs e)
        {
            //int groupIndex = groupsListBox.SelectedIndex;
            //string groupName = groupsList[groupIndex];

            //Group group = curTimeTable.getGroup(groupName);

            //string msg = "Раписание:" + Environment.NewLine;

            //for(int i = 0; i < group.SubGroupsCount; i++)
            //{
            //    msg += Environment.NewLine + "Группа №" + (i + 1);
            //    foreach(string pare in group.getTimetable(i))
            //    {
            //        msg += Environment.NewLine + pare;
            //    }
            //}
            //timeTableTextBox.Text = msg;
        }
    }
}
