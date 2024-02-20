
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

        private Dictionary<string, Group> groupsList = new Dictionary<string, Group>();
        private List<string> curStudentList;

        private TimeTable curTimeTable = null;
        public mainForm()
        {
            InitializeComponent();
            app = new ExcelApp();
            string rootPath = Directory.GetCurrentDirectory();
            fileManager = new FileManager(rootPath);

            List<string> groupsNamesList = fileManager.readDataToListFromFileMbNExist(@"студенты\группы");
            List<string> subGroups = fileManager.readDataToListFromFileMbNExist(@"студенты\подгруппы");

            for (int i = 0; i < groupsNamesList.Count; i++)
            {
                try
                {
                    groupsList.Add(groupsNamesList[i], new Group(groupsNamesList[i], int.Parse(subGroups[i])));

                }catch (Exception) { }

            }

            groupsListBox.Items.AddRange(groupsNamesList.ToArray());
            if (groupsListBox.Items.Count == 0)
                hideAllGroupItems();
            else
            {
                groupsListBox.SelectedIndex = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //foreach (string group in groupsList)
            //{
            //    MessageBox.Show(group + "123");
            //}
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
            string groupIndex = groupsListBox.SelectedText;
            Group group = groupsList[groupIndex];
            string groupName = group.Name;

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
            fileManager.writeListToFile(@"\студенты\" + groupName, studentList);

        }

        private List<string> readStudentListFromFile(string groupName)
        {
            return fileManager.readDataToListFromFile(@"\студенты\" + groupName);
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
            if(groupsList[groupName] != null)
            {
                MessageBox.Show("Группа с названием \""+ groupName + "\" уже найдена" + Environment.NewLine 
                    + "Новая группа не может быть создана!", "Ошибка при добавлении группы", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            writeStudentListToFile(groupName, students);

            //groupsList.Add(groupName, );
            //if(groupsList.Count == 1)
            //{
            //    showAllGroupItems();
            //}
            //fileManager.writeListToFile("groups", groupsList);


            //groupsListBox.Items.Clear();
            //groupsListBox.Items.AddRange(groupsList.ToArray());

            //groupsListBox.SelectedIndex = groupsList.IndexOf(groupName);
        }
        private void removeGroup(string groupName)
        {
            Group index = this.groupsList[groupName];
            if (index == null)
            {
                MessageBox.Show("Группа с названием \"" + groupName + "\" не найдена" + Environment.NewLine
                    + "Удаление невозможно", "Ошибка при удалении группы",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.groupsList.Remove(groupName);
            List<string> groupsList = new List<string>();
            foreach(var group in this.groupsList)
            {
                groupsList.Add(group.Value.Name);
            }
            fileManager.writeListToFile(@"студенты\группы", groupsList);

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
            Group group = groupsList[groupsListBox.SelectedText];
            var resutl = MessageBox.Show("Вы уверены, что хотите удалить группу \""+ group.Name + "\"?",
                "Удаление группы", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(resutl == DialogResult.Yes)
                removeGroup(group.Name);
        }

        private void createTable_Click(object sender, EventArgs e)
        {
            app.createTable(curStudentList);
        }

        private void timeTableButton_Click(object sender, EventArgs e)
        {
            curTimeTable = app.readTimeTable(Directory.GetCurrentDirectory() + @"\2_kurs_2_sem.xls");
            MessageBox.Show("Загрузка расписания окончена успешно");

            GroupTimeTable timeTable;

            foreach(var group in curTimeTable.getGroups())
            {
                timeTable = group.Value.getTimetable();
                fileManager.writeListToFile(@"расписание\числитель\" + group.Key, timeTable.Numerator);
                fileManager.writeListToFile(@"расписание\знаменатель\" + group.Key, timeTable.Denominator);
            }

        }

        private void writeTimetableButton_Click(object sender, EventArgs e)
        {
            Group group = groupsList[groupsListBox.SelectedText];

            //Group group = curTimeTable.getGroup(groupName);

            string msg = "Раписание:" + Environment.NewLine;

            for (int i = 0; i < group.SubGroupsCount; i++)
            {
                msg += Environment.NewLine + "Группа №" + (i + 1);
                foreach (string pare in group.getTimetable(i).Numerator)
                {
                    msg += Environment.NewLine + pare;
                }
            }
            timeTableTextBox.Text = msg;
        }
    }
}
