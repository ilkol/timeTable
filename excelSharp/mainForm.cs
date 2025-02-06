
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

            List<List<string>> studList;
            Group curGroup;
            string name;

            for (int i = 0; i < groupsNamesList.Count; i++)
            {
                try
                {
                    name = groupsNamesList[i];
                    studList = readStudentListFromFile(name);
                    curGroup = new Group(name);
                    curGroup.setGroupList(studList);

                    groupsList.Add(name, curGroup);
                }catch (Exception e) {
                    MessageBox.Show(e.Message, "Ошибка!");
                }

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
            string groupIndex = groupsListBox.Items[groupsListBox.SelectedIndex].ToString();
            Group group = groupsList[groupIndex];
            string groupName = group.Name;

            List<List<string>> curStudentList = group.GroupList;

            groupListBox.Text = "Список группы \"" + groupName + "\":" + Environment.NewLine;
            int counter = 1;
            int i = 1;
            foreach (List<string> subgroup in curStudentList)
            {
                groupListBox.Text += Environment.NewLine + Environment.NewLine + "Подгруппа №" + i + Environment.NewLine;
                i++;
                foreach(string student in subgroup)
                {
                    groupListBox.Text += Environment.NewLine + (counter++) + ". " + student;

                }
            }
        }
        private void writeStudentListToFile(string groupName, List<string> studentList)
        {
            fileManager.writeListToFile(@"\студенты\" + groupName, studentList);

        }

        private List<List<string>> readStudentListFromFile(string groupName)
        {
            List<List<string>> grouplist = new List<List<string>>();
            List<string> curList = new List<string>();
            List<string> list = fileManager.readDataToListFromFile(@"\студенты\" + groupName);

            foreach(string student in list)
            {
                if(student == "-")
                {
                    grouplist.Add(curList);
                    curList = new List<string>();
                }
                else
                {
                    curList.Add(student);
                }
            }
            grouplist.Add(curList);

            return grouplist;
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
                    Group group = new Group(groupName);
                    addGroup(group);
                }
            }

        }

        private void addGroup(Group group)
        {
            string groupName = group.Name;
            if (groupsList.ContainsKey(groupName))
            {
                MessageBox.Show("Группа с названием \""+ groupName + "\" уже найдена" + Environment.NewLine 
                    + "Новая группа не может быть создана!", "Ошибка при добавлении группы", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            saveGroupStudentList(group);

            groupsList.Add(groupName, group);
            if (groupsList.Count == 1)
            {
                showAllGroupItems();
            }
            
            updateGroupListBox(saveGroupListToFile());
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
            updateGroupListBox(saveGroupListToFile());

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
            string groupName = groupsListBox.Items[groupsListBox.SelectedIndex].ToString();
            Group group = groupsList[groupName];
            var resutl = MessageBox.Show("Вы уверены, что хотите удалить группу \""+ group.Name + "\"?",
                "Удаление группы", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(resutl == DialogResult.Yes)
                removeGroup(group.Name);
        }

        private void createTable_Click(object sender, EventArgs e)
        {
            string groupName = groupsListBox.Items[groupsListBox.SelectedIndex].ToString();
            Group group = groupsList[groupName];
            app.createTable(group);
        }

        private void timeTableButton_Click(object sender, EventArgs e)
        {
            curTimeTable = app.readTimeTable(Directory.GetCurrentDirectory() + @"\2_kurs_2_sem.xls");
            MessageBox.Show("Загрузка расписания окончена успешно");

            Group group;
            Group oldGroup;
            foreach (var items in curTimeTable.getGroups())
            {
                group = items.Value;
                if (groupsList.ContainsKey(items.Key))
                {
                    oldGroup = groupsList[items.Key];
                    group.setGroupList(oldGroup.GroupList);
                    
                    groupsList[items.Key] = group;
                }
                else
                {
                    addGroup(group);
                }
                group.TimetableList = items.Value.TimetableList;
            }
        }

        private void writeTimetableButton_Click(object sender, EventArgs e)
        {
            string groupIndex = groupsListBox.Items[groupsListBox.SelectedIndex].ToString();
            Group group = groupsList[groupIndex];

            string msg = "Раписание:" + Environment.NewLine;

            int i = 1;
            foreach(var item in group.TimetableList)
            {
                msg += Environment.NewLine + "Группа №" + i;
                i++;
                msg += Environment.NewLine + item.NumeratorString;
                

            }

            timeTableTextBox.Text = msg;
        }
        private List<string> saveGroupListToFile()
        {
            List<string> groupsList = new List<string>();
            List<string> subgroupsList = new List<string>();
            foreach (var group in this.groupsList)
            {
                groupsList.Add(group.Value.Name);
                subgroupsList.Add(group.Value.SubGroupsCount.ToString());

            }
            fileManager.writeListToFile(@"студенты\группы", groupsList);
            //fileManager.writeListToFile(@"студенты\подгруппы", subgroupsList);

            return groupsList;
        }
        private void updateGroupListBox(List<string> groupList)
        {
            groupsListBox.Items.Clear();
            groupsListBox.Items.AddRange(groupList.ToArray());
            if (groupsListBox.Items.Count > 0)
            {
                groupsListBox.SelectedIndex = 0;
            }
            else
            {
                hideAllGroupItems();
            }
        }

        private void changeGroupList_Click(object sender, EventArgs e)
        {
            string groupIndex = groupsListBox.Items[groupsListBox.SelectedIndex].ToString();
            Group group = groupsList[groupIndex];

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
                    group.addStudentList(list);
                    
                    updateGroupListBox();
                    saveGroupStudentList(group);
                }
            }
        }
        private void saveGroupStudentList(Group group)
        {
            writeStudentListToFile(group.Name, group.StudentListToSave);
        }
    }
}
