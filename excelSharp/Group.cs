using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelSharp
{
    internal class Group
    {
        private readonly string _name;
        private List<GroupTimeTable> tibetableList = new List<GroupTimeTable>();
        private List<List<string>> studentsList = new List<List<string>>();
        public string Name { get { return _name; } }

        public Group(string name)
        {
            this._name = name;
        }
        public void addStudentList(List<string> list)
        {
            studentsList.Add(list);
           
        }
        public List<string> StudentList
        {
            get
            {
                List<string> list = new List<string>();
                foreach (List<string> studList in studentsList)
                {
                    list.AddRange(studList);
                }
                return list;
            }

        }
        public List<string> StudentListToSave
        {
            get
            {
                List<string> list = new List<string>();
                foreach (List<string> studList in studentsList)
                {
                    list.AddRange(studList);
                    list.Add("-");
                }
                list.Remove(list[list.Count - 1]);

                return list;
            }

        }

        public List<List<string>> GroupList
        {
            get
            {
                
                return studentsList;
            }

        }
        public void setGroupList(List<List<string>> list)
        {
            studentsList = list;
        }
        public void addTimetable(GroupTimeTable tibetable)
        {
            tibetableList.Add(tibetable);
        }
        public GroupTimeTable getTimetable(int subGroup = 0) 
        {
            if (subGroup < 0 || subGroup >= tibetableList.Count)
                return new GroupTimeTable();
            var tmp = tibetableList[subGroup];
            if (tmp == null)
                return new GroupTimeTable();
            return tmp;
        }
        public int SubGroupsCount
        {
            get { return studentsList.Count; }
        
        }
        public List<GroupTimeTable> TimetableList
        {
            set
            {
                tibetableList = value;

            }
            get
            {
                return tibetableList;
            }
        }
    }
}
