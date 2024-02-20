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
        private readonly int subGroups;
        private readonly List<GroupTimeTable> tibetableList = new List<GroupTimeTable>();
        private readonly List<List<string>> sutdentsList = new List<List<string>>();
        public string Name { get { return _name; } }

        public Group(string name, int subgroups)
        {
            this._name = name;
            this.subGroups = subgroups;
        }
        public void addStudentList(List<string> list)
        {
            sutdentsList.Add(list);
           
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
            get { return subGroups; }
        }
    }
}
