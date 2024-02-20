using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelSharp
{
    internal class TimeTable
    {
        private Dictionary<string, Group> timeTable;
        private int[] paresCount;
        public TimeTable(Dictionary<string, Group> groupsTimeTables, int[] paresCount)
        {
            timeTable = groupsTimeTables;
            this.paresCount = paresCount;
        }
        public GroupTimeTable getGroupTimeTable(string name, int subGroup = 0)
        {
            return getGroup(name).getTimetable(subGroup);
        }
        public Group getGroup(string name)
        {
            //timeTable
            return timeTable[name];
        }
        public Dictionary<string, Group> getGroups()
        {
            return timeTable;
        }
    }
}
