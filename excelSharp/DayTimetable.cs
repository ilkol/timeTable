using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelSharp
{
    internal class DayTimetable
    {
        private int count;

        public int paresCount
        {
            get { return count; }
        }

        private List<string> timetable;

        public List<string> Timetable { get { return timetable; } }

        public DayTimetable(List<string> timetable, int pareCount = 0) {
            this.timetable = timetable;
            if(pareCount > 0)
            {
                this.count = pareCount;
            }
            else count = timetable.Count;
        }
    }
}
