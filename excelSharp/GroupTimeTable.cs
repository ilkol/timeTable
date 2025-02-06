using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelSharp
{
    internal class GroupTimeTable
    {
        private List<DayTimetable> numerator = new List<DayTimetable>();
        private List<DayTimetable> denominator = new List<DayTimetable>();
        private int[] paresCount = new int[6];
        public GroupTimeTable() { }
        public int[] PareCount
        {
            set
            {
                paresCount = value;
            }
            get
            {
                return paresCount;
            }
        }
        public void set(List<string> numerator, List<string> denominator, int[] count)
        {
            prepareTimeTable(ref this.denominator, numerator, count);
            prepareTimeTable(ref this.numerator, numerator, count);
        }
        private void prepareTimeTable(ref List<DayTimetable> table, List<string> source, int[] count)
        {
            int day = 0;
            int index = 0;
            List<string> list = new List<string>();
            foreach (var pare in source)
            {
                if (count[day] >= index)
                {
                    table.Add(new DayTimetable(list, count[day]));
                    index = 0;
                    day++;
                    list = new List<string>();
                }
                list.Add(pare);
                index++;
            }
        }
        public List<DayTimetable> Numerator
        {
            get { return numerator; }
        }
        public List<DayTimetable> Denominator
        {
            get { return denominator; }
        }
        
        public string NumeratorString
        {
            get
            {
                string value = "";
                
                foreach(DayTimetable day  in numerator)
                {
                    foreach(string pare in day.Timetable)
                    {
                        value += pare + Environment.NewLine;

                    }
                }
                return value;
            }
        }
        public string DenominatorString
        {
            get
            {
                string value = "";
                foreach (DayTimetable day in denominator)
                {
                    foreach (string pare in day.Timetable)
                    {
                        value += pare + Environment.NewLine;

                    }
                }
                return value;
            }
        }
    }
}
