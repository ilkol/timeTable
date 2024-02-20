using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelSharp
{
    internal class GroupTimeTable
    {
        private List<string> numerator = new List<string>();
        private List<string> denominator = new List<string>();
        public GroupTimeTable() { }

        public void Add(string numerator, string denominator = null)
        {
            if(denominator == null)
            {
                this.denominator.Add(numerator);
            }
            else this.denominator.Add(denominator);
            this.numerator.Add(numerator);
        }
        public List<string> Numerator
        {
            get { return numerator; }
        }
        public List<string> Denominator
        {
            get { return denominator; }
        }

        public string NumeratorString
        {
            get
            {
                string value = "";
                foreach(string s in numerator)
                {
                    value += s + Environment.NewLine;
                }
                return value;
            }
        }
        public string DenominatorString
        {
            get
            {
                string value = "";
                foreach (string s in denominator)
                {
                    value += s + Environment.NewLine;
                }
                return value;
            }
        }
    }
}
