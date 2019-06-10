using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BH.Engine.Excel.Profiling
{
    public class Timer : IDisposable
    {
        public Timer(string name)
        {
            m_name = name;
            m_start = DateTime.Now;
        }

        private string m_name;
        private DateTime m_start;
        
        private static Dictionary<string, List<double>> records = new Dictionary<string, List<double>>();

        private static void RecordTime(string name, double time)
        {
            if (records.ContainsKey(name))
            {
                records[name].Add(time);
            } else
            {
                records.Add(name, new List<double> { time });
            }
        }

        public static double GetTotal(string name)
        {
            if(records.ContainsKey(name))
            {
                return records[name].Sum();
            }
            return 0;
        }

        public static double GetMean(string name)
        {
            if(records.ContainsKey(name) && records[name].Count > 0)
            {
                return records[name].Sum() / records[name].Count;
            }
            return 0;
        }

        public void Dispose()
        {
            RecordTime(m_name, (DateTime.Now - m_start).TotalMilliseconds);
        }
    }
}
