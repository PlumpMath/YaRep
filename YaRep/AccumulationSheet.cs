using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YaRep
{
    public sealed class AccumulationSheet : Sheet
    {

        internal AccumulationSheet()
        {
        }

        internal override object[,] GetArray()
        {
            if (/*(maxValues < 1) || */(data.Count < 1)) return null;

            object[,] result = new object[maxValues+1, data.Count];

            int setNum = 0;
            foreach (var set in data)
            {
                result[0, setNum] = set.Key;

                int valueIndex = 0;
                foreach (object value in set.Value)
                {
                    result[++valueIndex, setNum] = value;
                }

                setNum++;
            }


            return result;
        }

        int maxValues = 0;
        Dictionary<string, List<object>> data = new Dictionary<string, List<object>>();

        public void AddValue(string setName, object value)
        {
            List<object> setList = null;
            if (!data.TryGetValue(setName, out setList)) { setList = new List<object>(); data.Add(setName, setList); }
            setList.Add(value);
            if (setList.Count > maxValues) maxValues = setList.Count;
        }

        public void AddValues(string setName, IEnumerable<object> values)
        {
            List<object> setList = null;
            if (!data.TryGetValue(setName, out setList)) { setList = new List<object>(); data.Add(setName, setList); }
            setList.AddRange(values);
            if (setList.Count > maxValues) maxValues = setList.Count;
        }

        public void AddSet(string setName)
        {
            if (data.ContainsKey(setName)) return;
            data.Add(setName, new List<object>());
        }

    }
}
