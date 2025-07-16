using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Outlook2Excel
{
    public class ImportObject
    {
        private string[,] _data = new string[1, 2];

        public ImportObject(string key, string regex)
        {
            _data[0, 0] = key;
            _data[0, 1] = regex;
        }

        public string Key => _data[0, 0];
        public string Regex => _data[1, 0];
    }
}
