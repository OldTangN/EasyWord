using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWord.Core
{
    public class BookMark : ObservableObject
    {
        string _name;
        string _value = "null";

        public BookMark(string name)
        {
            this.Name = name;
        }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get => _name; set => Set(ref _name, value); }

        /// <summary>
        /// 文本
        /// </summary>
        public string Value { get => _value; set => Set(ref _value, value); }
    }
}
