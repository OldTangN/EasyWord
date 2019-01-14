using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyWord.UI
{
    public class ReplacePara
    {
        public Dictionary<string, string> ReplaceDatas { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// 文档目录
        /// </summary>
        public string FilePath { get; set; } = "";

        /// <summary>
        /// 是否替换同目录文件
        /// </summary>
        public bool All { get; set; } = false;

        public string FileNameFrom { get; set; } = "";
        public string FileNameTo { get; set; } = "";
    }
}
