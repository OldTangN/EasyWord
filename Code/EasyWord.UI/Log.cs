using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using log4net;

namespace EasyWord.UI
{
    public static class Log
    {
        static ILog logger = LogManager.GetLogger("");

        /// <summary>
        /// 异常
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="ex"></param>
        public static void Error(string msg, Exception ex)
        {
            logger.Error(msg, ex);
        }

        /// <summary>
        /// 消息
        /// </summary>
        /// <param name="msg"></param>
        public static void Info(string msg)
        {
            logger.Info(msg);
        }
    }
}
