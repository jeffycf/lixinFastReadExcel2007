using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace LixinFastReadExcel07
{
    class Common
    {
     static    Regex letterReg = new Regex("[a-z|A-Z]+");//将excel的字母变成数字的正则
        public static int letter2Num(string letter)
        {
            Match myMatch = letterReg.Match(letter);
            int result = 0;
            if (myMatch.Success)
            {
                char[] words = myMatch.Value.ToUpper().ToCharArray();
                for (int i = 0; i < words.Length; i++)
                {
                    result += Convert.ToInt32((words[i] - 64) * Math.Pow(26, words.Length - i - 1));
                }
            }
            return result;
        }
    }
}
