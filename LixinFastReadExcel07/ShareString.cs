using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
namespace LixinFastReadExcel07
{
    /// <summary>
    /// sharestring处理类
    /// 该类提供一个查询功能，根据序号查询对应的值。
    /// 预设计2种查询方式，1：把所有sharestring存进内存。2：如果太大的sharestring使用缓存+sqlite
    /// </summary>
    class ShareString
    {

        public static string[] shareWords
        {
            get;
            private set;
        }
        public static string file;
        private ShareString(Stream ss)
        {

        }
        public static void LoadShareString(string file)
        {
            if (ShareString.file == file)
                return;
            ZipFile myZip = new ZipFile(file);
            int index = myZip.FindEntry("xl/sharedStrings.xml", true);
            Stream ss = myZip.GetInputStream(index);
            XmlReader myReader = XmlReader.Create(ss);
            string content = string.Empty;
            bool inEntry = false;
            bool inContent = false;
            int postion = 0;
            while (myReader.Read())
            {
                switch (myReader.NodeType)
                {
                    case XmlNodeType.Element:
                        if (myReader.Name == "sst")
                        {
                            int count = Convert.ToInt32(myReader.GetAttribute("uniqueCount"));
                            shareWords = new string[count];
                        }
                        else if (myReader.Name == "si")
                        {
                            inEntry = true;
                        }
                        else if (myReader.Name == "t" && inEntry)
                        {
                            inContent = true;
                        }
                        break;
                    case XmlNodeType.Text:
                        if (inContent)
                        {
                            content += myReader.Value;
                        }
                        break;
                    case XmlNodeType.EndElement:
                        if (myReader.Name == "si")
                        {
                            inEntry = false;
                            shareWords[postion++] = content;
                            content = string.Empty;
                        }
                        else if (myReader.Name == "t")
                        {
                            inContent = false;
                        }
                        break;
                }
            }
            myReader.Close();
            ss.Dispose();
        }



    }
}
