using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Xml;
using System.Text.RegularExpressions;
//作者：李鑫
//email：lixin@lixin.me
//2013年5月20日
//完成第一个版本

//Demo
//LixinFastReadExcel myExcel = new LixinFastReadExcel(@"d:\website\test2.xlsx");
// myExcel.loadShareString();
//string[] mySheetNames = myExcel.getSheetName();
//myExcel.OpenSheet(0);
// while (myExcel.SheetRead())
//{
//  string[] data=  myExcel.mySheetRowData;
   
//}
namespace LixinFastReadExcel07
{
    public class LixinFastReadExcel
    {
        #region 变量定义
        XmlReader myReader; //读取sheet数据的
        ZipFile myZip;
        Stream sheetStream;
        public string[] mySheetRowData;//保存通过sheetRead得到的行数据
        Regex letterReg = new Regex("[a-z|A-Z]+");//将excel的字母变成数字的正则
        XmlNamespaceManager ExcelNameSpace = new XmlNamespaceManager(new NameTable());
        XmlNamespaceManager ExcelPackageSpace = new XmlNamespaceManager(new NameTable());
        public string filePath = string.Empty;//excel文件的路径
        List<SheetStruct> mySheetList;//sheet的集合
        public string[] myShareString;//

        #endregion
        /// <summary>
        /// 版本号
        /// </summary>
        public string version
        {
            get { return "v0.1 update 2013-5-17 by lixin"; }
        }

        public LixinFastReadExcel(string file)
        {
            ExcelNameSpace.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            ExcelPackageSpace.AddNamespace("x", "http://schemas.openxmlformats.org/package/2006/relationships");
            if (!File.Exists(file))
                throw new IOException("系统找不到指定路径；file not exists");

            this.filePath = file;
            myZip = new ZipFile(filePath);
        }
        public void close()
        {
            myZip.Close();
            myReader.Close();
        }
        /// <summary>
        /// 获取所有sheet的名字
        /// </summary>
        /// <returns></returns>
        public string[] getSheetName()
        {

            if (mySheetList == null)
                loadSheetsInfo();
            return mySheetList.OrderBy(a => a.index).Select(a => a.name).ToArray();
        }
        /// <summary>
        /// 读取所有sheet表的基本信息
        /// </summary>
        public void loadSheetsInfo()
        {

            mySheetList = new List<SheetStruct>();
            int index = myZip.FindEntry("xl/workbook.xml", true);
            Stream myXmlStream = myZip.GetInputStream(index);
            XmlDocument doc = new XmlDocument();
            doc.Load(myXmlStream);
            XmlNodeList sheets = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", ExcelNameSpace);
            index = myZip.FindEntry("xl/_rels/workbook.xml.rels", true);
            Stream relStream = myZip.GetInputStream(index);
            XmlDocument doc2 = new XmlDocument();
            doc2.Load(relStream);
            int tempIndex = 0;
            foreach (XmlNode sheet in sheets)
            {
                SheetStruct item = new SheetStruct() { name = sheet.Attributes["name"].Value, id = sheet.Attributes["r:id"].Value, index = tempIndex++ };
                item.path = "xl/" + doc2.SelectSingleNode("/x:Relationships/x:Relationship[@Id='" + item.id + "']", ExcelPackageSpace).Attributes["Target"].Value;
                mySheetList.Add(item);
            }
            myXmlStream.Dispose();
            relStream.Dispose();
        }

        /// <summary>
        ///把字母转换成十进制。A=>1; AA=>27
        /// </summary>
        /// <param name="letter"></param>
        /// <returns></returns>
        public int letter2Num(string letter)
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
        /// <summary>
        /// 加载sharestring
        /// 该方法将全部sharestring加入内存数字，如果体积太大的话，可能出现性能问题，但不一次加载，读取sheet内容时又会慢些。
        /// </summary>
        public void loadShareString()
        {
            int index = myZip.FindEntry("xl/sharedStrings.xml", true);
            Stream myStream = myZip.GetInputStream(index);
            XmlReader myXReader = XmlReader.Create(myStream);
            string content = string.Empty;
            bool inEntry = false;
            bool inContent = false;
            int postion = 0;
            while (myXReader.Read())
            {
                switch (myXReader.NodeType)
                {
                    case XmlNodeType.Element:
                        if (myXReader.Name == "sst")
                        {
                            int count = Convert.ToInt32(myXReader.GetAttribute("uniqueCount"));
                            myShareString = new string[count];
                        }
                        else if (myXReader.Name == "si")
                        {
                            inEntry = true;
                        }
                        else if (myXReader.Name == "t" && inEntry)
                        {
                            inContent = true;
                        }
                        break;
                    case XmlNodeType.Text:
                        if (inContent)
                        {
                            content += myXReader.Value;
                        }
                        break;
                    case XmlNodeType.EndElement:
                        if (myXReader.Name == "si")
                        {
                            inEntry = false;
                            myShareString[postion++] = content;
                            content = string.Empty;
                        }
                        else if (myXReader.Name == "t")
                        {
                            inContent = false;
                        }
                        break;
                }
            }
            myXReader.Close();
            myStream.Dispose();
        }

        public void OpenSheet(int index)
        {
            if (mySheetList == null)
                loadSheetsInfo();
            string sheetPath = mySheetList[index].path;
            int indexEntry = myZip.FindEntry(sheetPath, true);
            sheetStream = myZip.GetInputStream(indexEntry);
            myReader = XmlReader.Create(sheetStream);
        }
        /// <summary>
        /// 读取一行sheet内容。读完后把该行的数据保存在mySheetRowData数组中。
        /// </summary>
        /// <returns></returns>
        public bool SheetRead()
        {
            bool endRow = false;
            int colIndex = 0;
            bool isShareString = false;
            bool inContent = false;
            while (!endRow && myReader.Read())
            {
                switch (myReader.NodeType)
                {
                    case XmlNodeType.Element:
                        switch (myReader.Name)
                        {
                            case "row":
                                int[] spans = myReader.GetAttribute("spans").Split(':').Select(a => Convert.ToInt32(a)).ToArray();
                                mySheetRowData = new string[spans[1]];
                                break;
                            case "c":
                                string letter = letterReg.Match(myReader.GetAttribute("r")).Value;
                                colIndex = letter2Num(letter) - 1;
                                isShareString = myReader.GetAttribute("t") == "s";
                                break;
                            case "v":
                                inContent = true;
                                break;
                        }
                        break;
                    case XmlNodeType.EndElement:
                        switch (myReader.Name)
                        {
                            case "row":
                                endRow = true;
                                break;
                            case "c":
                                break;
                            case "v":
                                inContent = false;
                                break;
                        }
                        break;
                    case XmlNodeType.Text:
                        string content = string.Empty;
                        if (inContent && isShareString)
                        {
                            content = myShareString[Convert.ToInt32(myReader.Value)];
                        }
                        else if (inContent && !isShareString)
                        {
                            content = myReader.Value;
                        }
                        mySheetRowData[colIndex] = content;
                        break;
                }
            }
            return endRow;
        }//end func


    }
}
