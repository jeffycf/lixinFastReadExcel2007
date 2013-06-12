using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Xml;
using System.Text.RegularExpressions;

namespace LixinFastReadExcel07
{

    public class LixinFastReadExcel
    {
        #region 变量定义

        public string filePath = string.Empty;//excel文件的路径
        List<SheetStruct> mySheetList;//sheet的集合


        #endregion
        /// <summary>
        /// 版本号
        /// </summary>
        public string version
        {
            get { return "v0.2 update 2013-6-12 by lixin"; }
        }

        public LixinFastReadExcel(string file)
        {
            if (!File.Exists(file))
                throw new IOException("系统找不到指定路径；file not exists");

            this.filePath = file;
            Space.excelName = new XmlNamespaceManager(new NameTable());
            Space.excelName.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            Space.excelPackage = new XmlNamespaceManager(new NameTable());
            Space.excelPackage.AddNamespace("x", "http://schemas.openxmlformats.org/package/2006/relationships");
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
            ZipFile myZip = new ZipFile(filePath);
            mySheetList = new List<SheetStruct>();
            int index = myZip.FindEntry("xl/workbook.xml", true);
            Stream myXmlStream = myZip.GetInputStream(index);
            XmlDocument doc = new XmlDocument();
            doc.Load(myXmlStream);
            XmlNodeList sheets = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", Space.excelName);
            index = myZip.FindEntry("xl/_rels/workbook.xml.rels", true);
            Stream relStream = myZip.GetInputStream(index);
            XmlDocument doc2 = new XmlDocument();
            doc2.Load(relStream);
            int tempIndex = 0;
            foreach (XmlNode sheet in sheets)
            {
                SheetStruct item = new SheetStruct() { name = sheet.Attributes["name"].Value, id = sheet.Attributes["r:id"].Value, index = tempIndex++ };
                item.path = "xl/" + doc2.SelectSingleNode("/x:Relationships/x:Relationship[@Id='" + item.id + "']", Space.excelPackage).Attributes["Target"].Value;
                mySheetList.Add(item);
            }
            myXmlStream.Dispose();
            relStream.Dispose();
            myZip.Close();
        }
        /// <summary>
        /// 获取sheet表单内容。
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public SheetData GetSheetData(int index)
        {
            if (ShareString.shareWords == null)
            {
                ShareString.LoadShareString(filePath);
            }
            ZipFile myZip = new ZipFile(filePath);
            int num = myZip.FindEntry(mySheetList[index].path, true);
            Stream ss = myZip.GetInputStream(num);
            SheetData mySheetData = new SheetData(ss);
   
            return mySheetData;
        }

    }
}
