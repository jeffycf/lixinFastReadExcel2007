using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
namespace LixinFastReadExcel07
{
    public class SheetData
    {
        XmlReader rowReader;
        Stream ss;
        Row[] _rows;
        public Row[] rows
        {
            get
            {
                if (_rows != null)
                    return _rows;
                else
                {
                    parseRows();
                    return _rows;
                }
            }

        }


        public SheetData(Stream ss)
        {
            this.ss = ss;
        }
       /// <summary>
       /// 一次性把所有的内容都加载进rows
       /// </summary>
        void parseRows()
        {
            XmlReader myReader = XmlReader.Create(ss);
            List<Row> ls = new List<Row>();
            while (myReader.Read())
            {
                while (myReader.IsStartElement() && myReader.Name == "row")
                {
                    ls.Add(new Row(myReader.ReadOuterXml()));

                }
            }
            _rows = ls.OrderBy(a => a.r).ToArray();
            ls.Clear();
            myReader.Close();
        }
        /// <summary>
        /// 每次调用读取一行row返回
        /// </summary>
        /// <returns></returns>
        public Row ReadOneRow()
        {
            if (rowReader == null)
                rowReader = XmlReader.Create(ss);
            do
            {
                if (rowReader.EOF)
                    return null;
                if (rowReader.IsStartElement() && rowReader.Name == "row")
                    return new Row(rowReader.ReadOuterXml());
            } while (rowReader.Read());
            return null;
        }

        public void close()
        {
            _rows = null;
            rowReader.Close();
            ss = null;
        }
    }
}
