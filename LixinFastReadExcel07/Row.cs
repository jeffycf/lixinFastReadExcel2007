using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace LixinFastReadExcel07
{
  public   class Row
    {
        /// <summary>
        /// Row Index 第几行的row
        /// </summary>
        public int r
        {
            get;
            set;
        }
        public string innerString { get; set; }
        public List<Cell> cells { get; set; }
        public Row(string content)
        {
            XmlDocument myDoc = new XmlDocument();
            myDoc.LoadXml(content);
            if (myDoc.DocumentElement.Attributes["r"] != null)
                r = Convert.ToInt32(myDoc.DocumentElement.Attributes["r"].Value);
            else
                r = 0;
            cells = new List<Cell>();
            foreach (XmlNode myNode in myDoc.DocumentElement.SelectNodes("/x:row/x:c",Space.excelName))
            {
                Cell myCell = new Cell(myNode);
              cells.Insert(myCell.colIndex, myCell);
            }

        }
    }
}
