using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace LixinFastReadExcel07
{
    public class Cell
    {
        public int colIndex { get; set; }

        public CellType dataType { get; set; }
        public string v { get; set; }
        public string t { get; set; }
        public string value
        {
            get
            {
                if (dataType == CellType.s)
                {
                    return ShareString.shareWords[Convert.ToInt32(v)];
                }
                else if (dataType == CellType.inlineStr)
                {
                    return t;
                }
                else
                {
                    return v;
                }
            }
         
        }

        public Cell(XmlNode node)
        {
            colIndex = Common.letter2Num(node.Attributes["r"].Value) - 1;
            if (node.Attributes["t"] != null)
                dataType = (CellType)Enum.Parse(typeof(CellType), node.Attributes["t"].Value);
            else
                dataType = CellType.n;

            if (node.SelectSingleNode("./x:v", Space.excelName) != null)
                v = node.SelectSingleNode("./x:v", Space.excelName).InnerText;
            foreach (XmlNode myT in node.SelectNodes("./x:is/x:t", Space.excelName))
            {
                t += myT.InnerText;
            }
        }
    }
    public enum CellType
    {
        b,//Boolean
        n,//Number
        e,//Error
        s,//Shared String
        str,//String
        inlineStr//Inline String
    }

}
