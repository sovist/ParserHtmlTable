using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace HtmlDocument
{
    class ParserHtmlTable
    {
        private readonly string _fileName;
        public bool HaveTable
        {
            get
            {
                WebBrowser web = new WebBrowser { DocumentText = "" };
                web.Document.OpenNew(true);
                web.Document.Write(File.ReadAllText(_fileName, Encoding.Default));
                int tablesCount = web.Document.GetElementsByTagName("table").Count;
                web.Dispose();
                return tablesCount != 0;
            }
        }

        public ParserHtmlTable(string fileName)
        {
            _fileName = fileName;
        }

        public List<List<string>> GetTableData()
        {
            if (!HaveTable)
                return new List<List<string>>();

            List<List<string>> rezList = new List<List<string>>();
            WebBrowser webRows = new WebBrowser {DocumentText = ""};
            WebBrowser webCell = new WebBrowser {DocumentText = ""};
            WebBrowser webAllTables = new WebBrowser {DocumentText = ""};

            webAllTables.Document.OpenNew(true);
            webAllTables.Document.Write(File.ReadAllText(_fileName, Encoding.Default));
            HtmlElementCollection tables = webAllTables.Document.GetElementsByTagName("table");

            for (int i = 0; i < tables.Count; i++)
            {
                webRows.Document.OpenNew(true);
                webRows.Document.Write("<HTML><BODY><TABLE>" + tables[i].InnerHtml + "</TABLE></BODY></HTML>");
                HtmlElementCollection rows = webRows.Document.GetElementsByTagName("tr");

                for (int j = 0; j < rows.Count; j++)
                {
                    webCell.Document.OpenNew(true);
                    webCell.Document.Write("<HTML><BODY><TABLE>" + rows[j].InnerHtml + "</TABLE></BODY></HTML>");
                    HtmlElementCollection cells = webCell.Document.GetElementsByTagName("td");

                    List<string> tList = (from HtmlElement col in cells select col.InnerText).ToList();
                    if (tList.Any(elem => elem != null))
                        rezList.Add(tList);
                }
            }
            webCell.Dispose();
            webRows.Dispose();
            webAllTables.Dispose();

            return rezList;
        }
    }
}
