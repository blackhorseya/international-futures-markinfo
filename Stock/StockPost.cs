using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Stock
{
    class StockPost
    {
        public HtmlAgilityPack.HtmlDocument Posthtmldoc()
        {
            try
            {
                WebClient wc = new WebClient();
                MemoryStream ms = new MemoryStream(wc.DownloadData("http://histock.tw/global"));
                HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
                htmldoc.Load(ms, Encoding.UTF8);
                string xpath = "//html[1]//body[1]//div[@class='globals']//div[3]//div[3]//div[1]//ul[1]";

                htmldoc.LoadHtml(htmldoc.DocumentNode.SelectSingleNode(xpath).InnerHtml);
                
                ms.Close();
                wc.Dispose();

                return htmldoc;
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                return null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public HtmlAgilityPack.HtmlDocument PostStockItem(string Item)
        {
            try
            {
                WebClient wc = new WebClient();
                MemoryStream ms = new MemoryStream(wc.DownloadData(@"http://histock.tw/index/" + Item));
                HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
                htmldoc.Load(ms, Encoding.UTF8);
                //string xpath = "//html[1]//body[1]//div[1]//div[8]//div[1]//div[1]//div[4]";
                string xpath = "//html[1]//body[1]//div[@id='getFixed']//div[@class='index-data2 large-no']";

                htmldoc.LoadHtml(htmldoc.DocumentNode.SelectSingleNode(xpath).InnerHtml);
                ms.Close();
                wc.Dispose();

                return htmldoc;
            }
            catch(NullReferenceException ex)
            {
                //MessageBox.Show(ex.ToString());
                return null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public HtmlNodeCollection PostTaiwanFutures()
        {
            try
            {
                WebClient wc = new WebClient();
                MemoryStream ms = new MemoryStream(wc.DownloadData(@"http://www.cnyes.com/futures/indexftr3.aspx"));
                HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
                htmldoc.Load(ms, Encoding.UTF8);
                string xpath = "//div[@id='main3']//table[1]";
                HtmlNodeCollection nodes = htmldoc.DocumentNode.SelectNodes(xpath);

                ms.Close();
                wc.Dispose();

                return nodes;
            }
            catch(NullReferenceException ex)
            {
                return null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
