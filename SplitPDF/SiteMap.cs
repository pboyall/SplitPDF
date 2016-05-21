using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
namespace SplitPDF
{
    public class SiteMap
    {

        public string siteMapFile;
        //Update a node's label
        public string updateSiteMap(string ExistingValue, string NewValue)
        {
            XDocument doc = XDocument.Load(siteMapFile);
            ///graphml/graph/node/data/y:ShapeNode/y:NodeLabel`
            XDocument doc1 = XDocument.Parse(doc.ToString());
            string yednamespace = "{http://www.yworks.com/xml/graphml}";
            XNamespace ns = doc1.Root.Name.Namespace;
            //var Nodes = doc.Root.Elements().Select(x => x.Element("NodeLabel"));
            var Nodes = doc.Descendants(yednamespace + "NodeLabel");
            //Crude - need to work out how to do the select properly above based on value
            foreach (var Node in Nodes)
            {
                Console.WriteLine(Node.Value);
                System.Diagnostics.Debug.WriteLine(Node.Value);
                //Hard coded an update to test it out!
                if (Node.Value == ExistingValue)
                {
                    Node.Value = NewValue;
                }
            }
            doc.Save("NewSiteMap.graphml");

            return "0";
        }

    }

}
