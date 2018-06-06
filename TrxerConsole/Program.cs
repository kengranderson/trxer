using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Xsl;
using System.Text.RegularExpressions;
using System.IO;

namespace TrxerConsole
{
    class Program
    {
        /// <summary>
        /// Embedded Resource name
        /// </summary>     
        private const string XSLT_FILE = "Trxer.xslt";

        /// <summary>
        /// Trxer output format
        /// </summary>
        private const string OUTPUT_FILE_EXT = ".html";

        /// <summary>
        /// Main entry of TrxerConsole
        /// </summary>
        /// <param name="args">First cell shoud be TRX path</param>
        static void Main(string[] args)
        {
            if (args.Any() == false)
            {
                Console.WriteLine("No trx file,  Trxer.exe <filename>");
                return;
            }
            Console.WriteLine("Trx File\n{0}", args[0]);
            Transform(args[0], PrepareXsl());
        }

        /// <summary>
        /// Transforms trx int html document using xslt
        /// </summary>
        /// <param name="fileName">Trx file path</param>
        /// <param name="xsl">Xsl document</param>
        private static void Transform(string fileName, XmlDocument xsl)
        {
            XslCompiledTransform x = new XslCompiledTransform();

            var settings = new XsltSettings(true, false);

            // https://msdev.pro/2012/09/14/xslloadexception-the-type-or-namespace-name-securityrulesattribute-does-not-exist-in-the-namespace-system-security-are-you-missing-an-assembly-reference/
            XsltArgumentList xsltArgs = new XsltArgumentList();
            XsltFunctions ext = new XsltFunctions();
            xsltArgs.AddExtensionObject("urn:my-scripts", ext);

            x.Load(xsl, settings, null);

            Console.WriteLine("Transforming...");
            FileStream fs = File.Create(fileName + OUTPUT_FILE_EXT);
            x.Transform(fileName, xsltArgs, fs);
            fs.Close();
            Console.WriteLine("Done transforming xml into html");
        }

        /// <summary>
        /// Loads xslt form embedded resource
        /// </summary>
        /// <returns>Xsl document</returns>
        private static XmlDocument PrepareXsl()
        {
            XmlDocument xslDoc = new XmlDocument();
            Console.WriteLine($"Loading xslt template {XSLT_FILE}...");
            xslDoc.Load(ResourceReader.StreamFromResource(XSLT_FILE));
            MergeCss(xslDoc);
            MergeJavaScript(xslDoc);
            return xslDoc;
        }

        /// <summary>
        /// Merges all javascript linked to page into Trxer html report itself
        /// </summary>
        /// <param name="xslDoc">Xsl document</param>
        private static void MergeJavaScript(XmlDocument xslDoc)
        {
            Console.WriteLine("Loading javascript...");
            XmlNode scriptEl = xslDoc.GetElementsByTagName("script")[0];
            XmlAttribute scriptSrc = scriptEl.Attributes["src"];
            string script = ResourceReader.LoadTextFromResource(scriptSrc.Value);
            scriptEl.Attributes.Remove(scriptSrc);
            scriptEl.InnerText = script;
        }

        /// <summary>
        /// Merges all css linked to page ito Trxer html report itself
        /// </summary>
        /// <param name="xslDoc">Xsl document</param>
        private static void MergeCss(XmlDocument xslDoc)
        {
            Console.WriteLine("Loading css...");
            XmlNode headNode = xslDoc.GetElementsByTagName("head")[0];
            XmlNodeList linkNodes = xslDoc.GetElementsByTagName("link");
            List<XmlNode> toChangeList = linkNodes.Cast<XmlNode>().ToList();

            foreach (XmlNode xmlElement in toChangeList)
            {
                XmlElement styleEl = xslDoc.CreateElement("style");
                styleEl.InnerText = ResourceReader.LoadTextFromResource(xmlElement.Attributes["href"].Value);
                headNode.ReplaceChild(styleEl, xmlElement);
            }
        }

        public class XsltFunctions
        {
            public string RemoveAssemblyName(string asm)
            {
                int idx = asm.IndexOf(',');
                if (idx == -1)
                    return asm;
                return asm.Substring(0, idx);
            }

            public string RemoveNamespace(string asm)
            {
                int coma = asm.IndexOf(',');
                return asm.Substring(coma + 2, asm.Length - coma - 2);
            }

            public string GetShortDateTime(string time)
            {
                if (string.IsNullOrEmpty(time))
                {
                    return string.Empty;
                }

                return DateTime.Parse(time).ToString();
            }

            private string ToExtactTime(double ms)
            {
                if (ms < 1000)
                    return ms + " ms";

                if (ms >= 1000 && ms < 60000)
                    return string.Format("{0:0.00} seconds", TimeSpan.FromMilliseconds(ms).TotalSeconds);

                if (ms >= 60000 && ms < 3600000)
                    return string.Format("{0:0.00} minutes", TimeSpan.FromMilliseconds(ms).TotalMinutes);

                return string.Format("{0:0.00} hours", TimeSpan.FromMilliseconds(ms).TotalHours);
            }

            public string ToExactTimeDefinition(string duration)
            {
                if (string.IsNullOrEmpty(duration))
                {
                    return string.Empty;
                }

                return ToExtactTime(TimeSpan.Parse(duration).TotalMilliseconds);
            }

            public string ToExactTimeDefinition(string start, string finish)
            {
                TimeSpan datetime = DateTime.Parse(finish) - DateTime.Parse(start);
                return ToExtactTime(datetime.TotalMilliseconds);
            }

            public string CurrentDateTime()
            {
                return DateTime.Now.ToString();
            }

            public string ExtractImageUrl(string text)
            {
                Match match = Regex.Match(text, "('|\")([^\\s]+(\\.(?i)(jpg|png|gif|bmp)))('|\")",
                   RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    return match.Value.Replace("\'", string.Empty).Replace("\"", string.Empty).Replace("\\", "\\\\");
                }
                return string.Empty;
            }
        }
    }
}
