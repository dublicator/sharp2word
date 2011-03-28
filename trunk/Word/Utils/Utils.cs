using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Word.Utils
{
    public class Util
    {
        /// <summary>
        ///   The root of the app as String.
        /// </summary>
        /// <value></value>
        public static string AppRoot
        {
            get
            {
                try
                {
                    return Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                }
                catch (IOException e)
                {
                    Trace.Write(e.StackTrace);
                    throw new IOException("Can't get app root directory", e);
                }
            }
        }

        /// <summary>
        /// </summary>
        /// <param name = "file">It is the full path to the file</param>
        /// <returns>String with the content of the file</returns>
        public static string ReadFile(string file)
        {
            return File.ReadAllText(file);
        }

        public static string Pretty(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            StringBuilder sb = new StringBuilder();
            XmlWriter xw = XmlWriter.Create(sb, new XmlWriterSettings {Indent = true});
            doc.WriteTo(xw);
            xw.Flush();
            return sb.ToString();
        }
    }
}