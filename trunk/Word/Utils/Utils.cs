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
            return FormatXml(xml);
        }

        /// <summary>
        /// Formats the provided XML so it's indented and humanly-readable.
        /// </summary>
        /// <param name="inputXml">The input XML to format.</param>
        /// <returns></returns>
        private static string FormatXml(string inputXml)
        {
            XmlDocument document = new XmlDocument();
            document.Load(new StringReader(inputXml));

            StringBuilder builder = new StringBuilder();
            using (XmlTextWriter writer = new XmlTextWriter(new StringWriter(builder)))
            {
                writer.Formatting = Formatting.Indented;
                document.Save(writer);
            }

            return builder.ToString();
        }

    }
}