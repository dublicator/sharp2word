using System;
using System.IO;
using System.Net;
using System.Text;
using Word.Api.Interfaces;
using Word.Utils;
using BufferedImage = System.Drawing.Image;

namespace Word.W2004.Elements
{
    /// <summary>
    ///   Use this class when you want to add images to your document.
    ///   You can insert images inside Paragraphs, Tables, Headings, Header, Footer and obviously, at the body of a Document.
    /// </summary>
    public class Image : IImage, IFluentElement<Image>
    {
        private readonly BufferedImage bufferedImage;
        private readonly string path = "";
        private readonly StringBuilder txt = new StringBuilder("");
        private bool hasBeenCalledBefore;
        // size
        private string height = ""; // to be able to set this to override default

        private const string img_template = "\n<w:pict>"
                                      +
                                      "\n	<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">"
                                      + "		<v:stroke joinstyle=\"miter\"/>"
                                      + "		<v:formulas>"
                                      + "			<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>"
                                      + "			<v:f eqn=\"sum @0 1 0\"/><v:f eqn=\"sum 0 0 @1\"/>"
                                      + "			<v:f eqn=\"prod @2 1 2\"/>"
                                      + "			<v:f eqn=\"prod @3 21600 pixelWidth\"/>"
                                      + "			<v:f eqn=\"prod @3 21600 pixelHeight\"/>"
                                      + "			<v:f eqn=\"sum @0 0 1\"/>"
                                      + "			<v:f eqn=\"prod @6 1 2\"/>"
                                      + "			<v:f eqn=\"prod @7 21600 pixelWidth\"/>"
                                      + "			<v:f eqn=\"sum @8 21600 0\"/>"
                                      + "			<v:f eqn=\"prod @7 21600 pixelHeight\"/>"
                                      + "			<v:f eqn=\"sum @10 21600 0\"/>"
                                      + "		</v:formulas>"
                                      + "		<v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>"
                                      + "		<o:lock v:ext=\"edit\" aspectratio=\"t\"/>"
                                      + "	</v:shapetype>"
                                      +
                                      "\n<w:binData w:name=\"wordml://{internalFileName}\" xml:space=\"preserve\">{binary}</w:binData>"
                                      +
                                      "\n	<v:shape id=\"_x0000_i1026\" type=\"#_x0000_t75\" style=\"width:{width}pt;height:{height}pt\"><v:imagedata src=\"wordml://{internalFileName}\" o:title=\"{fileName}\"/>"
                                      + "\n	</v:shape>" + "\n</w:pict>";

        private string width = ""; // to be able to set this to override default

        // size

        /// <summary>
        ///   This is the location on your image. If you specify  "imageLocation" as WEB_URL, your path should start with "http://..."
        ///   But if you choose "imageLocation" as FULL_LOCAL_PATH, you should specify path value as absolute path from the root of your server. Eg.: /Users/YourName/imgs/...
        /// </summary>
        /// <param name = "path">Path of the image. It will depend on the location: web, local or classpath.</param>
        /// <param name = "imageLocation">
        ///   FULL_LOCAL_PATH:  Full path absolute (from the root of your server.) including file name and extension.
        ///   It has to start from the root of your system.
        /// 
        ///   WEB_URL: It can be http://localhost/your_app/img/xxx.gif or http://google.com/img/logoWhatever.png
        /// </param>
        public Image(string path, ImageLocation imageLocation)
        {
            this.path = path;
            try
            {
                if (imageLocation.Equals(ImageLocation.FULL_LOCAL_PATH))
                {
                    bufferedImage = BufferedImage.FromFile(path);
                }
                else if (imageLocation.Equals(ImageLocation.WEB_URL))
                {
                    Uri url = new Uri(path);
                    bufferedImage = ImageFromUrl(url);
                }
                else if (imageLocation.Equals(ImageLocation.CLASSPATH))
                {
                    
                    throw new NotImplementedException();
                    
                    //InputStream is = getClass().getResourceAsStream(path);
                    //bufferedImage = ImageIO.read(is);
                }
            }
            catch (IOException e)
            {
                throw new Exception(
                    "Can't create ImageIO. Maybe the path is not valid. Path: \n" + path + "\nImageLocation: " +
                    imageLocation, e);
            }
        }

        public Image(byte[] path)
        {
            try
            {
            }
            catch (IOException e)
            {
                throw new Exception(
                    "Can't create ImageIO. Maybe the path is not valid. Path: \n" +  "\nImageLocation: " +
                     e);
            }
        }


        public string OriginalWidthHeight
        {
            get
            {
                string res = bufferedImage.Width + "#" + bufferedImage.Height + "";
                return res;
            }
        }

        #region IFluentElement<Image> Members

        public Image create()
        {
            return this;
        }

        #endregion

        #region IImage Members

        public string Content
        {
            get
            {
                if (hasBeenCalledBefore)
                {
                    return txt.ToString();
                }
                else
                {
                    hasBeenCalledBefore = true;
                }
                // Placeholders: internalFileName, fileName, binary, width and height

                string[] arr = path.Split('/');
                string fileName = arr[arr.Length - 1];

                string internalFileName = DateTime.Now.Millisecond + fileName;

                // string binary = ImageUtils.getImageHexaBase64(path);
                string imageformat = path.Substring(path.LastIndexOf('.') + 1);
                string binary = ImageUtils.getImageHexaBase64(bufferedImage,
                                                              imageformat);

                setUpSize();

                string res = img_template;
                res = res.Replace("{fileName}", fileName);
                res = res.Replace("{internalFileName}", internalFileName);
                res = res.Replace("{binary}", binary);
                res = res.Replace("{width}", this.width);
                res = res.Replace("{height}", this.height);

                txt.Append(res);
                return txt.ToString();
            }
        }

        public Image setWidth(string value)
        {
            this.width = value;
            return this;
        }

        public Image setHeight(string value)
        {
            this.height = value;
            return this;
        }

        #endregion

        private static BufferedImage ImageFromUrl(Uri url)
        {
            byte[] imageData = DownloadData(url);
            MemoryStream stream = new MemoryStream(imageData);
            System.Drawing.Image img = System.Drawing.Image.FromStream(stream);
            stream.Close();
            return img;
        }

        private static byte[] DownloadData(Uri url)
        {
            try
            {
                WebRequest req = WebRequest.Create(url);
                WebResponse response = req.GetResponse();
                if (response != null)
                {
                    Stream stream = response.GetResponseStream();
                    return StreamToByteArray(stream);
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static byte[] StreamToByteArray(Stream stream)
        {
            byte[] total_stream = new byte[0];
            using (Stream input = stream)
            {
                // Setup whatever read size you want (small here for testing)
                byte[] buffer = new byte[32]; // * 1024];
                int read;

                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    byte[] stream_array = new byte[total_stream.Length + read];
                    total_stream.CopyTo(stream_array, 0);
                    Array.Copy(buffer, 0, stream_array, total_stream.Length, read);
                    total_stream = stream_array;
                }
            }
            return total_stream;
        }

        private void setUpSize()
        {
            if ("".Equals(this.width) || "".Equals(this.height))
            {
                string[] wh = OriginalWidthHeight.Split('#');
                string ww = wh[0];
                string hh = wh[1];
                if ("".Equals(this.width))
                {
                    this.width = ww;
                }
                if ("".Equals(this.height))
                {
                    this.height = hh;
                }
            }
        }


        /// <summary>
        ///   It creates an image from the Web.
        /// </summary>
        /// <param name = "path">Image full path. To know if it will work, you should be able to see this image in your browser</param>
        /// <returns></returns>
        public static Image from_WEB_URL(string path)
        {
            return new Image(path, ImageLocation.WEB_URL);
        }

        /// <summary>
        ///   It creates an image from your local machine.
        /// </summary>
        /// <param name = "path">Image full path. To know if it will work, probably you should specify full path from the root of your system.</param>
        /// <returns></returns>
        public static Image from_FULL_LOCAL_PATHL(string path)
        {
            return new Image(path, ImageLocation.FULL_LOCAL_PATH);
        }
    }
}