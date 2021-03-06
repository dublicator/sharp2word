using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using BufferedImage = System.Drawing.Image;

namespace Word.Utils
{
    public static class ImageUtils
    {
        private static string ImageToBase64(BufferedImage image, string imageformat)
        {
            ImageFormat format;
            switch (imageformat)
            {
                case "png":
                    format = ImageFormat.Png;
                    break;
                case "bmp":
                    format = ImageFormat.Bmp;
                    break;
                case "gif":
                    format = ImageFormat.Gif;
                    break;
                case "ico":
                    format = ImageFormat.Icon;
                    break;
                case "jpg":
                    format = ImageFormat.Jpeg;
                    break;
                case "wmf":
                    format = ImageFormat.Wmf;
                    break;
                default:
                    format = null;
                    break;
            }

            using (MemoryStream ms = new MemoryStream())
            {
                // Convert Image to byte[]
                if (format==null)
                {
                    format = ImageFormat.Png;
                }
                image.Save(ms, format);
                byte[] imageBytes = ms.ToArray();

                // Convert byte[] to Base64 string
                string base64String = Convert.ToBase64String(imageBytes);
                return base64String;
            }
        }

        public static string GetImageHexaBase64(BufferedImage bufferedImage, string imageformat)
        {
            return ImageToBase64(bufferedImage, imageformat);
        }

        public static string ColorToHex(Color clr)
        {
            int red = clr.R;
            int green = clr.G;
            int blue = clr.B;

            string colorHex = "";
            colorHex += String.Format("{0:X02}", red);
            colorHex += String.Format("{0:X02}", green);
            colorHex += String.Format("{0:X02}", blue);
            return colorHex;
        }
    }
}