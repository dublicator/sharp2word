namespace Word.W2004.Elements
{
    public enum ImageLocation
    {
        /// <summary>
        /// Full path absolute (from the root of your server.) including file name and extension.
        /// It has to start from the root of your system.
        /// </summary>
        FULL_LOCAL_PATH,
        /// <summary>
        /// It can be http://localhost/your_app/img/xxx.gif or http://google.com/img/logoWhatever.png
        /// To know if it will work, you should be able to see this image in your browser
        /// </summary>
        WEB_URL,
        /// <summary>
        /// Application classpath
        /// </summary>
        CLASSPATH
    }
}