using System.IO;

namespace Common
{
    /// <summary>
    /// Class to support interactions for file content
    /// </summary>
    public static class StreamExtensions
    {
        /// <summary>
        /// Transforms a string into a stream
        /// </summary>
        /// <param name="s">String to transform</param>
        /// <returns>Stream</returns>
        public static Stream GenerateStream(this string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
    }
}
