using System.Collections.Generic;
using System;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Net;
using Microsoft.IdentityModel.Protocols.WSIdentity;

namespace Common
{
    public static class SiteExtensions
    {
        /// <summary>
        /// Gets all sub sites for a given site
        /// </summary>
        /// <param name="site">Site to find all sub site for</param>
        /// <returns>IEnumerable of strings holding the sub site urls</returns>
        public static IEnumerable<string> GetAllSubSites(this Site site)
        {
            var siteContext = site.Context;
            siteContext.Load(site, s => s.Url);

            try
            {
                siteContext.ExecuteQuery();
            }
            catch (System.Net.WebException clientException)
            {
                Console.WriteLine(clientException.Message.ToString());
                yield break;
            }
            catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
            {
                Console.WriteLine(unauthorizedException.Message.ToString());
                yield break;
            }

            var queue = new Queue<string>();
            queue.Enqueue(site.Url);
            while (queue.Count > 0)
            {
                var currentUrl = queue.Dequeue();
                try
                {
                    using (var webContext = CreateClientContext(currentUrl, siteContext.Credentials))
                    {
                        webContext.Load(webContext.Web, web => web.Webs.Include(w => w.Url, w => w.WebTemplate));
                        webContext.ExecuteQuery();
                        foreach (var subWeb in webContext.Web.Webs)
                        {
                            // Ensure these props are loaded...sometimes the initial load did not handle this
                            //subWeb.EnsureProperties(s => s.Url, s => s.WebTemplate);

                            // Don't dive into App webs and Access Services web apps
                            if (!subWeb.WebTemplate.Equals("App", StringComparison.InvariantCultureIgnoreCase) &&
                                !subWeb.WebTemplate.Equals("ACCSVC", StringComparison.InvariantCultureIgnoreCase))
                            {
                                queue.Enqueue(subWeb.Url);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Eat exceptions when certain subsites aren't accessible, better this then breaking the complete fMedium
                }

                yield return currentUrl;
            }
        }

        public static ClientContext Clone(this ClientContext cc, string url)
        {
            return CreateClientContext(url, cc.Credentials);
        }

        internal static ClientContext CreateClientContext(string url, ICredentials credential = null)
        {
            ClientContext cc = new ClientContext(url);
            try
            {
                if (credential == null)
                {
                    cc.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                }
                else
                {
                    cc.Credentials = credential;
                }
                Web web = cc.Web;
                cc.Load(web, website => website.Title);
                cc.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.Message);
                Logging.GetInstance().WriteToLogFile(Logging.Error, ex.StackTrace);

                // Return ClientContext with url even though it will be unusable
                return new ClientContext(url);
            }
            return cc;
        }
    }

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
