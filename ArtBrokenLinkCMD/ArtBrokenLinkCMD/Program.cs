using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Web;

namespace ArtBrokenLinkCMD
{
    static class Program
    {
        // Connect to GSX database
        private static SqlConnection sqlConnection_GSX = new SqlConnection();
        private static SqlCommand cmd = new SqlCommand();

        //Log List<string>
        private static List<string> log = new List<string>();
        private static int errorCounter = 0;

        //Main
        static void Main(string[] args)
        {
            // Get current date
            DateTime CurrentDate = DateTime.Now;

            Console.ForegroundColor = ConsoleColor.Green;

            Console.WriteLine("");
            Console.WriteLine("Current Timestamp: {0}", CurrentDate);
            Console.WriteLine("Connecting to GSX DB...");

            //Add the language to the log
            log.Add("Current Timestamp: " + CurrentDate.ToString());
            log.Add("Connecting to GSX DB...");

            //Connect to FactCSATBlackBox
            SQLConnectionToGSXDB();

            Console.WriteLine("Connected to GSX DB...");
            Console.WriteLine("");

            //Add the language to the log
            log.Add("Connected to GSX DB...");
            log.Add("");

            //Go to the main method
            Execute();

            // Close SQL Connection
            sqlConnection_GSX.Close();
        }

        //SQL Connection Info
        private static void SQLConnectionToGSXDB()
        {
            // Connect to GSX database
            sqlConnection_GSX = new System.Data.SqlClient.SqlConnection(@"Data Source=JCHOE_EliteDesk;Initial Catalog=LSXContentBI;Integrated Security=True");
            cmd = new System.Data.SqlClient.SqlCommand();
            cmd.CommandType = System.Data.CommandType.Text;

            // Execute SQL Query
            cmd.Connection = sqlConnection_GSX;
            sqlConnection_GSX.Open();
        }

        //Generate Log file
        public static void GenerateLogFile(List<string> logMessage, string lang)
        {
            // Get the current directory.
            string logpath = Directory.GetCurrentDirectory() + "\\log_" + lang + ".txt";
            //using (StreamWriter sw = File.AppendText(logpath))
            using (StreamWriter sw = new StreamWriter(logpath,true))
            {
                try
                {
                    foreach (string log in logMessage)
                    {
                        sw.WriteLine(log);
                    }
                    Console.WriteLine("Generating log file {0}", logpath);

                    //Add the language to the log
                    log.Add("");
                    log.Add("Generating log file: " + logpath);
                    log.Add("");
                }
                finally
                {
                    sw.Close();
                }
            }
        }

        //Check the URL to see if it resolves or not
        private static void checkURL(string articleURL, string artImageURL, SqlCommand cmd, string marketLLCC, int counter)
        {
            HttpWebResponse response = null;
            var request = (HttpWebRequest)WebRequest.Create(artImageURL);
            request.Method = "HEAD";
            string output = String.Empty;

            try
            {
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                // Insert into database
                cmd.CommandText = "INSERT INTO[dbo].[SOCArtBrokenLink]([ID],[LLCC],[ArticleURL],[ArtImageURL], [ErrorMessage], [Date]) VALUES (" +
                    errorCounter++ + ", '" + marketLLCC + "', N'" + articleURL + "', '" + artImageURL + "', '" + ex.Message + "', dateadd(d,0, getdate()));";

                // Execute SQL Query
                cmd.ExecuteNonQuery();

                //MessageBox.Show("/* A WebException will be thrown if the status of the response is not `200 OK` */");
                /* A WebException will be thrown if the status of the response is not `200 OK` */
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("ErrorID: " + errorCounter + "\n");
                Console.WriteLine("Locale: " + marketLLCC + "\n");
                Console.WriteLine("Article URL: " + articleURL + "\n");
                Console.WriteLine("Art Image URL: " + artImageURL + "\n");
                Console.WriteLine("Error Message: " + ex.Message + "\n");
                Console.WriteLine("Date: " + DateTime.Now + "\n");
                Console.WriteLine("\n\n");

                //Write to log file
                log.Add("ErrorID: " + errorCounter + "\n");
                log.Add("Locale: " + marketLLCC + "\n");
                log.Add("Article URL: " + articleURL + "\n");
                log.Add("Art Image URL: " + artImageURL + "\n");
                log.Add("Art Validation Status: " + ex.Message + "\n");
                log.Add("Date: " + DateTime.Now + "\n");
                log.Add("\n\n");
            }
            finally
            {
                // Don't forget to close your response.
                if (response != null)
                {
                    response.Close();
                }
            }
        }

        //Go to the sitemap and extract AssetID
        private static List<string> GetLinkToArticles(string lang)
        {
            // language counter
            int langCounter = 0;
            int counter = 0;
            string element = String.Empty;

            //string[] llcc = { "nb-NO", "zh-CN", "ko-KO" };
            StringBuilder sbSitemapURL = new StringBuilder();
            List<string> siteMapArticleAssetID = new List<string>();

            // Dictionary
            Dictionary<string, int> sitemapDetails = new Dictionary<string, int>();

            // Generate sitemap for each language
            string siteMapURL = "https://support.office.com/" + lang + "/sitemap";
            langCounter++;

            XmlTextReader sitemapReader = new XmlTextReader(siteMapURL);

            try
            {
                {
                    while (sitemapReader.Read())
                    {
                        if (sitemapReader.NodeType == XmlNodeType.Element)
                        {
                            element = sitemapReader.Name;
                        }
                        else if (sitemapReader.NodeType == XmlNodeType.Text)
                        {
                            switch (element)
                            {
                                case "loc":
                                    siteMapArticleAssetID.Add(sitemapReader.Value);
                                    counter++;
                                    break;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (sitemapReader != null)
                    sitemapReader.Close();
            }

            return siteMapArticleAssetID;
        }

        //Execute
        private static void Execute()
        {
            List<string> listOfAssetID = new List<string>();
            string imgLink = String.Empty;
            int assetIDCounter = 0;

            //string[] llcc = { "en-us", "de-de", "ar-sa", "bg-bg", "zh-cn", "zh-tw", "hr-hr", "cs-cz", "da-dk", "nl-nl", "et-ee", "fi-fi", "el-gr", "he-il", "hu-hu", "id-id", "it-it", "ja-jp", "ko-kr", "lv-lv", "lt-lt", "nb-no", "pl-pl", "pt-pt", "pt-br", "ru-ru", "sr-latn-rs", "sk-sk", "sl-si", "es-es", "sv-se", "th-th", "tr-tr", "uk-ua", "vi-vn", "zh-hk", "fr-fr", "ro-ro" };
            string[] llcc = { "zh-tw", "hr-hr", "cs-cz", "da-dk", "nl-nl", "et-ee", "fi-fi", "el-gr", "he-il", "hu-hu", "id-id", "it-it", "ja-jp", "ko-kr", "lv-lv", "lt-lt", "nb-no", "pl-pl", "pt-pt", "pt-br", "ru-ru", "sr-latn-rs", "sk-sk", "sl-si", "es-es", "sv-se", "th-th", "tr-tr", "uk-ua", "vi-vn", "zh-hk", "fr-fr", "ro-ro"};
            //string[] llcc = { "en-us", "zh-cn", "zh-tw", "hr-hr", "cs-cz", "da-dk", "nl-nl", "et-ee", "fi-fi", "el-gr", "he-il", "hu-hu", "id-id", "it-it", "ja-jp", "ko-kr", "lv-lv", "lt-lt", "nb-no", "bg-bg", "pl-pl", "pt-pt", "pt-br", "ru-ru", "sr-latn-rs", "sk-sk", "sl-si", "es-es", "sv-se", "th-th", "tr-tr", "uk-ua", "vi-vn", "zh-hk", "fr-fr", "ro-ro"};
            //string[] llcc = {"th-th", "tr-tr", "uk-ua", "vi-vn", "fr-fr", "ro-ro"};
            //string[] llcc = { "de-de", "ja-jp" };
            //string[] llcc = { "ar-SA" };

            // Generate sitemap for each language
            foreach (string lang in llcc)
            {
                // Get current date
                DateTime CurrentDate = DateTime.Now;

                //Output to the screen
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Current Timestamp: {0}", CurrentDate);
                Console.WriteLine("Executing the script...");
                Console.WriteLine("");

                //Add to the log
                log.Add("Current Timestamp: " + CurrentDate.ToString());
                log.Add("Executing the script...");
                log.Add("");

                //Grab the assetIDs for the target language
                listOfAssetID = GetLinkToArticles(lang);

                //Iterate through each language
                foreach (var assetID in listOfAssetID)
                {
                    try
                    {
                        //AssetID Counter
                        assetIDCounter++;

                        //Write out to user
                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("Processing AssetID(" + assetIDCounter + "): " + Uri.EscapeUriString(assetID));
                        Console.WriteLine("");

                        log.Add("Processing AssetID(" + assetIDCounter + "): " + Uri.EscapeUriString(assetID));
                        log.Add("");

                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(assetID);
                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        StreamReader sr = new StreamReader(response.GetResponseStream());
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.Load(sr);
                        var imgs = doc.DocumentNode.SelectNodes("//img[@src]");
                        if (imgs == null)
                        { //do nothing}
                          //return;
                        }

                        int counter = 0;

                        //richTextBox1.Text += "AssetID: " + assetID + "\n";

                        foreach (HtmlNode img in imgs)
                        {
                            if (img.Attributes["src"] == null)
                                continue;
                            HtmlAttribute src = img.Attributes["src"];
                            if (src.Value.Contains("https://") == true)
                            {
                                counter++;
                                checkURL(assetID, src.Value, cmd, lang, counter);
                            }
                            else if (src.Value.Contains(@"/Images/") == true)
                            {
                                imgLink = "https://support.office.com" + src.Value;
                                counter++;
                                checkURL(assetID, imgLink, cmd, lang, counter);
                            }
                        }
                        sr.Close();
                    }
                    catch (Exception)
                    {
                        string exceptionMsg = "Sorry, the page you’re looking for can’t be found. You may have clicked an old link or the page may have moved.";
                        string imgLinkMsg = "Not available";

                        // Insert into database
                        cmd.CommandText = "INSERT INTO[dbo].[SOCArtBrokenLink]([ID],[LLCC],[ArticleURL],[ArtImageURL], [ErrorMessage], [Date]) VALUES (" +
                            errorCounter++ + ", '" + lang + "', N'" + assetID + "', '" + imgLinkMsg + "', '" + exceptionMsg + "', dateadd(d,0, getdate()));";

                        // Execute SQL Query
                        cmd.ExecuteNonQuery();

                        //MessageBox.Show("/* A WebException will be thrown if the status of the response is not `200 OK` */");
                        /* A WebException will be thrown if the status of the response is not `200 OK` */
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("ErrorID: " + errorCounter + "\n");
                        Console.WriteLine("Locale: " + lang + "\n");
                        Console.WriteLine("Article URL: " + assetID + "\n");
                        Console.WriteLine("Art Image URL: " + imgLinkMsg + "\n");
                        Console.WriteLine("Error Message: " + exceptionMsg + "\n");
                        Console.WriteLine("Date: " + DateTime.Now + "\n");
                        Console.WriteLine("\n\n");

                        //Write exception to log file
                        log.Add("ErrorID: " + errorCounter + "\n");
                        log.Add("Locale: " + lang + "\n");
                        log.Add("Article URL: " + assetID + "\n");
                        log.Add("Art Image URL: " + imgLinkMsg + "\n");
                        log.Add("Error Message: " + exceptionMsg + "\n");
                        log.Add("Date: " + DateTime.Now + "\n");
                        log.Add("\n\n");
                    }
                }

                Console.WriteLine("");
                Console.WriteLine("Finished executing the script...");
                Console.WriteLine("Current Timestamp: {0}", DateTime.Now);
                Console.WriteLine("");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Completed successfully!");
                Console.WriteLine("");

                //Add the language to the log
                log.Add("");
                log.Add("Finished executing the script...");
                log.Add("Current Timestamp: " + DateTime.Now);
                log.Add("");
                log.Add("Completed successfully!");

                //Generate Log file
                GenerateLogFile(log, lang);
                log.Clear();

                //Reset
                listOfAssetID.Clear();
                assetIDCounter = 0;
            }

            Console.CursorVisible = false;
            Console.ForegroundColor = ConsoleColor.Green;
        }
    }
}
