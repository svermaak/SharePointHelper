using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointHelper.Core
{
    public class Config
    {
        public static void ProcessRules()
        {
            try
            {
                NameValueCollection rules = ConfigurationManager.GetSection("Rules") as NameValueCollection;

                foreach (string ruleName in rules)
                {
                    try
                    {
                        string ruleValue = rules[ruleName];
                        string ruleType = ruleValue.Split('|')[0];

                        if (ruleType == "RecordsCenterProcessList")
                        {
                            string siteUrl = ruleValue.Split('|')[1];
                            string site = ruleValue.Split('|')[2];
                            string web = ruleValue.Split('|')[3];
                            string list = ruleValue.Split('|')[4];
                            string triggeringColumn = ruleValue.Split('|')[5];
                            string triggeringValue = ruleValue.Split('|')[6];

                            Core.RecordCenter.ProcessList(siteUrl, site, web, list, triggeringColumn, triggeringValue);
                        }
                        else if (ruleType == "PDFThumbnailProcessList")
                        {
                            string siteUrl = ruleValue.Split('|')[1];
                            string site = ruleValue.Split('|')[2];
                            string web = ruleValue.Split('|')[3];
                            string pdfListName = ruleValue.Split('|')[4];
                            string thumbnailListName = ruleValue.Split('|')[5];

                            Core.PDFThumbnail.ProcessList(siteUrl, site, web, pdfListName, thumbnailListName);
                        }
                    }
                    catch
                    {
                        //Rule failed
                    }
                }
            }
            catch
            {
                //Config failed
            }
        }
    }
}
