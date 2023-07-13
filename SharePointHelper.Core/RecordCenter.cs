using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace SharePointHelper.Core
{
    public class RecordCenter
    {
        public static bool ProcessList(string SiteUrl, string Site, string Web, string ListName, string TriggeringColumn, string TriggeringValue)
        {
            Logging.LogMessage("RecordCenter - ProcessList - Starting");
            try
            {
                using (ClientContext clientContext = new ClientContext(SiteUrl))
                {
                    Web site = clientContext.Web;
                    List oList = site.Lists.GetByTitle(ListName);

                    CamlQuery oQuery = new CamlQuery();
                    oQuery.ViewXml = "<Query><Where><Eq><FieldRef Name='" + TriggeringColumn + "'></FieldRef><Value Type='Text'>" + TriggeringValue + "</Value></Eq></Where></Query>";


                    ListItemCollection collListItem = oList.GetItems(oQuery);

                    clientContext.Load(collListItem);
                    clientContext.ExecuteQuery();

                    if (collListItem.Count == 0)
                    {
                        Logging.LogMessage("RecordCenter - ProcessList - No items triggered");
                        return true;
                    }
                    else
                    {
                        foreach (ListItem listItem in collListItem)
                        {
                            try
                            {
                                string triggeringColumnValue = listItem[TriggeringColumn].ToString();
                                string filePath = listItem["FileRef"].ToString();

                                if ((triggeringColumnValue == TriggeringValue) && (!filePath.ToLower().EndsWith(".aspx"))) //Ready to be declared as record and not currently a record
                                {
                                    DeclareAsRecord(Site, Web, listItem["FileRef"].ToString());
                                }
                                else if (triggeringColumnValue != TriggeringValue)
                                {
                                    Logging.LogMessage("RecordCenter - ProcessList - Item not triggered (" + filePath + ")");
                                }
                                else if (filePath.ToLower().EndsWith(".aspx"))
                                {
                                    Logging.LogMessage("RecordCenter - ProcessList - Item already processed (" + filePath + ")");
                                }
                            }
                            catch
                            {
                                //Item failed
                                Logging.LogMessage("RecordCenter - ProcessList - Item failed");
                            }
                        }
                    }
                }
                Logging.LogMessage("RecordCenter - ProcessList - Done");
                return true;
            }
            catch (Exception ex)
            {
                Logging.LogMessage("RecordCenter - ProcessList - Error occurred (" + ex.Message + ")");
                return false;
            }
        }
        public static string DeclareAsRecord(string Site, string Web, string FileRef)
        {
            Logging.LogMessage("DeclareAsRecord - Starting");
            try
            {
                using (SPSite oSiteCollection = new SPSite(Site))
                {
                    SPWebCollection collWebsite = oSiteCollection.AllWebs;

                    using (SPWeb oWebsiteSrc = oSiteCollection.AllWebs[Web])
                    {
                        SPFile oFile = oWebsiteSrc.GetFile(FileRef);

                        string additionalInformation;
                        ArchiveFile(oFile, out additionalInformation);

                        Logging.LogMessage("DeclareAsRecord - Done (" + additionalInformation + ")");
                        return additionalInformation;
                    };
                }
            }
            catch (Exception ex)
            {
                Logging.LogMessage("DeclareAsRecord - Error occurred (" + ex.Message + ")");
                return "";
            }
        }
        private static bool ArchiveFile(SPFile file, out string additionalInformation)
        {
            Logging.LogMessage("ArchiveFile - Starting");
            try
            {
                String recordSeries = file.Item.ContentType.Name;
                OfficialFileResult returnValue;

                // WSS needs the file to be checked in to know which version to send.
                if (file.Level == SPFileLevel.Checkout)
                {
                    file.CheckIn(String.Empty, SPCheckinType.MinorCheckIn);
                }

                returnValue = file.SendToOfficialFile(recordSeries, out additionalInformation);

                // Custom code for handling the response from the service.
                switch (returnValue)
                {
                    case OfficialFileResult.MoreInformation:
                        // Notify user.
                        break;
                    case OfficialFileResult.Success:
                        Logging.LogMessage("ArchiveFile - Success (" + additionalInformation + ")");
                        return true;
                    //Notify user.
                    //break;
                    default:
                        // Handle error.
                        break;
                }
                Logging.LogMessage("ArchiveFile - Failed");
                additionalInformation = "";
                return false;
            }
            catch (Exception ex)
            {
                Logging.LogMessage("ArchiveFile - Error occurred (" + ex.Message + ")");

                additionalInformation = "";
                return false;
            }
        }
    }
}
