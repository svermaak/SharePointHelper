using Microsoft.SharePoint.Client;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace SharePointHelper.Core
{
    public class PDFThumbnail
    {
        public static bool ProcessList(string SiteUrl, string Site, string Web, string pdfListName, string thumbnailListName)
        {
            Logging.LogMessage("PDFThumbnail - ProcessList - Starting");
            try
            {
                using (ClientContext clientContext = new ClientContext(SiteUrl))
                {
                    Web site = clientContext.Web;
                    List opdfList = site.Lists.GetByTitle(pdfListName);
                    List othumbnailList = site.Lists.GetByTitle(thumbnailListName);

                    CamlQuery oQuery = new CamlQuery();
                    oQuery.ViewXml = "<OrderBy><FieldRef Ascending='False' Name='ID' /></OrderBy>";


                    ListItemCollection collPDFListItem = opdfList.GetItems(oQuery);

                    clientContext.Load(collPDFListItem);
                    clientContext.ExecuteQuery();

                    if (collPDFListItem.Count == 0)
                    {
                        Logging.LogMessage("PDFThumbnail - ProcessList - No PDFs");
                        return true;
                    }
                    else
                    {
                        foreach (ListItem listItem in collPDFListItem)
                        {
                            try
                            {
                                //Get fileref
                                string pdfFileRef = listItem["FileRef"].ToString();

                                //Get pdf filename from fileref
                                string pdfFileName = getFileName(pdfFileRef);

                                //Check if file is a pdf
                                if (pdfFileName.EndsWith(".pdf"))
                                {
                                    string thumbnailFileName = pdfFileName.Substring(0, pdfFileName.Length - 3) + "bmp";
                                    string thumbnailFileRef = "/Lists/" + thumbnailListName + "/" + thumbnailFileName;

                                    oQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileRef'></FieldRef><Value Type='Text'>" + thumbnailFileRef + "</Value></Eq></Where></Query></View>";
                                    ListItemCollection collThumbnailListItem = othumbnailList.GetItems(oQuery);

                                    clientContext.Load(collThumbnailListItem);
                                    clientContext.ExecuteQuery();


                                    string tempPath = Path.GetTempPath();


                                    if (collThumbnailListItem.Count() > 0)
                                    {
                                        //Thumbnail exist
                                    }
                                    else
                                    {
                                        using (var client = new WebClient())
                                        {
                                            client.UseDefaultCredentials = true;
                                            client.DownloadFile(SiteUrl + "/" + pdfFileRef, tempPath + pdfFileName);
                                        }

                                        PdfDocument doc = new PdfDocument();

                                        //Catching silly licencing error
                                        try
                                        {
                                            doc.LoadFromFile(tempPath + pdfFileName);
                                        }
                                        catch
                                        {

                                        }

                                        //Resize and save image to temp
                                        Image bmp = doc.SaveAsImage(0);
                                        bmp = ResizeImage(bmp, (int)(bmp.Width * 0.2), (int)(bmp.Height * 0.2));
                                        bmp.Save(tempPath + thumbnailFileName,ImageFormat.Bmp);

                                        //Uload image
                                        UploadPicture(SiteUrl, thumbnailListName, tempPath + thumbnailFileName);

                                        Logging.LogMessage("PDFThumbnail - ProcessList - Created thumbnail for '" + thumbnailFileName + "'");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                //Item failed
                                Logging.LogMessage("PDFThumbnail - ProcessList - Item failed (" + ex.Message + ")");
                            }
                        }
                    }

                    //Check in thumnails
                    oQuery.ViewXml = "<OrderBy><FieldRef Ascending='False' Name='ID' /></OrderBy>";

                    ListItemCollection collThumnailListItem =  othumbnailList.GetItems(oQuery);

                    clientContext.Load(collThumnailListItem);
                    clientContext.ExecuteQuery();

                    if (collThumnailListItem.Count == 0)
                    {
                        Logging.LogMessage("PDFThumbnail - ProcessList - All thumbnails checked in");
                    }
                    else
                    {
                        foreach (ListItem listItem in collThumnailListItem)
                        {
                            try
                            {
                                if (listItem.File.CheckedOutByUser != null)
                                {
                                    //listItem["Enterprise Meta Data"] = new System.Collections.Generic.Dictionary[System.String, System.Object];
                                    listItem.File.CheckIn("From code", CheckinType.MajorCheckIn);
                                    listItem.File.Publish("");
                                    clientContext.ExecuteQuery();
                                }
                            }
                            catch
                            {

                            }
                        }
                    }
                }
                Logging.LogMessage("PDFThumbnail - ProcessList - Done");
                return true;
            }
            catch (Exception ex)
            {
                Logging.LogMessage("PDFThumbnail - ProcessList - Error occurred (" + ex.Message + ")");
                return false;
            }
        }
        private static string getFileName(string fileRef)
        {
            return fileRef.Split('/')[fileRef.Split('/').GetUpperBound(0)];
        }

        public static byte[] ImageToByte(Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }
        public static void UploadPicture(string SiteUrl, string ListName, string FileName)
        {
            using (var clientContext = new ClientContext(SiteUrl))
            {
                using (var fs = new FileStream(FileName, FileMode.Open))
                {
                    var fi = new FileInfo(FileName);
                    var list = clientContext.Web.Lists.GetByTitle(ListName);
                    clientContext.Load(list.RootFolder);
                    clientContext.ExecuteQuery();
                    var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fi.Name);

                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, fileUrl, fs, true);
                }
            }
        }
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
    }
}
