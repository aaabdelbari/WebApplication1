using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Permissions;
using System.Threading.Tasks;
using System.Web;
using Microsoft.CSharp.RuntimeBinder;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;

namespace CorrespondanceCore
{
    public class PDFExporter
    {
        public string _connectionString;
        public int _dmbsType;
        public int _dmsId;
        public string _templateId;
        public string _basefilePath;
        public int _langId;
        public PdfDocument _document;
        public Dictionary<int, PdfPage> _pgDict;
        public int _systemNo;
        private string filePassword;
        string _token = "";
        public PDFExporter()
        {
            _token = BearerTokenHandler.GetTokenFromSessionASync();
        }
        public PDFExporter(string connectionString, string dmsId, int dbmsType, string templateId, string basefilePath, int langId, int systemNo = 0)
        {
            _connectionString = connectionString;
            _dmsId = Convert.ToInt32(dmsId);
            _templateId = templateId;
            _dmbsType = dbmsType;
            _basefilePath = basefilePath;
            _langId = langId;
            _systemNo = systemNo;
            //_token = BearerTokenHandler.GetTokenFromSessionASync();
        }

        public async Task<string> ExportToProtectedPdf(int documentId, string options, DataObjects.WatermarkOptions watermarkOptions, DataObjects.AnnotationOptions annotationOptions, string password, int fromPage, int toPage)
        {
            var fName = Guid.NewGuid().ToString() + ".pdf";
            var fileName = Path.Combine(_basefilePath, fName);

            //int toPage;
            //var fromPage = 1;
            //if (pages != null)
            //{
            //    fromPage = pages[0];
            //    toPage = pages[1];
            //}
            //else
            //    toPage = GetDocumentPageCount(_dmsId);

            List<Watermark.Watermark> allWatermarks = new List<Watermark.Watermark>();
            EDocumentsData obj = new EDocumentsData();
            if (watermarkOptions != null)
                allWatermarks = (await obj.GetResolvedWatermarks(watermarkOptions, this._connectionString)).ToList();

            List<DataObjects.TemplateImageAnnotationDD> allAnnotations = new List<DataObjects.TemplateImageAnnotationDD>();
            if (annotationOptions != null)
                allAnnotations = (await obj.GetApplicableAnnotationsForDocument(_dmsId, _connectionString)).ToList();

            var objPermission = new PermissionsData();
            var pServiceResult = await objPermission.GetApplicationSettingByKeyAsync("_APS_PDF_EXPORT_THREAD_THRESHOLD_", this._connectionString);
            var threshould = string.IsNullOrEmpty(pServiceResult) ? 4 : Convert.ToInt32(pServiceResult) - 1;


            // Dim threshould = 5 - 1
            var lowerEnd = fromPage;
            var upperEnd = lowerEnd + threshould >= toPage ? toPage : lowerEnd + threshould;

            _document = new PdfDocument();

            PdfSecuritySettings securitySettings = _document.SecuritySettings;

            securitySettings.UserPassword = password;
            securitySettings.OwnerPassword = password;
            securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;
            securitySettings.PermitAccessibilityExtractContent = false;
            securitySettings.PermitAnnotations = false;
            securitySettings.PermitAssembleDocument = false;
            securitySettings.PermitExtractContent = false;
            securitySettings.PermitFormsFill = true;
            securitySettings.PermitFullQualityPrint = false;
            securitySettings.PermitModifyDocument = true;
            securitySettings.PermitPrint = false;

            do
            {
                // Will Continue work here
                Dictionary<int, byte[]> pgResult = new Dictionary<int, byte[]>();
                for (var i = fromPage; i <= upperEnd; i++)
                {
                    var index = i;
                    var annotation = allAnnotations.Where(f => (string)f.DMSNumber == _dmsId + "-" + index).ToList();
                    var dataToPost = new { dmsId = this._dmsId, pageNo = index, options = "", strConnectionString = this._connectionString, baseFilePath = this._basefilePath, annotations = new List<DataObjects.TemplateImageAnnotationDD>(), resolvedWatermarks = new List<Watermark.Watermark>() };
                    try
                    {
                        var pResult = await (ConfigurationManager.AppSettings["BusinessApi"] + "Documents/GetDMSDocumentPageWithWatermarks").WithOAuthBearerToken(_token).PostJsonAsync(dataToPost).ReceiveBytes();

                        if (annotation.Count > 0)
                            pResult = (byte[])ApplyAnnotation(pResult, this._dmsId, index, this._basefilePath, annotation);

                        pgResult.Add(index, pResult);
                    }
                    catch (Exception ex)
                    {
                        var r = ex;
                    }
                }
                if (lowerEnd != fromPage)
                    _document = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);
                for (var pageNo = lowerEnd; pageNo <= upperEnd; pageNo++)
                {
                    using (MemoryStream memStream = new MemoryStream(pgResult[pageNo]))
                    {
                        using (var xImg = XImage.FromStream(memStream))
                        {
                            PdfPage pdfPage = new PdfPage();
                            _document.AddPage(pdfPage);
                            pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                            pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                gfx.DrawImage(xImg, 0, 0);
                            }
                        }
                    }
                }

                pgResult.Clear();
                _document.Save(fileName);
                _document.Close();
                _document.Dispose();
                if (upperEnd == toPage)
                    break;
                lowerEnd = upperEnd + 1;
                upperEnd = lowerEnd + threshould >= toPage ? toPage : lowerEnd + threshould;
            } while (true);

            return fName;
        }


        /// <summary>
        /// used to genearte PDF attachments for emails
        /// </summary>
        /// <param name="systemNo"></param>
        /// <param name="pages"></param>
        /// <param name="includeAnnotations"></param>
        /// <param name="applyWatermark"></param>
        /// <returns></returns>
        public string ExportToPdf(int systemNo, string pages = "", bool includeAnnotations = true, bool applyWatermark = true, string annotationsToInclde = "All")
        {

            // _document = New PdfDocument()
            string fName = Guid.NewGuid().ToString() + ".pdf";
            string fileName = Path.Combine(_basefilePath, fName);
            int toPage;
            int fromPage = 1;
            if (!string.IsNullOrEmpty(pages))
            {
                toPage = int.Parse(pages.Split('-')[1]);
                fromPage = int.Parse(pages.Split('-')[0]);
            }
            else
            {
                toPage = GetDocumentPageCount(_dmsId);
            }

            for (int pageNo = fromPage, loopTo = toPage; pageNo <= loopTo; pageNo++)
            {
                if (_document is null)
                {
                    _document = new PdfDocument();
                }
                else
                {
                    _document = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);
                }

                using (var image = GetImageWithAnnotationAndWatermark(_dmsId, pageNo, systemNo, Convert.ToInt32(_templateId), _basefilePath, _connectionString, includeAnnotations, applyWatermark, annotationsToInclde))
                {
                    using (var imgStream = new MemoryStream())
                    {
                        image.Save(imgStream, ImageFormat.Bmp);
                        using (var xImg = XImage.FromStream(imgStream))
                        {
                            var pdfPage = new PdfPage();
                            _document.AddPage(pdfPage);
                            pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                            pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                gfx.DrawImage(xImg, 0d, 0d);
                            }
                        }
                    }
                }

                _document.Save(fileName);
                _document.Close();
            }
            return fName;
            // End Using

        }

        public async Task<string> ExportToPdf(int documentId, string options, DataObjects.WatermarkOptions watermarkOptions, DataObjects.AnnotationOptions annotationOptions, string pages = "", int langId = 2)
        {
            int pageCount = GetDocumentPageCount(_dmsId);
            var dataToPostImage = new
            {
                docId = _dmsId,
                pageCount = pageCount
            };
            var IImageBy = await (ConfigurationManager.AppSettings["BusinessApi"] + "Documents/ExportImage").WithOAuthBearerToken(_token).PostJsonAsync(dataToPostImage).Result.GetBytesAsync();
            Bitmap bitmap = (Bitmap)Image.FromStream(new MemoryStream(IImageBy));

            int compressionRate = Convert.ToInt32(ConfigurationManager.AppSettings["compressionRate"].ToString());
            var fName = Guid.NewGuid().ToString() + ".pdf";
            var fileName = Path.Combine(_basefilePath, fName);

            int toPage;
            var fromPage = 1;
            if (pages != "")
            {
                toPage = int.Parse(pages.Split('-')[1]);
                fromPage = int.Parse(pages.Split('-')[0]);
            }
            else
                // toPage = GetDocumentPageCount(_dmsId)
                toPage = pageCount;

            List<Watermark.Watermark> allWatermarks = new List<Watermark.Watermark>();
            EDocumentsData obj = new EDocumentsData();

            try
            {
                if (watermarkOptions != null)
                    // disabling watermarks for now
                    allWatermarks = (await obj.GetResolvedWatermarks(watermarkOptions, this._connectionString))?.ToList();
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf: " + ex.ToString() + Environment.NewLine);
            }

            List<DataObjects.TemplateImageAnnotationDD> allAnnotations = new List<DataObjects.TemplateImageAnnotationDD>();
            if (annotationOptions != null)
                allAnnotations = (await obj.GetApplicableAnnotationsForDocument(_dmsId, _connectionString)).ToList();
            allAnnotations.All(c =>
            {
                c.LangId = langId;
                return true;
            });

            var lowerEnd = fromPage;
            var upperEnd = toPage;

            _document = new PdfDocument();
            // Do
            // Will Continue work here
            Dictionary<int, byte[]> pgResult = new Dictionary<int, byte[]>();

            // For i = fromPage To upperEnd + 1
            for (var i = fromPage; i <= upperEnd; i++)
            {
                var index = i;
                var creteria = "" + _dmsId + "-" + i + "";
                var annotation = allAnnotations.Where(f => f.DMSNumber.Equals(creteria)).ToList();
                try
                {
                    bitmap.SelectActiveFrame(FrameDimension.Page, index - 1);
                    MemoryStream byteStream = new MemoryStream();
                    bitmap.Save(byteStream, ImageFormat.Tiff);

                    // IImage.SelectActiveFrame(FrameDimension.Page, index - 1)
                    // Dim pResult = IImageBy
                    // pResult = ApplyWatermark(New Bitmap(New MemoryStream(pResult)), allWatermarks)
                    // Compress only in one stage 
                    var pResult = ApplyWatermark(Image.FromStream(byteStream), allWatermarks);

                    if (annotation.Count > 0)
                        pResult = (byte[])ApplyAnnotation(pResult, this._dmsId, index - 1, this._basefilePath, annotation);

                    // pgResult.Add(index - 1, byteStream.ToArray())
                    pgResult.Add(index - 1, pResult);
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf: " + ex.ToString() + Environment.NewLine);
                }
            }

            if (lowerEnd != fromPage)
                _document = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);

            for (var pageNo = lowerEnd; pageNo <= upperEnd; pageNo++)
            {
                using (MemoryStream memStream = new MemoryStream(pgResult[pageNo - 1]))
                {
                    // string path = ConfigurationManager.AppSettings["tempPath"].ToString() + "_1_" + _dmsId.ToString() + documentId.ToString() + _templateId.ToString() + pageNo.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".tiff";
                    // SaveTiff(path, Image.FromStream(memStream), compressionRate);
                    using (var xImg = XImage.FromStream(memStream))
                    {
                        // Using xImg = XImage.FromStream(memStream)
                        PdfPage pdfPage = new PdfPage();
                        _document.AddPage(pdfPage);
                        pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                        pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                        using (var gfx = XGraphics.FromPdfPage(pdfPage))
                        {
                            gfx.SmoothingMode = XSmoothingMode.HighSpeed;// = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                            gfx.DrawImage(xImg, 0, 0);
                        }
                    }
                }
            }

            pgResult.Clear();

            _document.Save(fileName);
            _document.Close();
            _document.Dispose();

            lowerEnd = upperEnd + 1;

            return fName;
        }

        public async Task<string> ExportToPdfEmail(int documentId, string pages = "", bool includeAnnotations = true, bool applyWatermarkOnPage = false, int langId = 2)
        {
            string _token = BearerTokenHandler.GetTokenFromSessionASync();
            int pageCount = GetDocumentPageCount(_dmsId);
            var dataToPostImage = new
            {
                docId = _dmsId,
                pageCount = pageCount
            };

            var IImageBy = await (ConfigurationManager.AppSettings["BusinessApi"] + "Documents/ExportImage").WithOAuthBearerToken(_token).PostJsonAsync(dataToPostImage).Result.GetBytesAsync();
            Image bitmap = Image.FromStream(new MemoryStream(IImageBy));

            int compressionRate = Convert.ToInt32(ConfigurationManager.AppSettings["compressionRate"].ToString());
            var fName = Guid.NewGuid().ToString() + ".pdf";
            var fileName = Path.Combine(_basefilePath, fName);

            int toPage;
            var fromPage = 1;
            if (pages != "")
            {
                toPage = int.Parse(pages.Split('-')[1]);
                fromPage = int.Parse(pages.Split('-')[0]);
            }
            else
                toPage = pageCount;

            var watermarkOptions = new DataObjects.WatermarkOptions();
            watermarkOptions = null;
            var annotationOptions = new DataObjects.AnnotationOptions();
            List<Watermark.Watermark> allWatermarks = new List<Watermark.Watermark>();
            EDocumentsData obj = new EDocumentsData();


            if (watermarkOptions != null)
                // disabling watermarks for now
                allWatermarks = (await obj.GetResolvedWatermarks(watermarkOptions, this._connectionString)).ToList();


            List<DataObjects.TemplateImageAnnotationDD> allAnnotations = new List<DataObjects.TemplateImageAnnotationDD>();
            var lowerEnd = fromPage;
            var upperEnd = toPage;

            _document = new PdfDocument();
            if (lowerEnd != fromPage)
                _document = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);

            Dictionary<int, byte[]> pgResult = new Dictionary<int, byte[]>();
            for (var i = fromPage; i <= upperEnd; i++)
            {
                var index = i;
                var creteria = "" + _dmsId + "-" + i + "";
                var annotation = allAnnotations.Where(f => f.DMSNumber.Equals(creteria)).ToList();
                try
                {
                    bitmap.SelectActiveFrame(FrameDimension.Page, index - 1);
                    MemoryStream byteStream = new MemoryStream();
                    bitmap.Save(byteStream, ImageFormat.Tiff);

                    // Compress only in one stage 
                    //var pResult = ApplyWatermark(Image.FromStream(byteStream), allWatermarks);
                    byte[] pResult = null;
                    /*if (watermarkOptions != null)
                        pResult = ApplyWatermark(Image.FromStream(byteStream), allWatermarks);
                    else*/
                    pResult = byteStream.ToArray();

                    try
                    {
                        if (annotation.Count > 0)
                            pResult = (byte[])ApplyAnnotation(pResult, this._dmsId, index - 1, this._basefilePath, annotation);
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf_ApplyAnnotation(): " + ex.ToString() + Environment.NewLine);
                    }

                    string path = ConfigurationManager.AppSettings["tempPath"].ToString() + "_1_" + _dmsId.ToString() + documentId.ToString() + _templateId.ToString() + i.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".tiff";
                    Image.FromStream(new MemoryStream(pResult)).Save(path);
                    //SaveTiff(path, Image.FromStream(new MemoryStream(pResult)), compressionRate);

                    using (var xImg = XImage.FromFile(path))
                    //using (var xImg = XImage.FromStream(new MemoryStream(pResult)))
                    {
                        PdfPage pdfPage = new PdfPage();
                        _document.AddPage(pdfPage);
                        pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                        pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                        using (var gfx = XGraphics.FromPdfPage(pdfPage))
                        {
                            gfx.SmoothingMode = XSmoothingMode.HighSpeed;// = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                            gfx.DrawImage(xImg, 0, 0);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf: " + ex.ToString() + Environment.NewLine);
                }
            }

            pgResult.Clear();

            _document.Save(fileName);
            _document.Close();
            _document.Dispose();

            SharedFunctions.CompressPdf(fileName);
            lowerEnd = upperEnd + 1;
            return fName;
        }

        public async Task<string> MultiExportToPdf(List<DataObjects.MultiDocsForPdfDO> multiDocs, bool applyWatermark, DataObjects.AnnotationOptions annotationOptions, string strConnectionString, int intUserId, string baseFilePath, int intDBMSType, int langId = 2)
        {
            if (multiDocs == null) return "";
            var fName = Guid.NewGuid().ToString() + ".pdf";
            var fileName = Path.Combine(baseFilePath, fName);
            var _doc = new PdfDocument();
            foreach (var singleDoc in multiDocs)
            {
                var waterMarkOptions = new DataObjects.WatermarkOptions();
                waterMarkOptions.LangId = langId;
                waterMarkOptions.SystemNo = singleDoc.SysNo.ToString();
                waterMarkOptions.TemplateId = singleDoc.TempId;
                waterMarkOptions.UserId = intUserId;
                waterMarkOptions.UserIPAddress = GetUserIPAddress();

                if (!applyWatermark)
                {
                    var objPermission = new PermissionsData();
                    var _objNo = Convert.ToDouble(Convert.ToInt32(ConfigurationManager.AppSettings["permissionStartKey"]) + singleDoc.TempId);
                    var objResult = objPermission.GetUserPermissions(intUserId, (int)Math.Round(_objNo), strConnectionString, intDBMSType);
                    if (AppPermission.IsAllowed(objResult.Data.Tables[0], ((int)AppPermission.PermissionName.REMOVE_WATERMARK).ToString()))
                    {
                        waterMarkOptions = null;
                    }
                }
                var dataToGetDmsId = new
                {
                    templateId = singleDoc.TempId,
                    systemNo = singleDoc.SysNo,
                    strConnectionString
                };
                string dmsId = null;
                try
                {
                    dmsId = await (ConfigurationManager.AppSettings["BusinessApi"] + "Documents/GetDMSBySysAndTempIdsAsync").WithOAuthBearerToken(_token).PostJsonAsync(dataToGetDmsId).ReceiveString();
                }
                catch (Exception ex)
                {
                }

                if (!string.IsNullOrEmpty(dmsId))
                {
                    int pageCount = GetDocumentPageCount(Convert.ToInt32(dmsId.Replace("\"", "")));
                    var dataToPostImage = new
                    {
                        docId = dmsId.Replace("\"", ""),
                        pageCount = pageCount
                    };
                    var IImageBy = await (ConfigurationManager.AppSettings["BusinessApi"] + "Documents/ExportImage").WithOAuthBearerToken(_token).PostJsonAsync(dataToPostImage).ReceiveBytes();
                    Bitmap bitmap = (Bitmap)Image.FromStream(new MemoryStream(IImageBy));
                    int compressionRate = Convert.ToInt32(ConfigurationManager.AppSettings["compressionRate"].ToString());

                    int toPage = pageCount;
                    var fromPage = 1;



                    List<Watermark.Watermark> allWatermarks = new List<Watermark.Watermark>();
                    EDocumentsData obj = new EDocumentsData();

                    try
                    {
                        if (waterMarkOptions != null)
                            // disabling watermarks for now
                            allWatermarks = (await obj.GetResolvedWatermarks(waterMarkOptions, strConnectionString)).ToList();
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf: " + ex.ToString() + Environment.NewLine);
                    }

                    List<DataObjects.TemplateImageAnnotationDD> allAnnotations = new List<DataObjects.TemplateImageAnnotationDD>();
                    if (annotationOptions != null)
                    {
                        try
                        {
                            allAnnotations = (await obj.GetApplicableAnnotationsForDocument(Convert.ToInt32(dmsId.Replace("\"", "")), strConnectionString)).ToList();
                        }
                        catch (Exception ex)
                        {

                        }

                    }

                    allAnnotations.All(c =>
                    {
                        c.LangId = langId;
                        return true;
                    });

                    var lowerEnd = fromPage;
                    var upperEnd = toPage;

                    Dictionary<int, byte[]> pgResult = new Dictionary<int, byte[]>();

                    // For i = fromPage To upperEnd + 1
                    for (var i = fromPage; i <= upperEnd; i++)
                    {
                        var index = i;
                        var creteria = "" + dmsId + "-" + i + "";
                        var annotation = allAnnotations.Where(f => f.DMSNumber.Equals(creteria.Replace("\"", ""))).ToList();
                        try
                        {
                            bitmap.SelectActiveFrame(FrameDimension.Page, index - 1);
                            MemoryStream byteStream = new MemoryStream();
                            bitmap.Save(byteStream, ImageFormat.Tiff);

                            var pResult = ApplyWatermark(Image.FromStream(byteStream), allWatermarks);

                            if (annotation.Count > 0)
                                pResult = (byte[])ApplyAnnotation(pResult, Convert.ToInt32(dmsId.Replace("\"", "")), index - 1, baseFilePath, annotation);

                            pgResult.Add(index - 1, pResult);
                        }
                        catch (Exception ex)
                        {
                            System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "ExportToPdf: " + ex.ToString() + Environment.NewLine);
                        }
                    }

                    if (lowerEnd != fromPage)
                        _doc = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);

                    for (var pageNo = lowerEnd; pageNo <= upperEnd; pageNo++)
                    {
                        using (MemoryStream memStream = new MemoryStream(pgResult[pageNo - 1]))
                        {
                            string path = ConfigurationManager.AppSettings["tempPath"].ToString() + "_1_" + dmsId.Replace("\"", "") + singleDoc.SysNo.ToString() + singleDoc.TempId.ToString() + pageNo.ToString() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".tiff";
                            SaveTiff(path, Image.FromStream(memStream), compressionRate);
                            using (var xImg = XImage.FromFile(path))
                            {
                                // Using xImg = XImage.FromStream(memStream)
                                PdfPage pdfPage = new PdfPage();
                                _doc.AddPage(pdfPage);
                                pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                                pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                                using (var gfx = XGraphics.FromPdfPage(pdfPage))
                                {
                                    gfx.SmoothingMode = XSmoothingMode.HighSpeed;// = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                                    gfx.DrawImage(xImg, 0, 0);
                                }
                            }
                        }
                    }

                    pgResult.Clear();
                    lowerEnd = upperEnd + 1;
                }

            }
            _doc.Save(fileName);
            _doc.Close();
            _doc.Dispose();
            return fName;

        }

        public void SaveTiff(string path, Image img, int quality)
        {
            try
            {
                if (quality < 0 || quality > 100)
                    throw new ArgumentOutOfRangeException("quality must be between 0 and 100.");
                var imgStream = new MemoryStream();
                EncoderParameter qualityParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                ImageCodecInfo jpegCodec = GetEncoderInfo("image/tiff");
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = qualityParam;
                img.Save(path, jpegCodec, encoderParams);
            }

            catch (Exception ex)
            {
                System.IO.File.AppendAllText(ConfigurationManager.AppSettings["errorlog"].ToString(), "SaveTiff: " + ex.ToString() + Environment.NewLine);
                throw ex;
            }
        }


        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

            for (int i = 0; i <= codecs.Length - 1; i++)
            {
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];
            }
            return null/* TODO Change to default(_) if this is not a reference type */;
        }

        public async Task<string> ExportToMultiPdf(int documentId, string options, DataObjects.WatermarkOptions watermarkOptions, DataObjects.AnnotationOptions annotationOptions, string fName, PdfDocument pdfdocument, string pages = "", int langId = 2)
        {
            string fileName = Path.Combine(_basefilePath, fName);
            int toPage;
            int fromPage = 1;
            if (!string.IsNullOrEmpty(pages))
            {
                toPage = int.Parse(pages.Split('-')[1]);
                fromPage = int.Parse(pages.Split('-')[0]);
            }
            else
            {
                toPage = GetDocumentPageCount(_dmsId);
            }

            // Dim service As New EDocuments.EDocumentsSoapClient("EDocumentsSoap")
            var allWatermarks = new List<Watermark.Watermark>();
            var obj = new EDocumentsData();
            if (watermarkOptions is object)
            {

                // disabling watermarks for now
                allWatermarks = (await obj.GetResolvedWatermarks(watermarkOptions, _connectionString)).ToList();
                // allWatermarks = service.GetResolvedWatermarks(watermarkOptions, Me._connectionString).ToList()
            }

            var allAnnotations = new List<DataObjects.TemplateImageAnnotationDD>();
            if (annotationOptions is object)
            {
                allAnnotations = (await obj.GetApplicableAnnotationsForDocument(_dmsId, _connectionString)).ToList();
                // allAnnotations = service.GetApplicableAnnotationsForDocument(_dmsId, _connectionString).ToList()
            }

            allAnnotations.All(c =>
            {
                c.LangId = langId;
                return true;
            });
            // Dim pService As New PermissionService.PermissionsServiceSoapClient("PermissionsServiceSoap")
            var objPermission = new PermissionsData();
            string pServiceResult = await objPermission.GetApplicationSettingByKeyAsync("_APS_PDF_EXPORT_THREAD_THRESHOLD_", _connectionString);
            // Dim pServiceResult = pService.GetApplicationSettingByKey("_APS_PDF_EXPORT_THREAD_THRESHOLD_", Me._connectionString)
            int threshould = int.Parse(Conversions.ToString(Interaction.IIf(string.IsNullOrEmpty(pServiceResult), 5, pServiceResult))) - 1;


            // Dim threshould = 5 - 1
            int lowerEnd = fromPage;
            int upperEnd = int.Parse(Conversions.ToString(Interaction.IIf(lowerEnd + threshould >= toPage, toPage, lowerEnd + threshould)));

            // _document = New PdfDocument()
            _document = pdfdocument;
            do
            {
                // Will Continue work here
                var pgResult = new Dictionary<int, byte[]>();
                for (int i = fromPage, loopTo = upperEnd + 1; i <= loopTo; i++)
                {
                    int index = i;
                    var annotation = allAnnotations.Where(f => Operators.ConditionalCompareObjectEqual(f.DMSNumber, _dmsId.ToString() + "-" + index.ToString(), false)).ToList();
                    var dataToPost = new
                    {
                        dmsId = _dmsId,
                        pageNo = index,
                        options = "",
                        strConnectionString = _connectionString,
                        baseFilePath = _basefilePath,
                        annotations = new List<DataObjects.TemplateImageAnnotationDD>(),
                        resolvedWatermarks = new List<Watermark.Watermark>()
                    };
                    try
                    {
                        var pResult = await DocumentCache.GetDocMainPage(_systemNo, Convert.ToInt32(_templateId), index, _dmsId);
                        pResult = ApplyWatermark(new Bitmap(new MemoryStream(pResult)), allWatermarks);
                        // End If
                        if (annotation.Count > 0)
                        {
                            pResult = (byte[])ApplyAnnotation(pResult, _dmsId, index, _basefilePath, annotation);
                        }
                        // End If

                        pgResult.Add(index, pResult);
                    }
                    catch (Exception ex)
                    {
                        var r = ex;
                    }
                }

                if (lowerEnd != fromPage)
                {
                    _document = PdfReader.Open(fileName, PdfDocumentOpenMode.Modify);
                }

                for (int pageNo = lowerEnd, loopTo1 = upperEnd; pageNo <= loopTo1; pageNo++)
                {
                    using (var memStream = new MemoryStream(pgResult[pageNo]))
                    {
                        using (var xImg = XImage.FromStream(memStream))
                        {
                            var pdfPage = new PdfPage();
                            _document.AddPage(pdfPage);
                            pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                            pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                gfx.DrawImage(xImg, 0d, 0d);
                            }
                        }
                    }
                }

                pgResult.Clear();
                _document.Save(fileName);
                _document.Close();
                _document.Dispose();
                if (upperEnd == toPage)
                {
                    break;
                }

                lowerEnd = upperEnd + 1;
                upperEnd = Convert.ToInt32(Interaction.IIf(lowerEnd + threshould >= toPage, toPage, lowerEnd + threshould));
            }
            while (true);
            return fName;
        }

        private static byte[] CompressImageQuality(Bitmap img)
        {
            try
            {

                // Dim xRes = CSng(img.HorizontalResolution - (img.HorizontalResolution * 0.7))
                // Dim yRes = CSng(img.VerticalResolution - (img.VerticalResolution * 0.7))
                // img.SetResolution(xRes, yRes)

                long compressionRate;
                if (long.TryParse(ConfigurationManager.AppSettings["DOC_IMAGE_COMPRESSION"], out compressionRate))
                {
                    if (compressionRate >= 0L && compressionRate <= 100L)
                    {
                        compressionRate = 100L - compressionRate;
                    }
                    else
                    {
                        compressionRate = 30L;
                    }
                }
                else
                {
                    compressionRate = 30L;
                }

                using (img)
                {
                    var jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                    var myEncoder = Encoder.Quality;
                    var myEncoderParameters = new EncoderParameters(1);
                    var myEncoderParameter = new EncoderParameter(myEncoder, compressionRate);
                    myEncoderParameters.Param[0] = myEncoderParameter;
                    using (var memStream = new MemoryStream())
                    {
                        img.Save(memStream, jpgEncoder, myEncoderParameters);
                        return memStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            var codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }

            return null;
        }

        // Private Function resizeImage(ByVal inputImage As System.Drawing.Image, width As Integer, height As Integer) As System.Drawing.Image
        // Dim maxWidth As Integer = width
        // Dim maxHeight As Integer = height
        // inputImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        // inputImage.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
        // Dim newWidth As Integer = inputImage.Width

        // If inputImage.Width >= maxWidth Then
        // newWidth = maxWidth
        // End If

        // Dim newHeight As Integer = CInt(((CDbl(inputImage.Height) / CDbl(inputImage.Width)) * newWidth))

        // If newHeight > maxHeight Then
        // newHeight = maxHeight
        // newWidth = CInt(((CDbl(inputImage.Width) / CDbl(inputImage.Height)) * maxHeight))
        // End If

        // Dim newImage As System.Drawing.Image = inputImage.GetThumbnailImage(newWidth, newHeight, Nothing, IntPtr.Zero)
        // Dim baseImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(maxWidth, maxHeight)

        // Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(baseImage)
        // Dim x As Integer = CInt((CDbl(((maxWidth - newImage.Width))) / 2))
        // Dim y As Integer = CInt((CDbl(((maxHeight - newImage.Height))) / 2))
        // g.Clear(Drawing.Color.White)
        // g.DrawImage(newImage, New System.Drawing.Rectangle(x, y, newImage.Width, newImage.Height))
        // End Using

        // Return baseImage
        // End Function

        private static Image ResizeImage(Image inputImage, int reducSizePercentage)
        {
            using (var image = inputImage)
            {
                int newWidth = Convert.ToInt32(image.Width * (reducSizePercentage / 100d));
                int newHeight = Convert.ToInt32(image.Height * (reducSizePercentage / 100d));
                var resizeImg = new Bitmap(image, newWidth, newHeight);
                return resizeImg;
            }
        }

        public void PageHelper(int pageNo, int documentId, bool includeAnnotations)
        {

            // Dim service As New EDocuments.EDocumentsSoapClient("EDocumentsSoap")
            using (var image = GetImageWithAnnotationAndWatermark(_dmsId, pageNo, documentId, Convert.ToInt32(_templateId), _basefilePath, _connectionString, includeAnnotations))
            {
                using (var imgStream = new MemoryStream())
                {
                    image.Save(imgStream, ImageFormat.Bmp);
                    var xImg = XImage.FromStream(imgStream);
                    var pdfPage = _pgDict[pageNo];
                    pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                    pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                    using (var gfx = XGraphics.FromPdfPage(pdfPage))
                    {
                        gfx.DrawImage(xImg, 0d, 0d);
                    }
                    xImg.Dispose();
                }
            }
        }

        public static List<DataObjects.TemplateImageAnnotationDD> GetApplicableAnnotationsForPage(int dmsId, int pageNo, string connectionString)
        {

            // Dim service As New EDocuments.EDocumentsSoapClient("EDocumentsSoap")
            string pageDMSNumber = dmsId.ToString() + "-" + pageNo.ToString();
            var obj = new EDocumentsData();
            string strResult = obj.TemplateImagesAnnotationGet(" AND DMSNumber='" + pageDMSNumber + "'", connectionString, 1);
            // Dim strResult = service.TemplateImagesAnnotationGet(" AND DMSNumber='" + pageDMSNumber + "'", connectionString, 1)

            // Will filter annotation here...

            var filteredAnnotations = SharedFunctions.XmlResponseToAnnotationDD(strResult, pageDMSNumber);
            return (List<DataObjects.TemplateImageAnnotationDD>)Interaction.IIf(filteredAnnotations is null, new List<DataObjects.TemplateImageAnnotationDD>(), filteredAnnotations);
        }

        public static int GetDocumentPageCount(int dmsId)
        {
            // Dim service As New EDocuments.EDocumentsSoapClient("EDocumentsSoap")
            var obj = new EDocumentsData();
            int pageCount = obj.GetDMSDocumentPageCount(dmsId);
            // Dim pageCount = service.GetDMSDocumentPageCount(dmsId)
            return pageCount;
        }

        public static Image ConvertAnnotationToImage(DataObjects.TemplateImageAnnotationDD annotation, string basefilePath)
        {
            try
            {
                JObject props = JsonConvert.DeserializeObject<JObject>(annotation.Properties);
                var mainProps = props.Value<JObject>("main");
                var fontProps = props.Value<JObject>("font");
                var bitImage = new Bitmap(int.Parse(annotation.Width.ToString()), int.Parse(annotation.Height.ToString()));
                Color bgColor;
                if (annotation.Type == "highlight")
                {
                    if (mainProps.Value<object>("background-color") is null)
                    {
                        bgColor = Color.FromArgb(255, 255, 0);
                    }
                    else
                    {
                        bgColor = RGBToColor(Conversions.ToString(mainProps.Value<object>("background-color")), Color.FromArgb(255, 255, 0));
                    }

                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        using (var brsh = new SolidBrush(bgColor))
                        {
                            gfx.FillRectangle(brsh, 0, 0, bitImage.Width, bitImage.Height);
                        }
                    }

                    var imgOpcty = new Bitmap(bitImage.Width, bitImage.Height);
                    using (var gfx = Graphics.FromImage(imgOpcty))
                    {
                        var matrix = new ColorMatrix();
                        matrix.Matrix33 = 0.5f;
                        var attributes = new ImageAttributes();
                        attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                        gfx.DrawImage(bitImage, new Rectangle(0, 0, bitImage.Width, bitImage.Height), 0, 0, bitImage.Width, bitImage.Height, GraphicsUnit.Pixel, attributes);
                    }

                    return imgOpcty;
                }
                else if (annotation.Type == "blackout")
                {
                    if (mainProps.Value<object>("background-color") is null)
                    {
                        bgColor = Color.FromArgb(0, 0, 0);
                    }
                    else
                    {
                        bgColor = RGBToColor(Conversions.ToString(mainProps.Value<object>("background-color")), Color.FromArgb(0, 0, 0));
                    }

                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        using (var brsh = new SolidBrush(bgColor))
                        {
                            gfx.FillRectangle(brsh, 0, 0, bitImage.Width, bitImage.Height);
                        }
                    }

                    return bitImage;
                }
                else if (annotation.Type == "text")
                {
                    if (mainProps.Value<object>("background-color") is null)
                    {
                        bgColor = Color.White;
                    }
                    else
                    {
                        bgColor = RGBToColor(Conversions.ToString(mainProps.Value<object>("background-color")), Color.White);
                    }

                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        using (var brsh = new SolidBrush(bgColor))
                        {
                            gfx.FillRectangle(brsh, 0, 0, bitImage.Width, bitImage.Height);
                        }

                        var textBox = new Rectangle(3, 3, bitImage.Width - 6, bitImage.Height - 6);
                        gfx.DrawRectangle(new Pen(Color.Gray), textBox);
                        object fontColor;
                        if (fontProps.Value<object>("color") is null)
                        {
                            fontColor = Color.Black;
                        }
                        else
                        {
                            fontColor = RGBToColor(Conversions.ToString(fontProps.Value<object>("color")), Color.Black);
                        }

                        StringFormat stringFormat;
                        if (annotation.LangId == 1)
                        {
                            stringFormat = new StringFormat();
                        }
                        else
                        {
                            stringFormat = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                        }

                        using (var brsh = new SolidBrush((Color)fontColor))
                        {
                            //Font drawFont = new Font("Arial", 16);
                            // gfx.DrawString(annotation.Content, drawFont, brsh, textBox);
                            try
                            {
                                gfx.DrawString(annotation.Content, PDFExporter.CssFontConvertor(fontProps), brsh, textBox, stringFormat);
                            }
                            catch (Exception ex)
                            {
                                //TODO: Handl error
                            }

                        }
                    }

                    return bitImage;
                }
                else if (annotation.Type == "sticky")
                {
                    bgColor = Color.FromArgb(255, 255, 0);
                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        using (var brsh = new SolidBrush(bgColor))
                        {
                            gfx.FillRectangle(brsh, 0, 0, bitImage.Width, bitImage.Height);
                        }

                        // gfx.DrawRectangle(New Pen(Color.Gray), 10, 25, bitImage.Width - 10, bitImage.Height - 25)
                        var textBox = new Rectangle(10, 25, bitImage.Width - 20, bitImage.Height - 50);
                        object fontColor;
                        if (fontProps.Value<object>("color") is null)
                        {
                            fontColor = Color.Black;
                        }
                        else
                        {
                            fontColor = RGBToColor(Conversions.ToString(fontProps.Value<object>("color")), Color.Black);
                        }

                        using (var brsh = new SolidBrush((Color)fontColor))
                        {
                            gfx.DrawString(annotation.Content, PDFExporter.CssFontConvertor(fontProps), brsh, textBox);
                        }
                    }

                    return bitImage;
                }
                else if (annotation.Type == "freehand")
                {
                    try
                    {
                        bgColor = Color.FromArgb(255, 255, 0);
                        using (var gfx = Graphics.FromImage(bitImage))
                        {
                            var imgFreehand = new Bitmap(new MemoryStream(Convert.FromBase64String(annotation.Content.Split(',')[1])));
                            gfx.DrawImage(imgFreehand, 0, 0, int.Parse(annotation.Width.ToString()), int.Parse(annotation.Height.ToString()));
                        }
                    }
                    catch (Exception freehandEx)
                    {
                        //TODO: Handl error
                        //throw freehandEx;
                    }


                    return bitImage;
                }
                else if (annotation.Type == "signature")
                {
                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        dynamic imgSignature = null;
                        if (File.Exists(Path.Combine(basefilePath, "signatures", annotation.FilePath.Split('?')[0])))
                        {
                            imgSignature = Image.FromFile(Path.Combine(basefilePath, "signatures", annotation.FilePath.Split('?')[0]));
                        }
                        else
                        {
                            imgSignature = new Bitmap(new MemoryStream(Convert.FromBase64String(annotation.Content.Split(',')[1])));
                        }

                        gfx.DrawImage(imgSignature, 0, 0, int.Parse(annotation.Width.ToString()), int.Parse(annotation.Height.ToString()));
                    }

                    return bitImage;
                }
                else if (annotation.Type == "barcode")
                {
                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        using (var brsh = new SolidBrush(Color.White))
                        {
                            gfx.FillRectangle(brsh, 0, 0, bitImage.Width, bitImage.Height);
                        }

                        dynamic imgBarcode = null;
                        if (File.Exists(Path.Combine(basefilePath, "barcodes", annotation.FilePath.Split('?')[0])))
                        {
                            imgBarcode = Image.FromFile(Path.Combine(basefilePath, "barcodes", annotation.FilePath.Split('?')[0]));
                        }
                        else
                        {
                            imgBarcode = new Bitmap(new MemoryStream(Convert.FromBase64String(annotation.Content.Split(',')[1])));
                        }

                        gfx.DrawImage(imgBarcode, 0, 0, int.Parse(annotation.Width.ToString()), int.Parse(annotation.Height.ToString()));
                    }

                    return bitImage;
                }
                else if (annotation.Type == "stamp")
                {
                    bgColor = Color.FromArgb(255, 255, 0);
                    using (var gfx = Graphics.FromImage(bitImage))
                    {
                        var imgFreehand = new Bitmap(new MemoryStream(Convert.FromBase64String(annotation.Content.Split(',')[1])));
                        gfx.DrawImage(imgFreehand, 0, 0, int.Parse(annotation.Width.ToString()), int.Parse(annotation.Height.ToString()));
                    }

                    return bitImage;
                }
            }
            catch (Exception ex)
            {
            }

            return null;
        }

        public static object GetProperty(object o, string member)
        {
            if (o == null) throw new ArgumentNullException("o");
            if (member == null) throw new ArgumentNullException("member");
            Type scope = o.GetType();
            IDynamicMetaObjectProvider provider = o as IDynamicMetaObjectProvider;
            if (provider != null)
            {
                ParameterExpression param = Expression.Parameter(typeof(object));
                DynamicMetaObject mobj = provider.GetMetaObject(param);
                GetMemberBinder binder = (GetMemberBinder)Microsoft.CSharp.RuntimeBinder.Binder.GetMember(0, member, scope, new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(0, null) });
                DynamicMetaObject ret = mobj.BindGetMember(binder);
                BlockExpression final = Expression.Block(
                    Expression.Label(CallSiteBinder.UpdateLabel),
                    ret.Expression
                );
                LambdaExpression lambda = Expression.Lambda(final, param);
                Delegate del = lambda.Compile();
                return del.DynamicInvoke(o);
            }
            else
            {
                return o.GetType().GetProperty(member, BindingFlags.Public | BindingFlags.Instance).GetValue(o, null);
            }
        }

        private static Color RGBToColor(string rgb, Color defaultColor)
        {
            try
            {
                var rgbArr = rgb.Replace("rgb(", "").Replace(")", "").Split(',').Select(x => int.Parse(x.Trim())).ToArray();
                return Color.FromArgb(rgbArr[0], rgbArr[1], rgbArr[2]);
            }
            catch (Exception ex)
            {
                return defaultColor;
            }
        }

        private static Font CssFontConvertor(dynamic fontProps)
        {
            string fontFamily = "Times New Roman";
            var _fontFamily = fontProps["font-family"]?.ToString().Replace("\"", "");
            if (_fontFamily != null)
            {
                fontFamily = Convert.ToString(_fontFamily);
            }

            int fontSize = 16;
            if (fontProps["font-size"] is object)
            {
                fontSize = (int)Math.Round(float.Parse(fontProps["font-size"]?.ToString().Replace("px", "")));
            }

            var boldStyle = FontStyle.Regular;
            if (fontProps["font-weight"] is object)
            {
                if (fontProps["font-weight"]?.ToString() != "normal")
                {
                    boldStyle = FontStyle.Bold;
                }
            }

            var italicStyle = FontStyle.Regular;
            if (fontProps["font-style"] is object)
            {
                if (fontProps["font-style"]?.ToString() != "normal")
                {
                    italicStyle = FontStyle.Italic;
                }
            }

            return new Font(fontFamily, fontSize, boldStyle | italicStyle);
        }

        public static Image GetImageWithAnnotationAndWatermark(int dmsId, int pageNo, int documentId, int templateId, string basefilePath, string strConnectionString, bool includeAnnotations = true, bool applyWatermark = true, string annotationsToInclde = "All")
        {
            // Dim service As New EDocuments.EDocumentsSoapClient("EDocumentsSoap")
            var obj = new EDocumentsData();
            Image image;
            if (applyWatermark)
            {
                image = SharedFunctions.ApplyWatermark(new Bitmap(new MemoryStream(obj.DownloadDMSFile(dmsId.ToString(), pageNo.ToString(), ""))), documentId.ToString(), templateId.ToString());
            }
            // image = SharedFunctions.ApplyWatermark(New Bitmap(New MemoryStream(service.DownloadDMSFile(dmsId, pageNo, ""))), documentId, templateId)
            else
            {
                image = new Bitmap(new MemoryStream(obj.DownloadDMSFile(dmsId.ToString(), pageNo.ToString(), "")));
                // image = New Bitmap(New MemoryStream(service.DownloadDMSFile(dmsId, pageNo, "")))
            }

            if (includeAnnotations)
            {
                var annotations = GetApplicableAnnotationsForPage(dmsId, pageNo, strConnectionString);
                foreach (var annotation in annotations)
                {
                    if (annotationsToInclde != "All" && !annotationsToInclde.Contains(annotation.Type))
                        continue;
                    using (var imgAnnotation = ConvertAnnotationToImage(annotation, basefilePath))
                    {
                        if (imgAnnotation is object)
                        {

                            int ScalePercentage = int.Parse(ConfigurationManager.AppSettings["ScalePercentage"].ToString());
                            int intX, intY;
                            intX = (imgAnnotation.Width / ScalePercentage);
                            intY = (imgAnnotation.Height / ScalePercentage);
                            int top = (int)(annotation.Top / ScalePercentage);
                            int left = (int)(annotation.Left / ScalePercentage);
                            Bitmap bm = new Bitmap(intX, intY);
                            Graphics g = Graphics.FromImage(bm);
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                            g.DrawImage(imgAnnotation, left, top, intX, intY);
                            image = bm;


                            // int ScalePercentage = int.Parse(ConfigurationManager.AppSettings["ScalePercentage"].ToString());                           
                            // int intX = (image.Width / ScalePercentage);
                            // int intY = (image.Height / ScalePercentage);
                            // int top = (int)(annotation.Top / ScalePercentage);
                            // int left = (int)(annotation.Left / ScalePercentage);                                                   
                            // using (var g = Graphics.FromImage(image))
                            // {                               
                            //     g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                            //     g.DrawImage(imgAnnotation, left, top, intX, intY);
                            //    // g.DrawImage(imgAnnotation, int.Parse(annotation.Left.ToString()), int.Parse(annotation.Top.ToString()));
                            // }
                        }
                    }
                }
            }
            return image;
        }

        public object ApplyAnnotation(byte[] imgBytes, int dmsId, int pageNo, string basefilePath, List<DataObjects.TemplateImageAnnotationDD> annotations, bool? compress = true)
        {
            var image = new Bitmap(new MemoryStream(imgBytes));
            int ScalePercentage = int.Parse(ConfigurationManager.AppSettings["ScalePercentage"].ToString());
            foreach (var annotation in annotations)
            {
                using (var imgAnnotation = ConvertAnnotationToImage(annotation, basefilePath))
                {
                    if (imgAnnotation is object)
                    {


                        // int ScalePercentage = int.Parse(ConfigurationManager.AppSettings["ScalePercentage"].ToString());
                        // int intX, intY;
                        // intX = (imgAnnotation.Width / ScalePercentage);
                        // intY = (imgAnnotation.Height / ScalePercentage);
                        // int top = (int)(annotation.Top / ScalePercentage);
                        // int left = (int)(annotation.Left / ScalePercentage);
                        // Bitmap bm = new Bitmap(intX, intY);
                        // Graphics g = Graphics.FromImage(bm);
                        // g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                        // g.DrawImage(imgAnnotation, left, top, intX, intY);
                        // image = bm;




                        int annotationWidth = (int)(annotation.Width / ScalePercentage);
                        int annotationHeight = (int)(annotation.Height / ScalePercentage);
                        int top = (int)(annotation.Top / ScalePercentage);
                        int left = (int)(annotation.Left / ScalePercentage);
                        using (var g = Graphics.FromImage(image))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                            g.DrawImage(imgAnnotation, left, top, annotationWidth, annotationHeight);
                            //g.DrawImage(imgAnnotation, int.Parse(annotation.Left.ToString()), int.Parse(annotation.Top.ToString()));
                        }
                    }
                }
            }


            if (compress.GetValueOrDefault(true))
            {
                int intX, intY;
                intX = (image.Width / ScalePercentage);
                intY = (image.Height / ScalePercentage);
                Bitmap bm = new Bitmap(intX, intY);
                Graphics mg = Graphics.FromImage(bm);
                mg.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                mg.DrawImage(image, 0, 0, intX, intY);
                image = bm;
            }

            using (var stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Jpeg);
                return stream.ToArray();
            }

            return null;
        }

        public static byte[] ApplyWatermark(Image image, List<Watermark.Watermark> allWatermarks, bool? compress = true)
        {
            using (image)
            {
                var objWatermarker = new Watermark.Watermarker(image);
                if (allWatermarks is object)
                {
                    foreach (var wMark in allWatermarks)
                    {
                        if (wMark.Type == "Repeating" || wMark.Type == "تكرار")
                        {
                            objWatermarker.FontSize = wMark.Font.Size;
                        }
                        else
                        {
                            objWatermarker.SpanPercent = wMark.SpanPercent;
                            objWatermarker.FontSize = wMark.Font.Size;
                            objWatermarker.Angle = wMark.Angle;
                            objWatermarker.PositionHorizontal = wMark.Horizontal;
                            objWatermarker.PositionVertical = wMark.Vertical;
                        }

                        objWatermarker.FontStyleName = "Bold";
                        objWatermarker.FontName = "Verdana";
                        objWatermarker.Opacity = (float)(wMark.Darkness / 100d);
                        objWatermarker.TransparentColor = Color.Transparent;
                        if (wMark.Type == "Repeating" || wMark.Type == "تكرار")
                        {
                            objWatermarker.DrawText(wMark.Text);
                        }
                        else
                        {
                            objWatermarker.DrawTextDiagonal(wMark.Text);
                        }
                    }
                }

                if (compress.GetValueOrDefault(true))
                {
                    int ScalePercentage = int.Parse(ConfigurationManager.AppSettings["ScalePercentage"].ToString());
                    int intX, intY;
                    intX = (objWatermarker.Image.Width / ScalePercentage);
                    intY = (objWatermarker.Image.Height / ScalePercentage);
                    Bitmap bm = new Bitmap(intX, intY);
                    Graphics g = Graphics.FromImage(bm);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                    g.DrawImage(objWatermarker.Image, 0, 0, intX, intY);
                    image = bm;
                }
                else
                {
                    image = objWatermarker.Image;
                }



            }

            using (var stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Jpeg);
                return stream.ToArray();
            }

            return null;
        }
        public static string GetUserIPAddress()
        {
            string objVal = "";
            try
            {
                objVal = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
                if (string.IsNullOrEmpty(objVal))
                {
                    objVal = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
                }

                if (objVal == "::1")
                {
                    objVal = "127.0.0.1";
                }
            }
            catch (Exception ex)
            {
            }

            return objVal;
        }
    }
}