using iTextSharp.text.pdf;
using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.util;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApplication1.Helpers;
using WebApplication1.Models;


namespace WebApplication1
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnAddWaterMark_Click(object sender, EventArgs e)
        {
            PdfWatermarkHelper.AddWatermarkTextC("D:\\test2.pdf", "D:\\watermark2.pdf", "66000634");
        }

        protected void btnPdfToImages_Click(object sender, EventArgs e)
        {
            //PdfImageExtractor.ExtractImages("D:\\watermark.pdf");

            //PdfUtils.ConvertPdfToJpg("D:\\watermark2.pdf", "D:\\new-watermark2.pdf");

            // Path to the PDF file
            string pdfFile = "D:\\watermark2.pdf";

            // Output image file path
            string outputFolder = "D:\\_watermark_pdf";

            // Convert PDF to Image
            var pdfImageList = GetPdfImages(pdfFile);

            var result = PdfUtils.GetPdfFromImageList(pdfImageList);

            string path = Path.Combine(outputFolder, "new-watermark2.pdf");

            File.WriteAllBytes(path, result);

            Console.WriteLine("PDF converted to image successfully.");
        }

        private List<ImageFileInfo> GetPdfImages(string pdfFile)
        {
            List<ImageFileInfo> result = new List<ImageFileInfo>();
            var pdfBytes = File.ReadAllBytes(pdfFile);

            using (var memoryStream = new MemoryStream(pdfBytes))
            {
                using (var document = PdfiumViewer.PdfDocument.Load(memoryStream))
                {
                    for (int i = 0; i < document.PageCount; i++)
                    {
                        float pageWidth = document.PageSizes[i].Width;
                        float pageHeight = document.PageSizes[i].Height;

                        int imageWidth = (int)(pageWidth * 300 / 72);
                        int imageHeight = (int)(pageHeight * 300 / 72);

                        using (var image = document.Render(i, imageWidth, imageHeight, 700 , 700, true))
                        {
                            using (var imageStream = new MemoryStream())
                            {
                                image.Save(imageStream, ImageFormat.Jpeg);
                                result.Add(new ImageFileInfo
                                {
                                    ImageBytes = imageStream.ToArray(),
                                    Width = document.PageSizes[i].Width,
                                    Height = document.PageSizes[i].Height
                                });

                                image.Save($"D:\\_watermark_pdf\\{i + 1}.jpeg", ImageFormat.Jpeg);
                            }
                        }
                    }
                }
            }

            return result;
        }
    }
}