using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using ImageMagick;
using PdfSharp.Pdf.Advanced;
using System.IO;
using PdfSharp.Drawing;
using System.Drawing;
using iTextSharp.text;
using WebApplication1.Models;

namespace WebApplication1.Helpers
{
    public class PdfUtils
    {
        public static byte[] GetPdfFromImageList(List<ImageFileInfo> images)
        {
            PdfDocument finalDocument = new PdfDocument();
            foreach (var image in images)
            {
                using (MemoryStream memStream = new MemoryStream(image.ImageBytes))
                {
                    using (var xImg = XImage.FromStream(memStream))
                    {
                        PdfPage pdfPage = finalDocument.AddPage();
                        pdfPage.Width = image.Width;
                        pdfPage.Height = image.Height;
                        using (var gfx = XGraphics.FromPdfPage(pdfPage))
                        {
                            gfx.DrawImage(xImg, 0, 0, image.Width, image.Height);
                        }
                    }
                }
            }

            using (MemoryStream memStream = new MemoryStream())
            {
                finalDocument.Save(memStream);

                return memStream.ToArray();
            }
        }

        public static void ConvertPdfToJpg(string pdfPath, string outputPath)
        {
            using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
            {
                int pageindex = 0;
                PdfDocument finalDocument = new PdfDocument();
                PdfDocument tempDocument;
                foreach (PdfPage page in document.Pages)
                {
                    using (MemoryStream memStream = new MemoryStream())
                    {
                        tempDocument = new PdfDocument();
                        tempDocument.AddPage(page);
                        tempDocument.Save(memStream);

                        using (var xImg = XImage.FromStream(memStream))
                        {
                            PdfPage pdfPage = new PdfPage();
                            finalDocument.AddPage(pdfPage);
                            pdfPage.Width = XUnit.FromPoint(xImg.PointWidth);
                            pdfPage.Height = XUnit.FromPoint(xImg.PointHeight);
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                gfx.DrawImage(xImg, 0, 0);
                            }
                        }
                    }

                    //using (MagickImage magickImage = new MagickImage(pdfPath, MagickFormat.Pdf))
                    //{
                    //    magickImage.Write($"D:\\page_{pageindex + 1}.png", MagickFormat.Jpeg);

                    //}
                    //pageindex++;
                    //using (var image = new MagickImage(page))
                    //{
                    //    image.Format = MagickFormat.Jpg;
                    //    image.Write($"{outputPath}/Page{pageindex++}.jpg");
                    //}
                }

                finalDocument.Save(outputPath);
            }
        }

        public static void ConvertPdfToJpg()
        {
            const string filename = "D:\\watermark.pdf";

            PdfDocument document = PdfReader.Open(filename);

            int imageCount = 0;
            // Iterate pages
            foreach (PdfPage page in document.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        // Iterate references to external objects
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                // Is external object an image?
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    ExportImage(xObject, ref imageCount);
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine(imageCount + " images exported.", "Export Images");
        }

        static void ExportImage(PdfDictionary image, ref int count)
        {
            string filter = image.Elements.GetName("/Filter");
            switch (filter)
            {
                case "/DCTDecode":
                    ExportJpegImage(image, ref count);
                    break;

                case "/FlateDecode":
                    ExportAsPngImage(image, ref count);
                    break;
            }
        }

        static void ExportJpegImage(PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;
            FileStream fs = new FileStream(String.Format("Image{0}.jpeg", count++), FileMode.Create, FileAccess.Write);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(stream);
            bw.Close();
        }

        static void ExportAsPngImage(PdfDictionary image, ref int count)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);
            int bitsPerComponent = image.Elements.GetInteger(PdfImage.Keys.BitsPerComponent);

            // TODO: You can put the code here that converts vom PDF internal image format to a Windows bitmap
            // and use GDI+ to save it in PNG format.
            // It is the work of a day or two for the most important formats. Take a look at the file
            // PdfSharp.Pdf.Advanced/PdfImage.cs to see how we create the PDF image formats.
            // We don't need that feature at the moment and therefore will not implement it.
            // If you write the code for exporting images I would be pleased to publish it in a future release
            // of PDFsharp.
        }
    }
}