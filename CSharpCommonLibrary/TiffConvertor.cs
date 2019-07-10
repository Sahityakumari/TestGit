using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace CSharpComLibrary
{
    [Guid("02FDD9A7-8AEF-4CF8-925B-69080DDF68A3"),
 InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ComClas2Events
    {
    }
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ITiffConvertor
    {
        [DispId(28)]
        System.Drawing.Image Base64ToTiff(String base64Code);
        [DispId(29)]
        void TiffToPdf(String path);
    }

    [Guid("C71BE6D8-D746-43F6-B09D-7B85D36852E7"),
    ClassInterface(ClassInterfaceType.None),
       ComSourceInterfaces(typeof(ComClas2Events))]
    [ProgId("CSharpComLibrary.TiffConvertor")]
    public class TiffConvertor : ITiffConvertor
    {
        [ComVisible(true)]
        public System.Drawing.Image Base64ToTiff(String base64Code)
        {
            Byte[] imageBytes = Convert.FromBase64String(base64Code);
            MemoryStream ms = new MemoryStream(imageBytes, 0, imageBytes.Length);
            System.Drawing.Image tiffimg = System.Drawing.Image.FromStream(ms, true);
            tiffimg.Save("Base64ToTiff.tiff", ImageFormat.Tiff);
            return tiffimg;

        }
        [ComVisible(true)]
        public void TiffToPdf(String inputpath)
        {
            string CPdfFiles = "";
            if (!Directory.Exists("C:\\TiffToPdfFolder"))
            {
                Directory.CreateDirectory("C:\\TiffToPdfFolder\\");
                CPdfFiles = "C:\\TiffToPdfFolder\\";
            }
            else
            {
                CPdfFiles = "C:\\TiffToPdfFolder\\";
            }
            string sTiffFiles = inputpath;
            string[] files = Directory.GetFiles(sTiffFiles, "*.tif");
            CPdfFiles = CPdfFiles + Path.GetFileName(files[0]) + ".PDF";
            Document document = new Document(PageSize.A4, 50, 50, 50, 50);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(CPdfFiles, FileMode.CreateNew));
            document.Open();
            AddTiff2Pdf(files, ref writer, ref document);
            document.Close();
        }
        private void AddTiff2Pdf(string[] tiffFileName, ref PdfWriter writer, ref Document document)
        {
            foreach (var filename in tiffFileName)
            {
                Bitmap bitmap = new Bitmap(filename);
                int numberOfPages = bitmap.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page);
                PdfContentByte cb = writer.DirectContent;
                for (int page = 0; page < numberOfPages; page++)
                {
                    bitmap.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, page);
                    System.IO.MemoryStream stream = new System.IO.MemoryStream();
                    bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(stream.ToArray());
                    stream.Close();
                    img.ScalePercent(72f / bitmap.HorizontalResolution * 100);
                    img.SetAbsolutePosition(0, 0);
                    cb.AddImage(img);
                    document.NewPage();
                }
            }

        }
    }
}
   
  



