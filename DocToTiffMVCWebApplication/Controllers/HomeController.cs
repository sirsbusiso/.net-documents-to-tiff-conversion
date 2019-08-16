using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocToTiffMVCWebApplication;
using Spire.Doc;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing;
using Spire.Pdf;

namespace DocToTiffMVCWebApplication.Controllers
{
    public class HomeController : Controller
    {
        Logic logic = new Logic();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            return View();
        }


        [HttpGet]
        public ActionResult Convert()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Convert(HttpPostedFileBase files)
        {
            var fileext = System.IO.Path.GetExtension(files.FileName).Substring(1);
            if (fileext == "doc")
            {
                MemoryStream target = new MemoryStream();
                files.InputStream.CopyTo(target);
                Document document = new Document(target);

                //Document document = new Document(@"C:\Users\User1\Downloads\TestWordDoc.doc");

                var file = logic.SaveAsImageDOC(document);

                logic.JoinTiffImages(file, Server.MapPath(@"~/" + files.FileName + ".tiff"), EncoderValue.CompressionLZW);

                System.Diagnostics.Process.Start(Server.MapPath(@"~/" + files.FileName + ".tiff"));
                return new EmptyResult();

                //ViewBag.Message = "Your contact page.";

                //return View();
            }
            else if (fileext == "jpg" || fileext == "png")
            {
                MemoryStream target = new MemoryStream();
                files.InputStream.CopyTo(target);
                Image imgage = Image.FromStream(target);

                Image[] imgs = new Image[1];
                imgs[0] = imgage;
                
                logic.JoinTiffImages(imgs, Server.MapPath(@"~/Content/Documents/" + files.FileName + ".tiff"), EncoderValue.CompressionLZW);
                var path = Server.MapPath(@"~/Content/Documents/" + files.FileName + ".tiff");
                System.Diagnostics.Process.Start(Server.MapPath(@"~/Content/Documents/" + files.FileName + ".tiff"));

                return new EmptyResult();
            }
            else if (fileext == "pdf")
            {
                MemoryStream target = new MemoryStream();
                files.InputStream.CopyTo(target);
                PdfDocument document = new PdfDocument(target);

                var file = logic.SaveAsImagePDF(document);

                //Call JoinTiffImages() method to save the PDF to Tiff and open the result 
                logic.JoinTiffImages(file, Server.MapPath(@"~/" + files.FileName + ".tiff"), EncoderValue.CompressionLZW);
                System.Diagnostics.Process.Start(Server.MapPath(@"~/" + files.FileName + ".tiff"));


            }
            return View();
        }
    

        public ActionResult Contact(HttpPostedFileBase files)
        {

            MemoryStream target = new MemoryStream();
            files.InputStream.CopyTo(target);
            Document doc = new Document(target);

            Document document = new Document(@"C:\Users\User1\Downloads\TestWordDoc.doc");

            var file = logic.SaveAsImageDOC(document);

            logic.JoinTiffImages(file, Server.MapPath(@"~/" + "result.tiff"), EncoderValue.CompressionLZW);
            System.Diagnostics.Process.Start(Server.MapPath("~/App_Data/Documents"),
                                               Path.GetFileName("result.tiff"));
            System.Diagnostics.Process.Start(Server.MapPath(@"~/" + "result.tiff"));
            return new EmptyResult();

            //ViewBag.Message = "Your contact page.";

            //return View();
        }
    }
}