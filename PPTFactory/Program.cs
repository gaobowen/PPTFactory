using System;
using System.IO;
using PPTAnalysisCore;
using System.Linq;
using System.Drawing;
//using DocumentFormat.OpenXml.Drawing;

namespace PPTFactory
{
    class Program
    {
        static void Main(string[] args)
        {
            // using (PresentationDocument presentationDocument =
            //     PresentationDocument.Open(Directory.GetCurrentDirectory()+"/test.pptx", true))
            // {
            //     var ids = presentationDocument.PresentationPart.Presentation.SlideIdList;
            //     var sldprt = (SlidePart)presentationDocument.PresentationPart
            //         .GetPartById((ids.ChildElements[0] as SlideId).RelationshipId);
            //     var tree = sldprt.Slide.CommonSlideData.ShapeTree;
            // }

            //TestAddPicture();
            //TestAddNewText();
            //TestAddNewSlide();
        }



        static void TestAddPicture()
        {
            var analysisCore = new AnalysisCore();
            Bounding bounding = new Bounding()
            {
                X = 960,
                Y = 540,
                Width = 384,
                Height = 384,
                Rotation = 45
            };
            var transform2D = analysisCore.CreateTransform2D(new Size(1920, 1080), bounding);
            analysisCore.AddPicture(0, AppDomain.CurrentDomain.BaseDirectory + "Image/test.png", transform2D);

            var path = AppDomain.CurrentDomain.BaseDirectory + "/addNewPicture.pptx";
            var ret = analysisCore.Doc.SaveAs(path);

            ret.Close();
            ret.Dispose();

            analysisCore.Dispose();
        }

        //添加场景页测试
        static void TestAddNewSlide()
        {
            AnalysisCore analysisCore = new AnalysisCore();
            var sldpart = analysisCore.AddNewSlide();
            Bounding bounding = new Bounding()
            {
                X = 1536,
                Y = 864,
                Width = 384,
                Height = 108,
                Rotation = 45
            };
            var transform2D = analysisCore.CreateTransform2D(new Size(1920, 1080), bounding);
            PPTTextStyle textStyle = new PPTTextStyle();
            analysisCore.AddText(sldpart, "第二个场景页", textStyle, transform2D);
            analysisCore.Doc.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/addNewSlide.pptx");
        }

        static void TestAddNewText()
        {
            AnalysisCore analysisCore = new AnalysisCore();
            var sldpart = analysisCore.Doc.GetSlidePart(0);
            Bounding bounding = new Bounding()
            {
                X = 192,
                Y = 108,
                Width = 384,
                Height = 108,
                Rotation = 90
            };
            var transform2D = analysisCore.CreateTransform2D(new Size(1920, 1080), bounding);
            PPTTextStyle textStyle = new PPTTextStyle()
            {
                Color = "#FF0000",
                IsUnderline = true,
                IsBold = true,
                IsItalic = true
            };
            analysisCore.AddText(sldpart, "第1个场景页", textStyle, transform2D);
            analysisCore.Doc.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/addNewText.pptx");
        }


        //测试创建空的ppt
        static void TestCreateBlankPPT()
        {
            var newpath = AppDomain.CurrentDomain.BaseDirectory + "/newone.pptx";
            AnalysisHelper.CreateBlankPPT(newpath);
        }


        private static void FixPowerpoint(string fileName)
        {
            Console.WriteLine(fileName);
            using (System.IO.Packaging.Package wdPackage = System.IO.Packaging.Package.Open(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                var binPartUri = new Uri("/ppt/printerSettings/printerSettings1.bin", UriKind.Relative);
                if (wdPackage.PartExists(binPartUri))
                {
                    var presPartUri = new Uri("/ppt/presentation.xml", UriKind.RelativeOrAbsolute);
                    var presPart = wdPackage.GetPart(presPartUri);
                    var presentationPartRels =
                        presPart.GetRelationships()
                            .Where(a =>
                                a.RelationshipType.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
                                StringComparison.InvariantCultureIgnoreCase)).SingleOrDefault();
                    if (presentationPartRels != null)
                    {
                        presPart.DeleteRelationship(presentationPartRels.Id);
                    }
                    wdPackage.DeletePart(binPartUri);
                }
                wdPackage.Close();
            }
        }
    }
}
