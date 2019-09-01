using System;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using System.IO;
using PPTAnalysisCore;
using System.Linq;
//using DocumentFormat.OpenXml.Drawing;

namespace PPTFactory
{
    class Program
    {
        static void Main(string[] args)
        {
            // using (PresentationDocument presentationDocument =
            //     PresentationDocument.Open(Directory.GetCurrentDirectory()+"/testsvg.pptx", true))
            // {
            //     var ids = presentationDocument.PresentationPart.Presentation.SlideIdList;
            //     var sldprt = (SlidePart)presentationDocument.PresentationPart
            //         .GetPartById((ids.ChildElements[0] as SlideId).RelationshipId);
            //     var tree = sldprt.Slide.CommonSlideData.ShapeTree;
            // }

            //TestAddPicture();
            //TestAddNewText();
            TestAddNewSlide();
        }

        private static void FixPowerpoint(string fileName)
        {
            //Opening the package associated with file
            Console.WriteLine(fileName);
            using (System.IO.Packaging.Package wdPackage = System.IO.Packaging.Package.Open(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                //Uri of the printer settings part
                var binPartUri = new Uri("/ppt/printerSettings/printerSettings1.bin", UriKind.Relative);
                if (wdPackage.PartExists(binPartUri))
                {
                    //Uri of the presentation part which contains the relationship
                    var presPartUri = new Uri("/ppt/presentation.xml", UriKind.RelativeOrAbsolute);
                    var presPart = wdPackage.GetPart(presPartUri);
                    //Getting the relationship from the URI
                    var presentationPartRels =
                        presPart.GetRelationships()
                            .Where(a =>
                                a.RelationshipType.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
                                StringComparison.InvariantCultureIgnoreCase)).SingleOrDefault();
                    if (presentationPartRels != null)
                    {
                        //Delete the relationship
                        presPart.DeleteRelationship(presentationPartRels.Id);
                    }

                    //Delete the part
                    wdPackage.DeletePart(binPartUri);
                }
                wdPackage.Close();
            }
        }

        static void TestAddPicture()
        {
            var analysisCore = new AnalysisCore();
            var transform2D = new D.Transform2D()
            {
                Offset = new Drawing.Offset() { X = (Int64)(analysisCore.Width * 0.5), Y = (Int64)(analysisCore.Height * 0.5) },
                Extents = new Drawing.Extents() { Cx = (Int64)(analysisCore.Width * 0.2), Cy = (Int64)(analysisCore.Width * 0.2) },
                Rotation = 45 * 60000,
            };
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
            var transform2D = new D.Transform2D()
            {
                Offset = new Drawing.Offset() { X = (Int64)(analysisCore.Width * 0.8), Y = (Int64)(analysisCore.Height * 0.8) },
                Extents = new Drawing.Extents() { Cx = (Int64)(analysisCore.Width * 0.2), Cy = (Int64)(analysisCore.Height * 0.1) },
                Rotation = 45 * 60000,
            };
            PPTTextStyle textStyle = new PPTTextStyle();
            analysisCore.AddText(sldpart, "第二个场景页", textStyle, transform2D);
            analysisCore.Doc.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "/addNewSlide.pptx");
        }

        static void TestAddNewText()
        {
            AnalysisCore analysisCore = new AnalysisCore();           
            var sldpart = analysisCore.Doc.GetSlidePart(0);
            var transform2D = new D.Transform2D()
            {
                Offset = new Drawing.Offset() { X = (Int64)(analysisCore.Width * 0.1), Y = (Int64)(analysisCore.Height * 0.1) },
                Extents = new Drawing.Extents() { Cx = (Int64)(analysisCore.Width * 0.2), Cy = (Int64)(analysisCore.Height * 0.1) },
                Rotation = 90 * 60000,
            };
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
    }
}
