using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using System.IO;

namespace PPTAnalysisCore
{
    public class AnalysisCore : IDisposable
    {
        private static PresentationDocument _prototype16x9 = null;
        private static P.Slide _prototypeSlide = null;

        private int _width;
        private int _height;

        public PresentationDocument Doc { get; }
        public int Width => _width;
        public int Height => _height;

        public AnalysisCore()
        {
            if (_prototype16x9 == null)
            {
                _prototype16x9 = PresentationDocument.Open(AppDomain.CurrentDomain.BaseDirectory + "/blankone.pptx", true);
            }
            Doc = (PresentationDocument)_prototype16x9.Clone();
            InitProperty();
        }

        public AnalysisCore(PresentationDocument presentationDocument)
        {
            Doc = presentationDocument;
            InitProperty();
        }

        public AnalysisCore(string filePath)
        {
            Doc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation, true);
            InitProperty();
        }

        public int GetSlideCount()
        {
            var ids = Doc.PresentationPart.Presentation.SlideIdList;
            return ids.ChildElements.Count;
        }


        public SlidePart AddNewSlide()
        {
            Slide slide = new Slide(
                new CommonSlideData(new ShapeTree()),
                new ColorMapOverride(new D.MasterColorMapping())
                );
            NonVisualGroupShapeProperties nonVisualProperties =
                slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
            nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = 1, Name = "" };
            nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
            nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
            slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());
            SlidePart slidePart = Doc.PresentationPart.AddNewPart<SlidePart>();
            var slideIdList = Doc.PresentationPart.Presentation.SlideIdList;
            slide.Save(slidePart);
            uint maxSlideId = 1;
            SlideId prevSlideId = null;
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                    prevSlideId = slideId;
                }

            }
            maxSlideId++;
            SlidePart lastSlidePart;
            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)Doc.PresentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)Doc.PresentationPart
                    .GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }
            if (null != lastSlidePart.SlideLayoutPart)
            {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = Doc.PresentationPart.GetIdOfPart(slidePart);
            Doc.PresentationPart.Presentation.Save();
            return slidePart;
        }

        public Shape AddText(int sldIdx, string text, PPTTextStyle textStyle, D.Transform2D transform)
        {
            var tree = Doc.GetShapeTree(sldIdx);
            return AddText(tree, text, textStyle, transform);
        }

        public Shape AddText(SlidePart sldpart, string text, PPTTextStyle textStyle, D.Transform2D transform)
        {
            return AddText(sldpart.Slide.CommonSlideData.ShapeTree, text, textStyle, transform);
        }

        /// <summary>
        /// 添加文本
        /// </summary>
        private Shape AddText(ShapeTree tree, string text, PPTTextStyle textStyle, D.Transform2D transform)
        {
            //Picture
            Shape textShape = tree.AppendChild(new Shape());
            uint maxid = AnalysisHelper.GetMaxId(tree);
            textShape.NonVisualShapeProperties = new P.NonVisualShapeProperties
            (
                new P.NonVisualDrawingProperties()
                {
                    Id = maxid + 1,
                    Name = $"TEXT{maxid + 1}"
                },
                new P.NonVisualShapeDrawingProperties
                (
                    new D.ShapeLocks()
                    {
                        NoGrouping = true
                    }
                ),
                new ApplicationNonVisualDrawingProperties
                (
                    new PlaceholderShape()
                    {
                        Type = PlaceholderValues.Body
                    }
                )
            );
            //位置
            textShape.ShapeProperties = new ShapeProperties()
            {
                Transform2D = transform
            };

            textShape.TextBody = new TextBody(
                new D.BodyProperties(),
                new D.ListStyle(),
                new D.Paragraph(new D.Run()
                {
                    Text = new D.Text() { Text = text },
                    RunProperties = new D.RunProperties
                    (
                        new D.SolidFill()
                        {
                            RgbColorModelHex = new D.RgbColorModelHex()
                            {
                                Val = HexBinaryValue.FromString(textStyle.Color.Replace("#", ""))
                            }
                        }
                    )
                    {
                        FontSize = (int)(textStyle.FontSize * 100), //20*100 字号20
                        Underline = textStyle.IsUnderline ? D.TextUnderlineValues.Single : D.TextUnderlineValues.None,
                        Italic = textStyle.IsItalic,
                        Bold = textStyle.IsBold,
                        AlternativeLanguage = "zh-CN",
                        Language = "en-US",
                        Kumimoji = true,
                        Dirty = false,
                        SpellingError = false
                    },


                },
                new D.EndParagraphRunProperties(
                    new D.SolidFill()
                    {
                        RgbColorModelHex = new D.RgbColorModelHex()
                        {
                            Val = HexBinaryValue.FromString(textStyle.Color.Replace("#", ""))
                        }
                    }
                )
                {
                    FontSize = (int)(textStyle.FontSize * 100), //20*100 字号20
                    Underline = textStyle.IsUnderline ? D.TextUnderlineValues.Single : D.TextUnderlineValues.None,
                    Italic = textStyle.IsItalic,
                    Bold = textStyle.IsBold,
                    AlternativeLanguage = "zh-CN",
                    Language = "en-US",
                    Kumimoji = true,
                    Dirty = false,
                    SpellingError = false
                }
                ));

            return textShape;
        }

        public Picture AddPicture(int sldindex, string filePath, D.Transform2D transform)
        {
            var sldpart = Doc.GetSlidePart(sldindex);
            return AddPicture(sldpart, filePath, transform);
        }

        /// <summary>
        /// 在指定位置添加图片 支持 png jpeg gif.
        /// </summary>
        public Picture AddPicture(SlidePart sldpart, string filePath, D.Transform2D transform)
        {
            if (!File.Exists(filePath))
            {
                return null;
            }
            // var type = Path.GetExtension(filePath).Replace(".", "");
            // if (type == "svg")
            // {
            //     type += "+xml";
            // }

            var imgprt = sldpart.AddImagePart(AnalysisHelper.GetImagePartType(filePath));

            string rlid = sldpart.GetIdOfPart(imgprt);
            //DocumentFormat.OpenXml.Packaging.CustomPropertyPartType
            // imgprt.AddExtendedPart
            // (
            //     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            //     $"image/{type}",
            //     $".{type}"
            // );
            imgprt.FeedData(new FileStream(filePath, FileMode.Open));

            
            var tree = sldpart.Slide.CommonSlideData.ShapeTree;
            uint maxid = AnalysisHelper.GetMaxId(tree);
            Picture pic = new Picture();
            pic.NonVisualPictureProperties = new NonVisualPictureProperties
            (
                new P.NonVisualDrawingProperties()
                {
                    Id = maxid + 1,
                    Name = $"PIC{maxid + 1}"
                },
                new P.NonVisualPictureDrawingProperties()
                {
                    PictureLocks = new D.PictureLocks()
                    {
                        NoChangeAspect = true
                    }
                }
            );
            pic.BlipFill = new BlipFill
            (
                new D.Blip
                (
                    // new D.BlipExtensionList
                    // (
                    //     new D.BlipExtension()
                    //     {
                    //         Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" //dpi
                    //     }
                    // )
                )
                {
                    Embed = rlid,
                },
                new D.Stretch(new D.FillRectangle())
            );
            pic.ShapeProperties = new P.ShapeProperties
            (
                new D.PresetGeometry(new D.AdjustValueList())
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            )
            {
                Transform2D = transform
            };
            tree.AppendChild(pic);
            return pic;
        }


        private void InitProperty()
        {
            if (Doc != null)
            {
                var cxysize = Doc.PresentationPart.Presentation.SlideSize;
                _width = Int32Value.ToInt32(cxysize.Cx);
                _height = Int32Value.ToInt32(cxysize.Cy);
            }
        }
        public void Dispose()
        {
            Doc?.Dispose();
        }
    }
}
