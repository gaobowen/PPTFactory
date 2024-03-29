﻿using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using System.IO;
using System.Linq;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

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

        private AnalysisCore()
        {
            if (_prototype16x9 == null)
            {
                _prototype16x9 = PresentationDocument.Open(AppDomain.CurrentDomain.BaseDirectory + "/blankone.pptx", true);
            }

            Doc = (PresentationDocument)_prototype16x9.Clone();
            InitProperty();
        }

        private AnalysisCore(PresentationDocument presentationDocument)
        {
            Doc = presentationDocument;
            InitProperty();
        }

        public AnalysisCore(string filePath)
        {
            Doc = PresentationDocument.Open(filePath, true);
            InitProperty();
        }

        public static AnalysisCore New(string path)
        {
            File.Copy(AppDomain.CurrentDomain.BaseDirectory + "/blankone.pptx", path, true);
            return new AnalysisCore(path);
        }

        public static AnalysisCore Open(string path)
        {
            return new AnalysisCore(path);
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
            slide.CommonSlideData.ShapeTree.AppendChild(
                new GroupShapeProperties()
                {
                    TransformGroup = new D.TransformGroup
                    (
                        new D.Offset() { X = 0L, Y = 0L },
                        new D.Extents() { Cx = 0L, Cy = 0L },
                        new D.ChildOffset() { X = 0L, Y = 0L },
                        new D.ChildExtents() { Cx = 0L, Cy = 0L }
                    )
                });
            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();
            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };
            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)4033567156U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
            commonSlideDataExtension1.Append(creationId1);
            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);
            slide.CommonSlideData.Append(commonSlideDataExtensionList1);
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

            var imgprt = sldpart.AddImagePart(AnalysisHelper.GetImagePartType(filePath));
            string rlid = sldpart.GetIdOfPart(imgprt);
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
                },
                //必须指定元素的类型，否则会报错。
                new ApplicationNonVisualDrawingProperties
                (
                    new PlaceholderShape()
                    {
                        Type = PlaceholderValues.Picture
                    }
                )
            );

            P.BlipFill blipFill = new P.BlipFill();
            D.Blip blip = new D.Blip() { Embed = rlid };
            D.BlipExtensionList blipExtensionList = new D.BlipExtensionList();
            D.BlipExtension blipExtension = new D.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi =
                new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14",
                "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            blip.Append(blipExtensionList);

            D.Stretch stretch = new D.Stretch();
            D.FillRectangle fillRectangle = new D.FillRectangle();
            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);
            pic.Append(blipFill);
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

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                imgprt.FeedData(fileStream);
            }

            return pic;
        }


        public void AddVideo(SlidePart slidepart, string videoFilePath, string videoCoverPath, D.Transform2D transform)
        {
            Slide slide = slidepart.Slide;
            ShapeTree shapeTree1 = slidepart.Slide.CommonSlideData.ShapeTree;
            var ptrlid = Doc.PresentationPart.GetIdOfPart(slidepart);
            var picID = AnalysisHelper.GetMaxId(shapeTree1);
            string imgEmbedId = string.Format("imgId{0}{1}{2}", ptrlid, picID, 1);
            string videoEmbedId = string.Format("vidId{0}{1}{2}", ptrlid, picID, 2);
            string mediaEmbedId = string.Format("medId{0}{1}{2}", ptrlid, picID, 3);

            Picture picture1 = new Picture();
            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 =
                new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = videoEmbedId + "" };
            D.HyperlinkOnClick hyperlinkOnClick1 = new D.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };
            nonVisualDrawingProperties2.Append(hyperlinkOnClick1);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            D.PictureLocks pictureLocks1 = new D.PictureLocks() { NoChangeAspect = true };
            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            D.VideoFromFile videoFromFile1 = new D.VideoFromFile() { Link = videoEmbedId };

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = 
                new ApplicationNonVisualDrawingPropertiesExtensionList();
            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = 
                new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

            P14.Media media1 = new P14.Media() { Embed = mediaEmbedId };
            media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
            applicationNonVisualDrawingPropertiesExtension1.Append(media1);
            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(videoFromFile1);
            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

            BlipFill blipFill1 = new BlipFill();
            D.Blip blip1 = new D.Blip() { Embed = imgEmbedId };

            D.Stretch stretch1 = new D.Stretch();
            D.FillRectangle fillRectangle1 = new D.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            D.PresetGeometry presetGeometry1 = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
            D.AdjustValueList adjustValueList1 = new D.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            shapeTree1.Append(picture1);

            if (!(slide.Timing?.ChildElements?.Count > 0))
            {
                AnalysisHelper.InitTiming(slide);
            }

            ImagePart imagePart = slidepart.AddImagePart(AnalysisHelper.GetImagePartType(videoCoverPath), imgEmbedId);
            using (var data = File.OpenRead(videoCoverPath))
            {
                imagePart.FeedData(data);
            };
            
            Doc.PartExtensionProvider.AddPartExtension("video/mp4", ".mp4");
            MediaDataPart mediaDataPart1 = Doc.CreateMediaDataPart("video/mp4", ".mp4");
         
            using (System.IO.Stream mediaDataPart1Stream = File.OpenRead(videoFilePath))
            {
                mediaDataPart1.FeedData(mediaDataPart1Stream);
            }
            slidepart.AddVideoReferenceRelationship(mediaDataPart1, videoEmbedId);
            slidepart.AddMediaReferenceRelationship(mediaDataPart1, mediaEmbedId);
            slide.Save();
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
