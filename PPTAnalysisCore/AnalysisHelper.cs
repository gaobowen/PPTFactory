using System;
using DocumentFormat.OpenXml.Packaging;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace PPTAnalysisCore
{
    public static class AnalysisHelper
    {
        public static ImagePartType GetImagePartType(string filepath)
        {
            var extstr = System.IO.Path.GetExtension(filepath);
            switch (extstr.ToLower())
            {
                case ".png":
                    return ImagePartType.Png;
                case ".jpeg":
                case ".jpg":
                    return ImagePartType.Jpeg;
                case ".gif":
                    return ImagePartType.Gif;
                default:
                    throw new Exception($"not support {extstr}");
            }
        }


        public static UInt32Value GetMaxId(ShapeTree tree)
        {
            var maxid = (uint)tree.ChildElements.Count;
            if (tree.ChildElements.Count > 0)
            {
                foreach (var child in tree.ChildElements)
                {
                    if (child is Shape shape)
                    {
                        if (shape?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id != null)
                        {
                            var id = UInt32Value.ToUInt32(shape?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id);
                            if (id > maxid)
                            {
                                maxid = id;
                            }
                        }
                    }
                }
            }
            return maxid;
        }

        public static void InitTiming(Slide slide)
        {
            Timing timing1 = new Timing();

            TimeNodeList timeNodeList1 = new TimeNodeList();

            ParallelTimeNode parallelTimeNode1 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode1 = new CommonTimeNode() { Id = (UInt32Value)1U, Duration = "indefinite", Restart = TimeNodeRestartValues.Never, NodeType = TimeNodeValues.TmingRoot };

            ChildTimeNodeList childTimeNodeList1 = new ChildTimeNodeList();

            SequenceTimeNode sequenceTimeNode1 = new SequenceTimeNode() { Concurrent = true, NextAction = NextActionValues.Seek };

            CommonTimeNode commonTimeNode2 = new CommonTimeNode() { Id = (UInt32Value)2U, Restart = TimeNodeRestartValues.WhenNotActive, Fill = TimeNodeFillValues.Hold, EventFilter = "cancelBubble", NodeType = TimeNodeValues.InteractiveSequence };

            StartConditionList startConditionList1 = new StartConditionList();

            Condition condition1 = new Condition() { Event = TriggerEventValues.OnClick, Delay = "0" };

            TargetElement targetElement1 = new TargetElement();
            ShapeTarget shapeTarget1 = new ShapeTarget() { ShapeId = "3" };

            targetElement1.Append(shapeTarget1);

            condition1.Append(targetElement1);

            startConditionList1.Append(condition1);

            EndSync endSync1 = new EndSync() { Event = TriggerEventValues.End, Delay = "0" };
            RuntimeNodeTrigger runtimeNodeTrigger1 = new RuntimeNodeTrigger() { Val = TriggerRuntimeNodeValues.All };

            endSync1.Append(runtimeNodeTrigger1);

            ChildTimeNodeList childTimeNodeList2 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode2 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode3 = new CommonTimeNode() { Id = (UInt32Value)3U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList2 = new StartConditionList();
            Condition condition2 = new Condition() { Delay = "0" };

            startConditionList2.Append(condition2);

            ChildTimeNodeList childTimeNodeList3 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode3 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode4 = new CommonTimeNode() { Id = (UInt32Value)4U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList3 = new StartConditionList();
            Condition condition3 = new Condition() { Delay = "0" };

            startConditionList3.Append(condition3);

            ChildTimeNodeList childTimeNodeList4 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode4 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode5 = new CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 2, PresetClass = TimeNodePresetClassValues.MediaCall, PresetSubtype = 0, Fill = TimeNodeFillValues.Hold, NodeType = TimeNodeValues.ClickEffect };

            StartConditionList startConditionList4 = new StartConditionList();
            Condition condition4 = new Condition() { Delay = "0" };

            startConditionList4.Append(condition4);

            ChildTimeNodeList childTimeNodeList5 = new ChildTimeNodeList();

            Command command1 = new Command() { Type = CommandValues.Call, CommandName = "togglePause" };

            CommonBehavior commonBehavior1 = new CommonBehavior();
            CommonTimeNode commonTimeNode6 = new CommonTimeNode() { Id = (UInt32Value)6U, Duration = "1", Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement2 = new TargetElement();
            ShapeTarget shapeTarget2 = new ShapeTarget() { ShapeId = "3" };

            targetElement2.Append(shapeTarget2);

            commonBehavior1.Append(commonTimeNode6);
            commonBehavior1.Append(targetElement2);

            command1.Append(commonBehavior1);

            childTimeNodeList5.Append(command1);

            commonTimeNode5.Append(startConditionList4);
            commonTimeNode5.Append(childTimeNodeList5);

            parallelTimeNode4.Append(commonTimeNode5);

            childTimeNodeList4.Append(parallelTimeNode4);

            commonTimeNode4.Append(startConditionList3);
            commonTimeNode4.Append(childTimeNodeList4);

            parallelTimeNode3.Append(commonTimeNode4);

            childTimeNodeList3.Append(parallelTimeNode3);

            commonTimeNode3.Append(startConditionList2);
            commonTimeNode3.Append(childTimeNodeList3);

            parallelTimeNode2.Append(commonTimeNode3);

            childTimeNodeList2.Append(parallelTimeNode2);

            commonTimeNode2.Append(startConditionList1);
            commonTimeNode2.Append(endSync1);
            commonTimeNode2.Append(childTimeNodeList2);

            NextConditionList nextConditionList1 = new NextConditionList();

            Condition condition5 = new Condition() { Event = TriggerEventValues.OnClick, Delay = "0" };

            TargetElement targetElement3 = new TargetElement();
            ShapeTarget shapeTarget3 = new ShapeTarget() { ShapeId = "3" };

            targetElement3.Append(shapeTarget3);

            condition5.Append(targetElement3);

            nextConditionList1.Append(condition5);

            sequenceTimeNode1.Append(commonTimeNode2);
            sequenceTimeNode1.Append(nextConditionList1);

            Video video1 = new Video();

            CommonMediaNode commonMediaNode1 = new CommonMediaNode() { Volume = 80000 };

            CommonTimeNode commonTimeNode7 = new CommonTimeNode() { Id = (UInt32Value)7U, Fill = TimeNodeFillValues.Hold, Display = false };

            StartConditionList startConditionList5 = new StartConditionList();
            Condition condition6 = new Condition() { Delay = "indefinite" };

            startConditionList5.Append(condition6);

            commonTimeNode7.Append(startConditionList5);

            TargetElement targetElement4 = new TargetElement();
            ShapeTarget shapeTarget4 = new ShapeTarget() { ShapeId = "3" };

            targetElement4.Append(shapeTarget4);

            commonMediaNode1.Append(commonTimeNode7);
            commonMediaNode1.Append(targetElement4);

            video1.Append(commonMediaNode1);

            childTimeNodeList1.Append(sequenceTimeNode1);
            childTimeNodeList1.Append(video1);

            commonTimeNode1.Append(childTimeNodeList1);

            parallelTimeNode1.Append(commonTimeNode1);

            timeNodeList1.Append(parallelTimeNode1);

            timing1.Append(timeNodeList1);

            slide.Append(timing1);
        }

        public static PresentationDocument CreateBlankPPT(string filepath)
        {
            PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation, true);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationPart);
            return presentationDoc;
        }

        private static void CreatePresentationParts(PresentationPart presentationPart)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (DocumentFormat.OpenXml.UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000 };
            NotesSize notesSize1 = new NotesSize() { Cy = 6858000, Cx = 12192000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;


            slidePart1 = CreateAddSlidePart(presentationPart, "rId2");
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }

        private static SlidePart CreateAddSlidePart(PresentationPart presentationPart, string addid)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>(addid);
            slidePart1.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new D.TransformGroup())
                            )),
                    new ColorMapOverride(new D.MasterColorMapping()));
            //缺少引用建立slideidlist
            return slidePart1;
        }

        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new D.TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new D.BodyProperties(),
                new D.ListStyle(),
                new D.Paragraph(new D.EndParagraphRunProperties()))))),
            new ColorMapOverride(new D.MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new D.TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new D.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new D.BodyProperties(),
                new D.ListStyle(),
                new D.Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
            { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }))
              { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 })
                { Val = D.SchemeColorValues.PhColor })
                { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 })
                { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 })
                  { Val = "000000" })
                { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 })
                  { Val = D.SchemeColorValues.PhColor })
                { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
              { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }
    }
}
