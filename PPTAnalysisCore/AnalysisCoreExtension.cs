using System;
using DocumentFormat.OpenXml.Packaging;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using System.Drawing;

namespace PPTAnalysisCore
{
    public static class AnalysisCoreExtension
    {
        /// <summary>
        /// 生成transform
        /// </summary>
        /// <param name="core">ppt</param>
        /// <param name="resolution">分辨率</param>
        /// <param name="bounding">位置</param>
        /// <returns>Transform2D</returns>
        public static D.Transform2D CreateTransform2D(this AnalysisCore core, Size resolution, Bounding bounding)
        {
            int bw = core.Width;
            int bh = core.Height;
            var transform2D = new D.Transform2D()
            {
                Offset = new D.Offset()
                {
                    X = (Int64)(bw * bounding.X / resolution.Width),
                    Y = (Int64)(bh * bounding.Y / resolution.Height)
                },
                Extents = new D.Extents()
                {
                    Cx = (Int64)(bw * bounding.Width / resolution.Width),
                    Cy = (Int64)(bh * bounding.Height / resolution.Height)
                },
                Rotation = (int)(bounding.Rotation * 60000),
                HorizontalFlip = bounding.IsFlipped,
                VerticalFlip = bounding.IsFlippedV
            };
            return transform2D;
        }


        public static ShapeTree GetShapeTree(this PresentationDocument doc, int index)
        {
            var ids = doc.PresentationPart.Presentation.SlideIdList;
            var slprt = (SlidePart)doc.PresentationPart.GetPartById((ids.ChildElements[index] as SlideId)?.RelationshipId);
            return slprt?.Slide?.CommonSlideData?.ShapeTree;
        }

        public static SlidePart GetSlidePart(this PresentationDocument doc, int index)
        {
            var ids = doc.PresentationPart.Presentation.SlideIdList;
            var slprt = (SlidePart)doc.PresentationPart.GetPartById((ids.ChildElements[index] as SlideId)?.RelationshipId);
            return slprt;
        }

    }
}
