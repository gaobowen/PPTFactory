using System;
namespace PPTAnalysisCore
{
    public class Bounding
    {
        public double X { get; set; }

        public double Y { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        public double Rotation { get; set; }

        public double Opacity { get; set; } = 1.0;

        public bool IsFlipped { get; set; }

        public bool IsFlippedV { get; set; }

    }
}
