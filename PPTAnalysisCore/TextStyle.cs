using System;
namespace PPTAnalysisCore
{
    public class PPTTextStyle
    {
        public double FontSize { get; set; } = 20;
        public string Color { get; set; } = "#000000";
        public bool IsUnderline { get; set; }
        public bool IsItalic { get; set; }
        public bool IsBold { get; set;}
        public float LineSpacing { get; set; }
        public float CharacterSpacing { get; set; }


    }
}
