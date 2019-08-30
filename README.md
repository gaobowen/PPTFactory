# PPTFactory
## 简介
     创建或修改PPT文件
## 例子
1. 创建PPT页并添加文字
```js
static void TestAddNewSlide()
{
    AnalysisCore analysisCore = new AnalysisCore();
    var sldpart = analysisCore.AddNewSlide();
    var transform2D = new D.Transform2D()
    {
        Offset = new Drawing.Offset() 
            { X = (Int6analysisCore.Width * 0.8), Y = (Int6analysisCore.Height * 0.8) },
        Extents = new Drawing.Extents() 
            { Cx = (Int6analysisCore.Width * 0.2), Cy = (Int6analysisCore.Height * 0.1) },
        Rotation = 45 * 60000,
    };
    PPTTextStyle textStyle = new PPTTextStyle();
    analysisCore.AddText(sldpart, "第二个场景页", textStyleransform2D);
    analysisCore.Doc.SaveAs(Directory.GetCurrentDirectory() + "addNewSlide.pptx");
}
```  
  
2. 添加图片
```js
        static void TestAddPicture()
        {
            var analysisCore = new AnalysisCore();
            analysisCore.Doc.PresentationPart.Presentation.AppendChild
            (
                new PresentationExtensionList
                (
                    new PresentationExtension
                    (
                        new DocumentFormat.OpenXml.Office2013.PowerPoint.SlideGuideList()
                    )
                    { Uri = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}" }
                )
            );

            var transform2D = new D.Transform2D()
            {
                Offset = new Drawing.Offset() { X = (Int64)(analysisCore.Width * 0.5), Y = (Int64)(analysisCore.Height * 0.5) },
                Extents = new Drawing.Extents() { Cx = (Int64)(analysisCore.Width * 0.2), Cy = (Int64)(analysisCore.Width * 0.2) },
                Rotation = 45 * 60000,
            };
            analysisCore.AddPicture(0, Directory.GetCurrentDirectory() + "/test.png", transform2D);
            var path = Directory.GetCurrentDirectory() + "/addNewImage.pptx";
            var ret = analysisCore.Doc.SaveAs(path);
            ret.Close();
            ret.Dispose();

            analysisCore.Dispose();

        }
```