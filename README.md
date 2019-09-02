# PPTFactory
## 简介
     创建或修改PPT文件
## 例子
1. 添加文字
```js
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
```
效果如下：
![](PPTFactory/Image/addnewtext.png)
  

2. 创建PPT页并添加文字
```js
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
```  
效果如下：
![](PPTFactory/Image/addnewslide.png)  

2. 添加图片
```js
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
            analysisCore.AddPicture(0, AppDomain.CurrentDomain.BaseDirectory + "/Image/test.png", transform2D);

            var path = AppDomain.CurrentDomain.BaseDirectory + "/addNewPicture.pptx";
            var ret = analysisCore.Doc.SaveAs(path);

            ret.Close();
            ret.Dispose();

            analysisCore.Dispose();
        }
```
效果如下：
![](PPTFactory/Image/addnewpicture.png)