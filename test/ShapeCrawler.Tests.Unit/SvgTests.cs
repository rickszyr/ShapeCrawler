using System.Drawing;
using System.Drawing.Drawing2D;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using NUnit.Framework;
using ShapeCrawler;
using ShapeCrawler.Tests.Unit.Helpers;
using SkiaSharp;
using Svg;
using Svg.Pathing;

namespace ShapeCrawler.Tests.Unit;

[TestFixture]
public class SvgTests : SCTest
{
    [Test]
    public void TestGetSvg()
    {
        var pptx = GetInputStream("autoshape-case019_custom-shapes.pptx");
        var pres = SCPresentation.Open(pptx);
        var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(shape => shape.Name == "Background");
        var result = autoShape.GetSvg();
        File.WriteAllText("result.svg", result);
    }


    [Test]
    public void RectangularGradientTests()
    {
        int width = 300;
        int height = 200;

        using (SKBitmap bitmap = new SKBitmap(width, height))
        {
            using (SKCanvas canvas = new SKCanvas(bitmap))
            {
                // Create a custom path (rectangle in this example)
                using (SKPath path = new SKPath())
                {
                    path.AddRect(new SKRect(0, 0, width, height));

                    // Create a rectangular gradient shader
                    using (SKShader shader = SKShader.CreateLinearGradient(
                        new SKPoint(width / 2, 0),
                        new SKPoint(width / 2, height / 2),
                        new SKColor[] { SKColors.Red, SKColors.Blue },
                        null,
                        SKShaderTileMode.Clamp))
                    {
                        // Use the shader to fill the rectangle
                        using (SKPaint paint = new SKPaint { Shader = shader })
                        {
                            canvas.DrawPath(path, paint);
                            path.Dispose();
                        }
                    }
                }

                // Save the bitmap to a file
                using (SKImage img = SKImage.FromBitmap(bitmap))
                using (SKData data = img.Encode(SKEncodedImageFormat.Png, 100))
                using (var stream = System.IO.File.OpenWrite("gradient.png"))
                {
                    data.SaveTo(stream);
                }
            }
        }
    }

    [Test]
    public void SvgNetTest()
    {
        var document = new SvgDocument();
        var width = 271f;
        var height = 359f;
        document.Width = width;
        document.Height = height;
        document.ViewBox = new SvgViewBox(0, 0, width, height);
        var path = new SvgPath();
        var rect = new SvgRectangle();
        rect.X = -width;
        rect.Y = 0;
        rect.Width = width * 2;
        rect.Height = height * 2;

        path.PathData = new SvgPathSegmentList();
        path.PathData.Add(new SvgMoveToSegment(false, new PointF(-width, 0)));
        path.PathData.Add(new SvgLineSegment(false, new PointF(width, height*2)));
        path.PathData.Add(new SvgLineSegment(false, new PointF(width, 0)));
        path.PathData.Add(new SvgLineSegment(false, new PointF(-width, height*2)));
        path.PathData.Add(new SvgClosePathSegment(false));

        var path2 = new SvgPath();
        path2.PathData = new SvgPathSegmentList();
        path2.PathData.Add(new SvgMoveToSegment(false, new PointF(-width, 0)));
        path2.PathData.Add(new SvgLineSegment(false, new PointF(width, 0)));
        path2.PathData.Add(new SvgLineSegment(false, new PointF(-width, height*2)));
        path2.PathData.Add(new SvgLineSegment(false, new PointF(width, height*2)));
        path2.PathData.Add(new SvgClosePathSegment(false));

        var gradient1 = new SvgLinearGradientServer();
        gradient1.ID = "horizontal";
        gradient1.SpreadMethod = SvgGradientSpreadMethod.Reflect;
        gradient1.X1 = 0.5f;

        var stop1 = new SvgGradientStop();
        stop1.StopColor = new SvgColourServer(Color.FromArgb(0, 0x20, 0x60));
        stop1.Offset = 0;
        var stop3 = new SvgGradientStop();
        stop3.StopColor = new SvgColourServer(Color.FromArgb(0x70, 0x30, 0xA0));
        stop3.Offset = 1f;
        stop3.StopOpacity = 0.45f;

        gradient1.Children.Add(stop1);
        // gradient1.Children.Add(stop2);
        gradient1.Children.Add(stop3);

        document.Children.Add(gradient1);

        var gradient2 = (SvgLinearGradientServer)gradient1.Clone();
        gradient2.ID = "vertical";
        gradient2.X1 = 0;
        gradient2.X2 = 0;
        gradient2.Y1 = 0.5f;
        gradient2.Y2 = 1f;

        document.Children.Add(gradient2);

        path.Fill = gradient1;
        path2.Fill = gradient2;

        var clipPathData = new SvgPath();
        clipPathData.ID = "customShape";
        clipPathData.PathData = new SvgPathSegmentList();
        clipPathData.PathData.Add(new SvgMoveToSegment(false, new PointF(6.95647f, 155.06f)));
        clipPathData.PathData.Add(new SvgCubicCurveSegment(false, new PointF(13.9129f, 228.867f), new PointF(119.772f, 359f), new PointF(190.849f, 351.555f)));
        clipPathData.PathData.Add(new SvgCubicCurveSegment(false, new PointF(261.624f, 343.785f), new PointF(271f, 201.027f), new PointF(264.044f, 127.22f)));
        clipPathData.PathData.Add(new SvgCubicCurveSegment(false, new PointF(257.087f, 53.413f), new PointF(193.874f, 0f), new PointF(123.099f, 7.44545f)));
        clipPathData.PathData.Add(new SvgCubicCurveSegment(false, new PointF(52.0223f, 15.2146f), new PointF(0f, 81.2525f), new PointF(6.95647f, 155.06f)));
        clipPathData.PathData.Add(new SvgClosePathSegment(false));

        var definitionList = new SvgDefinitionList();
        definitionList.Children.Add(clipPathData);

        var clipPath = new SvgClipPath();
        clipPath.ID = "clip-path";
        var use = new SvgUse();
        use.CustomAttributes.Add("href", "#customShape");
        clipPath.Children.Add(use);

        var group = new SvgGroup();
        rect.Fill = gradient1;

        var line = new SvgLine();
        line.StartX = 0;
        line.StartY = height;
        line.EndX = width;
        line.EndY = 0;
        line.StrokeWidth = 2;
        line.Stroke = gradient2;

        group.Children.Add(rect);
        group.Children.Add(path);
        group.Children.Add(path2);
        group.ClipPath = new Uri("url(#clip-path)", UriKind.Relative);
        group.ClipRule = SvgClipRule.NonZero;
        group.ShapeRendering = SvgShapeRendering.GeometricPrecision;

        document.Children.AddAndForceUniqueID(definitionList);
        document.Children.AddAndForceUniqueID(clipPath);
        document.Children.AddAndForceUniqueID(group);

        var xml = document.GetXML();
        File.WriteAllText("SvgNetResult.svg", xml);
    }
}