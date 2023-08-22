using System;
using System.Drawing;
using System.Xml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Svg;
using Svg.Pathing;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;
internal class SVGBuilder
{
    private readonly SvgDocument document;
    private SCSlideMaster? slideMaster;
    private bool shouldRedrawOutlinePath;
    public SVGBuilder()
    {
        this.document = new SvgDocument();
        this.shouldRedrawOutlinePath = false;
    }

    internal string BuildFromCustomGeometry(P.ShapeProperties shapeProperties, int width, int height, SCSlideMaster master)
    {
        this.slideMaster = master;

        var result = string.Empty;
        var customGeometry = shapeProperties?.GetFirstChild<D.CustomGeometry>()!;

        var shapePath = customGeometry.GetFirstChild<D.PathList>()?.GetFirstChild<D.Path>();
        var pixelWidth = UnitConverter.PointToPixel(UnitConverter.EmuToPoint((int)shapePath!.Width!));
        var horizontalRatio = width / pixelWidth;
        var pixelHeight = UnitConverter.PointToPixel(UnitConverter.EmuToPoint((int)shapePath!.Height!));
        var verticalRatio = height / pixelHeight;

        this.document.Width = width;
        this.document.Height = height;

        var svgPath = shapePath.ToSvgPath(horizontalRatio, verticalRatio);

        this.ProcessFill(shapeProperties!, svgPath);
        this.ProcessOutline(shapeProperties!, svgPath);


        var xml = new XmlDocument();
        xml.LoadXml(this.document.GetXML());
        var svgNode = xml.LastChild;

        return svgNode!.OuterXml;
    }

    private void ProcessOutline(ShapeProperties shapeProperties, SvgPath svgPath)
    {
        var outline = shapeProperties.GetFirstChild<D.Outline>();
        var strokeWidth = UnitConverter.PointToPixel(UnitConverter.EmuToPoint(outline?.Width!));
        if(strokeWidth > 1){
            this.document.Width += strokeWidth;
            this.document.Height += strokeWidth;
        }


        if (outline is not null)
        {
            var solidFill = outline.GetFirstChild<D.SolidFill>();
            var gradientFill = outline.GetFirstChild<D.GradientFill>();
            var noFill = outline.GetFirstChild<D.NoFill>();

            if (solidFill is not null)
            {
                var color = HexParser.GetSolidColorFromElement(solidFill, this.slideMaster!);
                if (this.shouldRedrawOutlinePath)
                {
                    svgPath = (SvgPath)svgPath.Clone();
                    this.document.Children.Add(svgPath);
                }

                svgPath.Stroke = new SvgColourServer(color);
                svgPath.StrokeOpacity = color.A / 255f;
                svgPath.StrokeWidth = strokeWidth;
            }
        }
    }

    /// <summary>
    /// Processes the fill, adding it to the Svg Document.
    /// If it's a simple solid color fill, it is added to the path.
    /// If it's a gradient fill, it depends on whether it's a simple linear gradient, or a path gradient.
    /// <para>Linear gradient - it is added to the document and then assigned to the path.</para>
    /// <para>Path gradient - it is processed separately for each type (rect, circle &lt; not currently supported, shape &lt; not currently supported.)</para>
    /// </summary>
    /// <param name="shapeProperties"></param>
    /// <param name="shapePath"></param>
    private void ProcessFill(P.ShapeProperties shapeProperties, SvgPath shapePath)
    {
        var solidFill = shapeProperties.GetFirstChild<D.SolidFill>();
        var gradientFill = shapeProperties.GetFirstChild<D.GradientFill>();
        var noFill = shapeProperties.GetFirstChild<D.NoFill>();
        var test = new ColorConverter();

        if (solidFill is not null)
        {
            shapePath.Fill = new SvgColourServer(HexParser.GetSolidColorFromElement(solidFill!, this.slideMaster!));
        }
        else if (gradientFill is not null)
        {
            var linearGradientFill = gradientFill.GetFirstChild<D.LinearGradientFill>();
            var pathGradientFill = gradientFill.GetFirstChild<D.PathGradientFill>();

            if (linearGradientFill is not null)
            {
                var gradient = linearGradientFill.ToSvgLinearGradient(this.slideMaster!);
                shapePath.Fill = gradient;

                this.document.Children.Add(gradient);
                this.document.Children.Add(shapePath);
            }
            else if (pathGradientFill is not null)
            {
                var pathType = pathGradientFill.Path!.Value;
                if (pathType == D.PathShadeValues.Rectangle)
                {
                    this.ProcessRectangleFill(shapePath, gradientFill);
                }
            }
        }
    }

    private void ProcessRectangleFill(SvgPath shapePath, D.GradientFill gradientFill)
    {
        var tileRect = gradientFill.GetFirstChild<D.TileRectangle>();
        var tileLeft = (tileRect?.Left.ToPercentValue() ?? 0) * this.document.Width;
        var tileRight = (tileRect?.Right.ToPercentValue() ?? 0) * this.document.Width;
        var tileTop = (tileRect?.Top.ToPercentValue() ?? 0) * this.document.Height;
        var tileBottom = (tileRect?.Bottom.ToPercentValue() ?? 0) * this.document.Height;

        var x = tileLeft;
        var y = tileTop;
        var gradientRectWidth = this.document.Width - tileRight - tileLeft;
        var gradientRectHeight = this.document.Height - tileTop - tileBottom;

        var horizontalGradient = new SvgLinearGradientServer();
        horizontalGradient.ID = "horizontal";
        horizontalGradient.SpreadMethod = SvgGradientSpreadMethod.Reflect;
        horizontalGradient.X1 = 0.5f;

        var stops = gradientFill.GetGradientStops(this.slideMaster!);
        stops.ForEach(s => horizontalGradient.Children.Add(s));


        var verticalGradient = (SvgLinearGradientServer)horizontalGradient.Clone();
        verticalGradient.ID = "vertical";
        verticalGradient.Y1 = 0.5f;
        verticalGradient.Y2 = 1f;
        verticalGradient.X1 = 0;
        verticalGradient.X2 = 0;

        var backgroundRect = new SvgRectangle();
        backgroundRect.X = x;
        backgroundRect.Y = y;
        backgroundRect.Width = gradientRectWidth;
        backgroundRect.Height = gradientRectHeight;
        backgroundRect.Fill = horizontalGradient;

        var verticalPath = new SvgPath();
        verticalPath.PathData = new SvgPathSegmentList();
        verticalPath.MoveTo(new PointF(x, y));
        verticalPath.AddLineTo(new PointF(x + backgroundRect.Width, y + backgroundRect.Height));
        verticalPath.AddLineTo(new PointF(x, y + backgroundRect.Height));
        verticalPath.AddLineTo(new PointF(x + backgroundRect.Width, y));
        verticalPath.ClosePath();
        verticalPath.Fill = verticalGradient;

        var group = new SvgGroup();
        group.Children.AddAndForceUniqueID(backgroundRect);
        group.Children.AddAndForceUniqueID(verticalPath);

        var definitions = new SvgDefinitionList();
        shapePath.ID = "custom-shape";
        definitions.Children.AddAndForceUniqueID(shapePath);

        var clipPath = new SvgClipPath();
        clipPath.ID = "shape-clip";
        var use = new SvgUse();
        use.CustomAttributes.Add("href", $"#{shapePath.ID}");
        clipPath.Children.Add(use);

        group.ClipPath = new Uri($"url(#{clipPath.ID})", UriKind.Relative);

        this.document.Children.AddAndForceUniqueID(verticalGradient);
        this.document.Children.AddAndForceUniqueID(horizontalGradient);
        this.document.Children.AddAndForceUniqueID(definitions);
        this.document.Children.AddAndForceUniqueID(clipPath);
        this.document.Children.AddAndForceUniqueID(group);
    }
}