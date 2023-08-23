using System;
using System.Drawing;
using System.Xml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Svg;
using Svg.Pathing;

namespace ShapeCrawler;
internal class SVGBuilder
{
    private readonly SvgDocument document;
    private SCSlideMaster? slideMaster;
    private bool shouldRedrawOutlinePath = false;
    private P.ShapeProperties? shapeProperties;
    private int id;

    public SVGBuilder()
    {
        this.document = new SvgDocument();
    }

    internal string BuildFromCustomGeometry(SCShape scShape, int width, int height, SCSlideMaster master)
    {
        this.slideMaster = master;
        this.shapeProperties = scShape.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !;
        this.id = scShape.Id;

        var result = string.Empty;
        var customGeometry = this.shapeProperties?.GetFirstChild<D.CustomGeometry>() !;

        var shapePath = customGeometry.GetFirstChild<D.PathList>()?.GetFirstChild<D.Path>();
        var pixelWidth = UnitConverter.PointToPixel(UnitConverter.EmuToPoint((int)shapePath!.Width!));
        var horizontalRatio = width / pixelWidth;
        var pixelHeight = UnitConverter.PointToPixel(UnitConverter.EmuToPoint((int)shapePath!.Height!));
        var verticalRatio = height / pixelHeight;

        this.document.Width = width;
        this.document.Height = height;

        var svgPath = shapePath.ToSvgPath(horizontalRatio, verticalRatio);

        this.ProcessFill(this.shapeProperties!, svgPath);
        this.ProcessOutline(this.shapeProperties!, svgPath);

        var xml = new XmlDocument();
        xml.LoadXml(this.document.GetXML());
        var svgNode = xml.LastChild;

        return svgNode!.OuterXml;
    }

    /// <summary>
    /// ToDo Implement for other shapes.
    /// </summary>
    /// <param name="shape"></param>
    /// <param name="width"></param>
    /// <param name="height"></param>
    /// <param name="master"></param>
    /// <returns></returns>
    internal string BuildFromAutoshape(SCShape shape, int width, int height, SCSlideMaster master){
        return string.Empty;
    }

    private void ProcessOutline(P.ShapeProperties shapeProperties, SvgPath svgPath)
    {
        var outline = shapeProperties.GetFirstChild<D.Outline>();

        if (outline is not null)
        {
            var solidFill = outline!.GetFirstChild<D.SolidFill>();
            var gradientFill = outline.GetFirstChild<D.GradientFill>();
            var noFill = outline.GetFirstChild<D.NoFill>();

            if (noFill is not null)
            {
                // no outline to process.
                return;
            }

            var strokeWidth = UnitConverter.PointToPixel(UnitConverter.EmuToPoint(outline?.Width!));
            if (strokeWidth > 1)
            {
                this.document.Width += strokeWidth / 2;
                this.document.Height += strokeWidth / 2;
            }

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
    /// It's possible that the fill comes from a group. In that case it will depend on the group fill.
    /// If it's a solid color, it's straight forward. Not yet implemented if it's a gradient.
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
        var groupFill = shapeProperties.GetFirstChild<D.GroupFill>();


        if ((solidFill is not null))
        {
            this.ProcessSolidFill(shapePath, solidFill);
        }
        else if (gradientFill is not null)
        {
            this.ProcessGradientFill(shapePath, gradientFill);
        }
        else if (groupFill is not null)
        {
            this.ProcessGroupFill(shapePath, groupFill);
        }
    }

    private void ProcessSolidFill(SvgPath shapePath, D.SolidFill? solidFill)
    {
        shapePath.Fill = new SvgColourServer(HexParser.GetSolidColorFromElement(solidFill!, this.slideMaster!));
        this.document.Children.Add(shapePath);
    }

    private void ProcessGradientFill(SvgPath shapePath, D.GradientFill gradientFill)
    {
        var linearGradientFill = gradientFill.GetFirstChild<D.LinearGradientFill>();
        var pathGradientFill = gradientFill.GetFirstChild<D.PathGradientFill>();

        if (linearGradientFill is not null)
        {
            var gradient = linearGradientFill.ToSvgLinearGradient(this.slideMaster!);
            gradient.ID = $"lg_{this.id}";
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
                this.shouldRedrawOutlinePath = true;
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
        horizontalGradient.ID = $"horizontal_{this.id}";
        horizontalGradient.SpreadMethod = SvgGradientSpreadMethod.Reflect;
        horizontalGradient.X1 = 0.5f;

        var stops = gradientFill.GetGradientStops(this.slideMaster!);
        stops.ForEach(s => horizontalGradient.Children.Add(s));


        var verticalGradient = (SvgLinearGradientServer)horizontalGradient.Clone();
        verticalGradient.ID = $"vertical_{this.id}";
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
        shapePath.ID = $"custom-shape_{this.id}";
        definitions.Children.AddAndForceUniqueID(shapePath);

        var clipPath = new SvgClipPath();
        clipPath.ID = $"shape-clip_{this.id}";
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

    private void ProcessGroupFill(SvgPath shapePath, D.GroupFill? groupFill)
    {
        var parentGroups = this.shapeProperties?.Ancestors<P.GroupShape>() !;
        foreach (var ancestor in parentGroups)
        {
            var visualProperties = ancestor.GetFirstChild<P.GroupShapeProperties>();
            if (visualProperties != null)
            {
                var solidFill = visualProperties.GetFirstChild<D.SolidFill>();
                if (solidFill is not null)
                {
                    this.ProcessSolidFill(shapePath, solidFill);
                }

                break;
            }
        }
    }
}