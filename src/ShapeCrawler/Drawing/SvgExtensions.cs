using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ShapeCrawler.Shared;
using Svg;
using Svg.Pathing;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;


namespace ShapeCrawler.Drawing
{
    internal static class SvgExtensions
    {
        internal static SvgPath ToSvgPath(this D.Path path, float horizontalRatio, float verticalRatio)
        {
            var result = new SvgPath();
            result.PathData = new SvgPathSegmentList();

            foreach (var pathElement in path.ChildElements)
            {
                if (pathElement is D.MoveTo moveTo)
                {
                    result.MoveTo(moveTo.Point!.ToPointF(horizontalRatio, verticalRatio));
                }
                else if (pathElement is D.CubicBezierCurveTo cubicBezierCurveTo)
                {
                    result.AddCubicBezierCurveTo(cubicBezierCurveTo.ToSvgCubicBezierCurve(horizontalRatio, verticalRatio));
                }
                else if (pathElement is D.LineTo lineTo)
                {
                    result.AddLineTo(lineTo.Point!.ToPointF(horizontalRatio, horizontalRatio));
                }
                else if (pathElement is D.CloseShapePath)
                {
                    result.ClosePath();
                }
            }

            return result;
        }

        /// <summary>
        /// Transforms an OpenXML point to a PointF while scaling the point to the shape's actual dimensions, 
        /// as sometimes the custom shape description doesn't match the actual shape's size.
        /// </summary>
        /// <param name="point"></param>
        /// <param name="horizontalRatio">The transform value to map the point correctly horizontally.</param>
        /// <param name="verticalRatio">The transform value to map the point correctly vertically.</param>
        /// <returns></returns>
        internal static PointF ToPointF(this D.Point point, float horizontalRatio = 1f, float verticalRatio = 1f)
        {
            var x = UnitConverter.PointToPixel(UnitConverter.EmuToPoint(int.Parse(point?.X?.Value ?? "0"))) * horizontalRatio;
            var y = UnitConverter.PointToPixel(UnitConverter.EmuToPoint(int.Parse(point?.Y?.Value ?? "0"))) * verticalRatio;
            return new PointF(x, y);
        }

        internal static void MoveTo(this SvgPath path, PointF point, bool isRelative = false)
        {
            path.PathData.Add(new SvgMoveToSegment(isRelative, point));
        }

        internal static SvgCubicCurveSegment ToSvgCubicBezierCurve(this D.CubicBezierCurveTo curve, float horizontalRatio, float verticalRatio, bool isRelative = false)
        {
            var point1 = (curve.ChildElements[0] as D.Point) !;
            var point2 = (curve.ChildElements[1] as D.Point) !;
            var point3 = (curve.ChildElements[2] as D.Point) !;

            return new SvgCubicCurveSegment(isRelative, point1.ToPointF(horizontalRatio, verticalRatio), point2.ToPointF(horizontalRatio, verticalRatio), point3.ToPointF(horizontalRatio, verticalRatio));
        }

        internal static void AddCubicBezierCurveTo(this SvgPath path, SvgCubicCurveSegment svgCubicCurveSegment)
        {
            path.PathData.Add(svgCubicCurveSegment);
        }

        internal static void AddLineTo(this SvgPath path, PointF point, bool isRelative = false)
        {
            path.PathData.Add(new SvgLineSegment(isRelative, point));
        }

        internal static void ClosePath(this SvgPath path, bool isRelative = false)
        {
            path.PathData.Add(new SvgClosePathSegment(isRelative));
        }

        internal static SvgLinearGradientServer ToSvgLinearGradient(this D.LinearGradientFill fill, SCSlideMaster slideMaster)
        {
            var result = new SvgLinearGradientServer();
            var gradientFillParent = fill.Parent as D.GradientFill;
            var stops = gradientFillParent!.GetGradientStops(slideMaster);
            var xmlAngle = fill.Angle?.Value;
            var scale = fill.Scaled?.Value ?? false;
            var properties = fill.Ancestors<P.ShapeProperties>().FirstOrDefault() !;            
            var gradientAngle = UnitConverter.AngleValueToRadians(xmlAngle ?? 0);
            var xfrm = properties.Transform2D;
            var shapePixelWidth = UnitConverter.HorizontalEmuToPixel(xfrm!.Extents!.Cx!);
            var shapePixelHeight = UnitConverter.VerticalEmuToPixel(xfrm!.Extents!.Cy!);
            var startPoint = GetStartXYFromAngle(gradientAngle, shapePixelWidth, shapePixelHeight);
            var endPoint = GetEndXYFromAngle(startPoint, gradientAngle, shapePixelWidth, shapePixelHeight);

            stops.ForEach(s => result.Children.Add(s));

            result.X1 = startPoint.X / shapePixelWidth;
            result.Y1 = startPoint.Y / shapePixelHeight;
            result.X2 = endPoint.X / shapePixelWidth;
            result.Y2 = endPoint.Y / shapePixelHeight;

            return result;
        }

        internal static List<SvgGradientStop> GetGradientStops(this D.GradientFill gradientFill, SCSlideMaster slideMaster)
        {
            var result = new List<SvgGradientStop>();
            var stops = gradientFill!.GetFirstChild<D.GradientStopList>()?.Elements<D.GradientStop>() !;
            var orderedStops = stops.OrderBy(s => s.Position);
            List<Color> colors = new List<Color>();
            List<float> positions = new List<float>();

            foreach (var stop in orderedStops)
            {
                var svgStop = new SvgGradientStop();
                var color = HexParser.GetSolidColorFromElement(stop, slideMaster);
                svgStop.StopColor = new SvgColourServer(HexParser.GetSolidColorFromElement(stop, slideMaster));
                svgStop.Offset = stop.Position!.Value / 100000f;
                svgStop.StopOpacity = color.A / 255f;
                result.Add(svgStop);
            }

            return result;
        }

        private static PointF GetStartXYFromAngle(double angle, int shapePixelWidth, int shapePixelHeight)
        {
            var result = angle switch
            {
                var a when a >= 0 && a < Math.PI / 2 => new PointF(0, 0),
                var a when a >= Math.PI / 2 && a < Math.PI => new PointF(shapePixelWidth, 0),
                var a when a >= Math.PI && a < 3 * Math.PI / 2 => new PointF(shapePixelWidth, shapePixelHeight),
                _ => new PointF(0, shapePixelHeight),
            };
            return result;
        }

        private static PointF GetEndXYFromAngle(PointF startXY, double gradientAngle, int width, int height)
        {
            var ninetyDegrees = Math.PI / 2;
            var quadrant = Math.Floor(gradientAngle / ninetyDegrees);

            // either the height or the width will be the Perpendicular for the diagonal triangle
            var perpendicular = (height * Math.Abs(Math.Cos(ninetyDegrees * quadrant))) + (width * Math.Abs(Math.Sin(ninetyDegrees * quadrant)));

            // the opposite one will be the base
            var @base = (height * Math.Abs(Math.Sin(ninetyDegrees * quadrant))) + (width * Math.Abs(Math.Cos(ninetyDegrees * quadrant)));

            var diagonalAngle = Math.Atan(perpendicular / @base);
            var diagonalLength = Math.Sqrt(Math.Pow(width, 2) + Math.Pow(height, 2));

            var vectorLength = (float)Math.Abs(diagonalLength * Math.Cos(diagonalAngle - (gradientAngle - (quadrant * ninetyDegrees))));

            var cosGradientAngle = (float)Math.Cos(gradientAngle);
            var sinGradientAngle = (float)Math.Sin(gradientAngle);

            var endPoint = quadrant switch
            {
                0 => new PointF(vectorLength * cosGradientAngle, vectorLength * sinGradientAngle),
                1 => new PointF(width - (vectorLength * cosGradientAngle), vectorLength * sinGradientAngle),
                2 => new PointF(width + (vectorLength * cosGradientAngle), height + (vectorLength * sinGradientAngle)),
                _ => new PointF(vectorLength * cosGradientAngle, height + (vectorLength * sinGradientAngle))
            };

            return endPoint;
        }
    }
}