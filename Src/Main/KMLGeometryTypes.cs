using System;

namespace USC.GISResearchLab.Common.Core.KML
{
    public enum KMLGeometryType { Unknown, Point, LineString, LinearRing, Polygon }

    public class KMLGeometryTypes
    {
        public static string KMLGeometryType_Name_Unknown = "Unknown";
        public static string KMLGeometryType_Name_Point = "Point";
        public static string KMLGeometryType_Name_LineString = "LineString";
        public static string KMLGeometryType_Name_LinearRing = "LinearRing";
        public static string KMLGeometryType_Name_Polygon = "Polygon";

        public static KMLGeometryType GetKMLGeometryTypeFromName(string name)
        {
            KMLGeometryType ret = KMLGeometryType.Unknown;

            if (String.Compare(name, KMLGeometryType_Name_LinearRing, true) == 0)
            {
                ret = KMLGeometryType.LinearRing;
            }
            else if (String.Compare(name, KMLGeometryType_Name_LineString, true) == 0)
            {
                ret = KMLGeometryType.LineString;
            }
            else if (String.Compare(name, KMLGeometryType_Name_Point, true) == 0)
            {
                ret = KMLGeometryType.Point;
            }
            else if (String.Compare(name, KMLGeometryType_Name_Polygon, true) == 0)
            {
                ret = KMLGeometryType.Polygon;
            }
            else
            {
                throw new Exception("Unexpected or implemented KMLGeometryType: " + name);
            }
            return ret;
        }

        public static string GetKMLGeometryTypeName(KMLGeometryType type)
        {
            string ret = null;

            switch (type)
            {
                case KMLGeometryType.LinearRing:
                    ret = KMLGeometryType_Name_LinearRing;
                    break;
                case KMLGeometryType.LineString:
                    ret = KMLGeometryType_Name_LineString;
                    break;
                case KMLGeometryType.Point:
                    ret = KMLGeometryType_Name_Point;
                    break;
                case KMLGeometryType.Polygon:
                    ret = KMLGeometryType_Name_Polygon;
                    break;
                case KMLGeometryType.Unknown:
                    ret = KMLGeometryType_Name_Unknown;
                    break;
                default:
                    throw new Exception("Unexpected or implemented KMLGeometryType: " + type);
            }

            return ret;
        }

    }
}
