using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using Microsoft.SqlServer.Types;
using USC.GISResearchLab.Common.Utils.Colors;

namespace USC.GISResearchLab.Common.Core.KML
{

    // this is from http://www.manfridayconsulting.it/index.php?option=com_content&view=article&id=31:generate-kml-file&catid=8:c&Itemid=22

    public class KMLDocument
    {
        public XmlDocument doc { get; set; }
        public XmlNode kmlNode { get; set; }
        public XmlNode documentNode { get; set; }
        private int _RandomCount;
        public int RandomCount
        {

            get
            {
                if (_RandomCount == 0)
                {
                    _RandomCount = 1;
                }
                else
                {
                    int newRandom = _RandomCount;
                    Thread.Sleep(2);
                    newRandom += new Random().Next();
                    while (newRandom == _RandomCount)
                    {
                        newRandom = new Random(newRandom).Next();
                    }

                    _RandomCount = newRandom;
                }

                return _RandomCount;
            }
        }

        public KMLDocument()
            : this(null)
        {
        }

        public KMLDocument(string name)
        {
            doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            doc.AppendChild(docNode);

            kmlNode = doc.CreateElement("kml");

            XmlAttribute xmlnsAttribute = doc.CreateAttribute("xmlns");
            xmlnsAttribute.Value = "http://www.opengis.net/kml/2.2";
            kmlNode.Attributes.Append(xmlnsAttribute);
            doc.AppendChild(kmlNode);

            documentNode = doc.CreateElement("Document");
            kmlNode.AppendChild(documentNode);

        }

        public string AsString()
        {
            string ret = "";
            if (doc != null)
            {
                StringWriter writer = new StringWriter();
                doc.Save(writer);
                ret = writer.ToString();
            }
            return ret;
        }

        public void WriteToStream(Stream stream, Encoding encoding)
        {
            XmlTextWriter xmlTextWriter = new XmlTextWriter(stream, encoding);
            WriteToXmlTextWriter(xmlTextWriter);
        }

        public void WriteToXmlTextWriter(XmlTextWriter xmlTextWriter)
        {
            doc.Save(xmlTextWriter);
            xmlTextWriter.Flush();
            xmlTextWriter.Close();
        }

        public override string ToString()
        {
            return doc.OuterXml;
        }

        //here we can save the document in a file, so it can be explored with google earth 
        public void SaveKml(string filePath)
        {
            try
            {
                doc.Save(filePath);
            }
            catch (Exception e)
            {
                throw new Exception("Error in SaveKml: " + e.Message, e);
            }
        }

        public void AddStyle(string name, Color lineColor, double lineWidth, Color polygonColor, bool polygonFill, bool polygonOutline)
        {
            // get the hex values
            string lineColorHex = ColorUtils.ColorToHexString(lineColor);
            string polygonColorHex = ColorUtils.ColorToHexString(polygonColor);

            // then add the alpha
            lineColorHex = "ff" + lineColorHex;
            polygonColorHex = "ff" + polygonColorHex;

            AddStyle(name, lineColorHex, lineWidth, polygonColorHex, polygonFill, polygonOutline);

        }

        public void AddStyle(string name, string lineColor, double lineWidth, string polygonColor, bool polygonFill, bool polygonOutline)
        {
            XmlNode StyleNode = doc.CreateElement("Style");
            XmlAttribute styleAttribute = doc.CreateAttribute("id");
            styleAttribute.Value = name;
            StyleNode.Attributes.Append(styleAttribute);
            documentNode.AppendChild(StyleNode);

            if (!String.IsNullOrEmpty(polygonColor))
            {
                // add the alpha to the front
                if (polygonColor.Length == 6)
                {
                    polygonColor = "ff" + polygonColor;
                }

                XmlNode PolyStyleNode = doc.CreateElement("PolyStyle");
                StyleNode.AppendChild(PolyStyleNode);

                XmlNode colorNode = doc.CreateElement("color");
                colorNode.AppendChild(doc.CreateTextNode(polygonColor));
                PolyStyleNode.AppendChild(colorNode);

                XmlNode fillNode = doc.CreateElement("fill");
                fillNode.AppendChild(doc.CreateTextNode(Convert.ToInt32(polygonFill).ToString()));
                PolyStyleNode.AppendChild(fillNode);

                XmlNode outlineNode = doc.CreateElement("outline");
                outlineNode.AppendChild(doc.CreateTextNode(Convert.ToInt32(polygonOutline).ToString()));
                PolyStyleNode.AppendChild(outlineNode);
            }

            if (!String.IsNullOrEmpty(lineColor))
            {
                // add the alpha to the front
                if (lineColor.Length == 6)
                {
                    lineColor = "ff" + lineColor;
                }

                XmlNode lineStyleNode = doc.CreateElement("LineStyle");
                StyleNode.AppendChild(lineStyleNode);

                XmlNode colorNode = doc.CreateElement("color");
                colorNode.AppendChild(doc.CreateTextNode(lineColor));
                lineStyleNode.AppendChild(colorNode);

                XmlNode widthNode = doc.CreateElement("width");
                widthNode.AppendChild(doc.CreateTextNode(lineWidth.ToString()));
                lineStyleNode.AppendChild(widthNode);
            }

        }


        public void AddStyleLine(string name, Color lineColor, double lineWidth)
        {
            // get the hex values
            string lineColorHex = ColorUtils.ColorToHexString(lineColor);

            // then add the alpha
            lineColorHex = "ff" + lineColorHex;

            AddStyleLine(name, lineColorHex, lineWidth);

        }

        public void AddStyleLine(string name, string lineColor, double lineWidth)
        {
            XmlNode StyleNode = doc.CreateElement("Style");
            XmlAttribute styleAttribute = doc.CreateAttribute("id");
            styleAttribute.Value = name;
            StyleNode.Attributes.Append(styleAttribute);
            documentNode.AppendChild(StyleNode);

            if (!String.IsNullOrEmpty(lineColor))
            {
                XmlNode lineStyleNode = doc.CreateElement("LineStyle");
                StyleNode.AppendChild(lineStyleNode);

                XmlNode colorNode = doc.CreateElement("color");
                colorNode.AppendChild(doc.CreateTextNode(lineColor));
                lineStyleNode.AppendChild(colorNode);

                XmlNode widthNode = doc.CreateElement("width");
                widthNode.AppendChild(doc.CreateTextNode(lineWidth.ToString()));
                lineStyleNode.AppendChild(widthNode);
            }

        }

        public void AddStylePolygon(string name, Color polygonColor, bool polygonFill, bool polygonOutline, double outlineWidth)
        {
            // get the hex values
            string polygonColorHex = ColorUtils.ColorToHexString(polygonColor);

            // then add the alpha
            polygonColorHex = "ff" + polygonColorHex;

            AddStylePolygon(name, polygonColorHex, polygonFill, polygonOutline, outlineWidth);

        }

        public void AddStylePolygon(string name, string polygonColor, bool polygonFill, bool polygonOutline, double outlineWidth)
        {
            XmlNode StyleNode = doc.CreateElement("Style");
            XmlAttribute styleAttribute = doc.CreateAttribute("id");
            styleAttribute.Value = name;
            StyleNode.Attributes.Append(styleAttribute);
            documentNode.AppendChild(StyleNode);

            if (!String.IsNullOrEmpty(polygonColor))
            {
                // set the line width
                XmlNode lineStyleNode = doc.CreateElement("LineStyle");
                StyleNode.AppendChild(lineStyleNode);

                XmlNode widthNode = doc.CreateElement("width");
                widthNode.AppendChild(doc.CreateTextNode(outlineWidth.ToString()));
                lineStyleNode.AppendChild(widthNode);

                XmlNode colorNode = doc.CreateElement("color");
                colorNode.AppendChild(doc.CreateTextNode(polygonColor));
                lineStyleNode.AppendChild(colorNode);

                // set the polygon attributes
                XmlNode PolyStyleNode = doc.CreateElement("PolyStyle");
                StyleNode.AppendChild(PolyStyleNode);

                XmlNode colorNode2 = doc.CreateElement("color");
                colorNode2.AppendChild(doc.CreateTextNode(polygonColor));
                PolyStyleNode.AppendChild(colorNode2);

                if (polygonFill)
                {
                    XmlNode fillNode = doc.CreateElement("fill");
                    fillNode.AppendChild(doc.CreateTextNode(Convert.ToInt32(polygonFill).ToString()));
                    PolyStyleNode.AppendChild(fillNode);
                }
                if (!polygonFill)
                {
                    XmlNode fillNode = doc.CreateElement("fill");
                    fillNode.AppendChild(doc.CreateTextNode(Convert.ToInt32(polygonFill).ToString()));
                    PolyStyleNode.AppendChild(fillNode);
                }

                if (polygonOutline)
                {
                    XmlNode outlineNode = doc.CreateElement("outline");
                    outlineNode.AppendChild(doc.CreateTextNode(Convert.ToInt32(polygonOutline).ToString()));
                    PolyStyleNode.AppendChild(outlineNode);
                }
            }
        }

        public void AddPoint(float lat, float lon, string name, string description)
        {
            AddPoint(lat, lon, 0, name, description);
        }

        public void AddPoint(float lat, float lon, float elevation, string name, string description)
        {
            XmlNode PlacemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(PlacemarkNode);
            
            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            PlacemarkNode.AppendChild(nameNode);

            if (!String.IsNullOrEmpty(description))
            {
                XmlNode descriptionNode = doc.CreateElement("description");
                descriptionNode.AppendChild(doc.CreateTextNode(description));
                PlacemarkNode.AppendChild(descriptionNode);
            }

            XmlNode PointNode = doc.CreateElement("Point");
            PlacemarkNode.AppendChild(PointNode);
            XmlNode coordinateNode = doc.CreateElement("coordinates");
            PointNode.AppendChild(coordinateNode);
            coordinateNode.AppendChild(doc.CreateTextNode(lon.ToString(new System.Globalization.CultureInfo("en-US")) + "," + lat.ToString(new System.Globalization.CultureInfo("en-US")) + "," + elevation.ToString(new System.Globalization.CultureInfo("en-US"))));

        }
        public void AddPoint(float lat, float lon, string name, string description, string recordId)
        {
            AddPoint(lat, lon, 0, name, description, recordId);
        }

        public void AddPoint(float lat, float lon, float elevation, string name, string description, string recordId)
        {
            XmlNode PlacemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(PlacemarkNode);

            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            PlacemarkNode.AppendChild(nameNode);


            if (!String.IsNullOrEmpty(description))
            {
                XmlNode dataNode = doc.CreateElement("ExtendedData");
                PlacemarkNode.AppendChild(dataNode);

                XmlElement dataRecordNode = doc.CreateElement("Data");
                dataRecordNode.SetAttribute("name", "recordId");

                XmlNode dataRecordElement = doc.CreateElement("value");

                dataRecordElement.AppendChild(doc.CreateTextNode(recordId));
                dataRecordNode.AppendChild(dataRecordElement);

                XmlElement dataSourceNode = doc.CreateElement("Data");
                dataSourceNode.SetAttribute("name", "source");

                XmlNode dataSourceElement = doc.CreateElement("value");

                dataSourceElement.AppendChild(doc.CreateTextNode(description));
                dataSourceNode.AppendChild(dataSourceElement);


                dataNode.AppendChild(dataRecordNode);
                dataNode.AppendChild(dataSourceNode);

                PlacemarkNode.AppendChild(dataNode);

            }
            XmlNode PointNode = doc.CreateElement("Point");
            PlacemarkNode.AppendChild(PointNode);
            XmlNode coordinateNode = doc.CreateElement("coordinates");
            PointNode.AppendChild(coordinateNode);
            coordinateNode.AppendChild(doc.CreateTextNode(lon.ToString(new System.Globalization.CultureInfo("en-US")) + "," + lat.ToString(new System.Globalization.CultureInfo("en-US")) + "," + elevation.ToString(new System.Globalization.CultureInfo("en-US"))));

        }

        public void AddSqlGeography(SqlGeography sqlGeography, string name, string styleName, string description, string recordId)
        {
            try
            {
                if (sqlGeography != null)
                {
                    string geographyType = sqlGeography.STGeometryType().Value;

                    switch (geographyType)
                    {
                        case "Point":
                            AddSqlGeographyPoint(sqlGeography, name, styleName, description, recordId);
                            break;
                        case "LineString":
                            AddSqlGeographyLine(sqlGeography, name, styleName, description, recordId);
                            break;
                        case "Polygon":
                            AddSqlGeographyPolygon(sqlGeography, name, styleName, description, recordId);
                            break;
                        case "MultiPoint":
                        case "MultiLineString":
                        case "MultiPolygon":
                            AddSqlGeographyMulti(sqlGeography, name, styleName, description, recordId);
                            break;
                        case "GeometryCollection":
                            AddSqlGeographyGeometryCollection(sqlGeography, name, description, recordId);
                            break;
                        default:
                            throw new InvalidOperationException("Unknown geography type: " + geographyType);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error in AddSqlGeography: " + e.Message, e);
            }
        }

        public void AddSqlGeography(SqlGeography sqlGeography, string name, string description, string recordId)
        {
            try
            {
                if (sqlGeography != null)
                {
                    string geographyType = sqlGeography.STGeometryType().Value;

                    switch (geographyType)
                    {
                        case "Point":
                            AddSqlGeographyPoint(sqlGeography, name, description, recordId);
                            break;
                        case "LineString":
                            AddSqlGeographyLine(sqlGeography, name, description, recordId);
                            break;
                        case "Polygon":
                            AddSqlGeographyPolygon(sqlGeography, name, description, recordId);
                            break;
                        case "MultiPoint":
                        case "MultiLineString":
                        case "MultiPolygon":
                            AddSqlGeographyMulti(sqlGeography, name, description, recordId);
                            break;
                        case "GeometryCollection":
                            AddSqlGeographyGeometryCollection(sqlGeography, name, description, recordId);
                            break;
                        default:
                            throw new InvalidOperationException("Unknown geography type: " + geographyType);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error in AddSqlGeography: " + e.Message, e);
            }
        }

        public void AddSqlGeographyMulti(SqlGeography sqlGeographyMulti, string name, string styleName, string description, string recordId)
        {
            if (sqlGeographyMulti != null && !sqlGeographyMulti.IsNull)
            {
                int numberOfItems = sqlGeographyMulti.STNumGeometries().Value;
                for (int i = 1; i <= numberOfItems; i++)
                {
                    SqlGeography sqlGeography = sqlGeographyMulti.STGeometryN(i);
                    if (sqlGeography != null && !sqlGeography.IsNull)
                    {
                        string geographyType = sqlGeography.STGeometryType().Value;

                        switch (geographyType)
                        {
                            case "Point":
                                AddSqlGeographyPoint(sqlGeography, name, styleName, description, recordId);
                                break;
                            case "LineString":
                                AddSqlGeographyLine(sqlGeography, name, styleName, description, recordId);
                                break;
                            case "Polygon":
                                AddSqlGeographyPolygon(sqlGeography, name, styleName, description, recordId);
                                break;
                            default:
                                throw new InvalidOperationException("Unknown or unImplemented geography type: " + geographyType);
                        }
                    }
                }
            }
        }

        public void AddSqlGeographyMulti(SqlGeography sqlGeographyMulti, string name, string description, string recordId)
        {
            if (sqlGeographyMulti != null && !sqlGeographyMulti.IsNull)
            {
                //Color color = ColorUtils.RandomColor();
                //string styleName = "style" + color.ToArgb();

                List<SqlGeography> sqlGeographyList = new List<SqlGeography>();
                int numberOfItems = sqlGeographyMulti.STNumGeometries().Value;
                for (int i = 1; i <= numberOfItems; i++)
                {
                    SqlGeography sqlGeography = sqlGeographyMulti.STGeometryN(i);
                    if (sqlGeography != null && !sqlGeography.IsNull)
                    {
                        sqlGeographyList.Add(sqlGeography);
                    }
                }


                SqlGeography sqlGeographyFirst = sqlGeographyMulti.STGeometryN(1);
                if (sqlGeographyFirst != null && !sqlGeographyFirst.IsNull)
                {
                    string geographyType = sqlGeographyFirst.STGeometryType().Value;

                    switch (geographyType)
                    {
                        case "Point":
                            AddSqlGeographyPoints(sqlGeographyList, name, null, description, recordId);
                            break;
                        case "LineString":
                            //if (i == 1)
                            //{
                            //    AddStyleLine(styleName, color, 3.0);
                            //}

                            AddSqlGeographyLines(sqlGeographyList, name,  description, recordId);
                            break;
                        case "Polygon":
                            //if (i == 1)
                            //{
                            //    AddStylePolygon(styleName, color, false, true, 3.0);
                            //}

                            AddSqlGeographyPolygons(sqlGeographyList, name,  description, recordId);
                            break;
                        default:
                            throw new InvalidOperationException("Unknown or unImplemented geography type: " + geographyType);
                    }
                }

            }
        }

        public void AddSqlGeographyGeometryCollection(SqlGeography sqlGeographyGeometryCollection, string name, string description, string recordId)
        {
            Color color = ColorUtils.RandomColor(RandomCount);
            string styleName = "style" + color.ToArgb();
            AddStylePolygon(styleName, color, false, true, 3.0);
            AddSqlGeographyGeometryCollection(sqlGeographyGeometryCollection, name, styleName, description, recordId);
        }



        public void AddSqlGeographyGeometryCollection(SqlGeography sqlGeographyGeometryCollection, string name, string styleName, string description, string recordId)
        {
            XmlNode placemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(placemarkNode);

            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            placemarkNode.AppendChild(nameNode);

            if (!String.IsNullOrEmpty(description))
            {
                XmlNode dataNode = doc.CreateElement("ExtendedData");
                placemarkNode.AppendChild(dataNode);

                XmlElement dataRecordNode = doc.CreateElement("Data");
                dataRecordNode.SetAttribute("name", "recordId");

                XmlNode dataRecordElement = doc.CreateElement("value");

                dataRecordElement.AppendChild(doc.CreateTextNode(recordId));
                dataRecordNode.AppendChild(dataRecordElement);

                XmlElement dataSourceNode = doc.CreateElement("Data");
                dataSourceNode.SetAttribute("name", "source");

                XmlNode dataSourceElement = doc.CreateElement("value");

                dataSourceElement.AppendChild(doc.CreateTextNode(description));
                dataSourceNode.AppendChild(dataSourceElement);

                dataNode.AppendChild(dataRecordNode);
                dataNode.AppendChild(dataSourceNode);

                placemarkNode.AppendChild(dataNode);

            }

            if (!String.IsNullOrEmpty(styleName))
            {
                XmlNode styleNode = doc.CreateElement("styleUrl");
                styleNode.AppendChild(doc.CreateTextNode("#" + styleName));
                placemarkNode.AppendChild(styleNode);
            }

            XmlNode multiGeometryNode = null;


            multiGeometryNode = doc.CreateElement("MultiGeometry");
            placemarkNode.AppendChild(multiGeometryNode);


            int numberOfItems = sqlGeographyGeometryCollection.STNumGeometries().Value;
            for (int i = 1; i <= numberOfItems; i++)
            {
                SqlGeography sqlGeography = sqlGeographyGeometryCollection.STGeometryN(i);
                if (sqlGeography != null)
                {
                    string geographyType = sqlGeography.STGeometryType().Value;

                    switch (geographyType)
                    {
                        case "Point":
                            AppendPointNode(multiGeometryNode, sqlGeography);
                            break;
                        case "LineString":
                            AppendLineNode(multiGeometryNode, sqlGeography);
                            break;
                        case "Polygon":
                            AppendPolygonNode(multiGeometryNode, sqlGeography);
                            break;
                        case "MultiPoint":
                        case "MultiLineString":
                        case "MultiPolygon":
                        case "GeometryCollection":
                            throw new NotImplementedException("UnImplemented geography type within a GeometryCollection: " + geographyType);
                        default:
                            throw new InvalidOperationException("Unknown geography type: " + geographyType);
                    }
                }
            }
        }

        public void AddSqlGeographyPolygon(SqlGeography sqlGeographyPolygon, string name, string description, string recordId)
        {
            Color color = ColorUtils.RandomColor(RandomCount);
            string styleName = "style" + color.ToArgb();
            AddStylePolygon(styleName, color, false, true, 3.0);

            List<SqlGeography> sqlGeographies = new List<SqlGeography>();
            sqlGeographies.Add(sqlGeographyPolygon);
            AddSqlGeographyPolygons(sqlGeographies, name, styleName, description, recordId);
        }

         public void AddSqlGeographyPolygon(SqlGeography sqlGeographyPolygon, string name, string styleName, string description, string recordId)
         {
             List<SqlGeography> sqlGeographies = new List<SqlGeography>();
             sqlGeographies.Add(sqlGeographyPolygon);
             AddSqlGeographyPolygons(sqlGeographies, name, styleName, description, recordId);
         }

        public void AddSqlGeographyPolygons(List<SqlGeography> sqlGeographyPolygons, string name, string description,string recordId)
        {
            Color color = ColorUtils.RandomColor(RandomCount);
            string styleName = "style" + color.ToArgb();
            AddStylePolygon(styleName, color, false, true, 3.0);

            AddSqlGeographyPolygons(sqlGeographyPolygons, name, styleName, description, recordId);
        }

        public void AddSqlGeographyPolygons(List<SqlGeography> sqlGeographyPolygons, string name, string styleName, string description, string recordId)
        {
            XmlNode PlacemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(PlacemarkNode);

            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            PlacemarkNode.AppendChild(nameNode);


            if (!String.IsNullOrEmpty(description))
            {
                XmlNode dataNode = doc.CreateElement("ExtendedData");
                PlacemarkNode.AppendChild(dataNode);

                XmlElement dataRecordNode = doc.CreateElement("Data");
                dataRecordNode.SetAttribute("name", "recordId");

                XmlNode dataRecordElement = doc.CreateElement("value");

                dataRecordElement.AppendChild(doc.CreateTextNode(recordId));
                dataRecordNode.AppendChild(dataRecordElement);

                XmlElement dataSourceNode = doc.CreateElement("Data");
                dataSourceNode.SetAttribute("name", "source");

                XmlNode dataSourceElement = doc.CreateElement("value");

                dataSourceElement.AppendChild(doc.CreateTextNode(description));
                dataSourceNode.AppendChild(dataSourceElement);


                dataNode.AppendChild(dataRecordNode);
                dataNode.AppendChild(dataSourceNode);

                PlacemarkNode.AppendChild(dataNode);

            }

            if (!String.IsNullOrEmpty(styleName))
            {
                XmlNode styleNode = doc.CreateElement("styleUrl");
                styleNode.AppendChild(doc.CreateTextNode("#" + styleName));
                PlacemarkNode.AppendChild(styleNode);
            }

            XmlNode multiGeometryNode = null;

            if (sqlGeographyPolygons.Count > 1)
            {
                multiGeometryNode = doc.CreateElement("MultiGeometry");
                PlacemarkNode.AppendChild(multiGeometryNode);
            }
            else
            {
                multiGeometryNode = PlacemarkNode;
            }


            foreach (SqlGeography sqlGeographyPolygon in sqlGeographyPolygons)
            {

                AppendPolygonNode(multiGeometryNode, sqlGeographyPolygon);

                //XmlNode polygonNode = doc.CreateElement("Polygon");
                //multiGeometryNode.AppendChild(polygonNode);

                //XmlNode extrudeNode = doc.CreateElement("extrude");
                //polygonNode.AppendChild(extrudeNode);
                //extrudeNode.AppendChild(doc.CreateTextNode("1"));

                //XmlNode tessellateNode = doc.CreateElement("tessellate");
                //polygonNode.AppendChild(tessellateNode);
                //tessellateNode.AppendChild(doc.CreateTextNode("1"));

                //XmlNode altitudeModeNode = doc.CreateElement("altitudeMode");
                //polygonNode.AppendChild(altitudeModeNode);
                //altitudeModeNode.AppendChild(doc.CreateTextNode("absolute"));


                //int numberOfRings = sqlGeographyPolygon.NumRings().Value;

                //for (int r = 1; r < sqlGeographyPolygon.NumRings().Value + 1; r++)
                //{
                //    SqlGeography ring = sqlGeographyPolygon.RingN(r);

                //    XmlNode boundaryIsNode = null;

                //    // first ring is the exterior ring, all others are interior
                //    if (r == 1)
                //    {
                //        boundaryIsNode = doc.CreateElement("outerBoundaryIs");

                //    }
                //    else
                //    {
                //        boundaryIsNode = doc.CreateElement("innerBoundaryIs");
                //    }

                //    polygonNode.AppendChild(boundaryIsNode);

                //    XmlNode linearRingNode = doc.CreateElement("LinearRing");
                //    boundaryIsNode.AppendChild(linearRingNode);

                //    XmlNode coordinatesNode = doc.CreateElement("coordinates");
                //    linearRingNode.AppendChild(coordinatesNode);


                //    string coordinates = "";

                //    coordinates += ring.STStartPoint().Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //    coordinates += "," + ring.STStartPoint().Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //    if (!ring.STStartPoint().Z.IsNull)
                //    {
                //        coordinates += "," + ring.STStartPoint().Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //    }
                //    else
                //    {
                //        coordinates += ",0";
                //    }

                //    for (int i = 2; i < ring.STNumPoints().Value + 1; i++)
                //    {
                //        coordinates += Environment.NewLine;

                //        coordinates += ring.STPointN(i).Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        coordinates += "," + ring.STPointN(i).Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        if (!ring.STPointN(i).Z.IsNull)
                //        {
                //            coordinates += "," + ring.STPointN(i).Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        }
                //        else
                //        {
                //            coordinates += ",0";
                //        }
                //    }

                //    if (ring.STStartPoint().Lat.Value != ring.STEndPoint().Lat.Value || ring.STStartPoint().Long.Value != ring.STEndPoint().Long.Value)
                //    {
                //        coordinates += Environment.NewLine;

                //        coordinates += ring.STEndPoint().Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        coordinates += "," + ring.STEndPoint().Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        if (!ring.STEndPoint().Z.IsNull)
                //        {
                //            coordinates += "," + ring.STEndPoint().Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //        }
                //        else
                //        {
                //            coordinates += ",0";
                //        }
                //    }

                //    coordinatesNode.AppendChild(doc.CreateTextNode(coordinates));

                //}
            }
        }

        public void AppendPolygonNode(XmlNode parentNode, SqlGeography sqlGeographyPolygon)
        {
            XmlNode polygonNode = doc.CreateElement("Polygon");
            parentNode.AppendChild(polygonNode);

            XmlNode extrudeNode = doc.CreateElement("extrude");
            polygonNode.AppendChild(extrudeNode);
            extrudeNode.AppendChild(doc.CreateTextNode("2"));

            XmlNode tessellateNode = doc.CreateElement("tessellate");
            polygonNode.AppendChild(tessellateNode);
            tessellateNode.AppendChild(doc.CreateTextNode("1"));

            XmlNode altitudeModeNode = doc.CreateElement("altitudeMode");
            polygonNode.AppendChild(altitudeModeNode);
            altitudeModeNode.AppendChild(doc.CreateTextNode("relativeToGround"));


            int numberOfRings = sqlGeographyPolygon.NumRings().Value;

            for (int r = 1; r < sqlGeographyPolygon.NumRings().Value + 1; r++)
            {
                SqlGeography ring = sqlGeographyPolygon.RingN(r);

                XmlNode boundaryIsNode = null;

                // first ring is the exterior ring, all others are interior
                if (r == 1)
                {
                    boundaryIsNode = doc.CreateElement("outerBoundaryIs");

                }
                else
                {
                    boundaryIsNode = doc.CreateElement("innerBoundaryIs");
                }

                polygonNode.AppendChild(boundaryIsNode);

                XmlNode linearRingNode = doc.CreateElement("LinearRing");
                boundaryIsNode.AppendChild(linearRingNode);

                XmlNode coordinatesNode = doc.CreateElement("coordinates");
                linearRingNode.AppendChild(coordinatesNode);


                string coordinates = "";

                coordinates += ring.STStartPoint().Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                coordinates += "," + ring.STStartPoint().Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                if (!ring.STStartPoint().Z.IsNull)
                {
                    coordinates += "," + ring.STStartPoint().Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                }
                else
                {
                    coordinates += ",0";
                }

                for (int i = 2; i < ring.STNumPoints().Value + 1; i++)
                {
                    coordinates += Environment.NewLine;

                    coordinates += ring.STPointN(i).Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    coordinates += "," + ring.STPointN(i).Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    if (!ring.STPointN(i).Z.IsNull)
                    {
                        coordinates += "," + ring.STPointN(i).Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    }
                    else
                    {
                        coordinates += ",0";
                    }
                }

                if (ring.STStartPoint().Lat.Value != ring.STEndPoint().Lat.Value || ring.STStartPoint().Long.Value != ring.STEndPoint().Long.Value)
                {
                    coordinates += Environment.NewLine;

                    coordinates += ring.STEndPoint().Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    coordinates += "," + ring.STEndPoint().Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    if (!ring.STEndPoint().Z.IsNull)
                    {
                        coordinates += "," + ring.STEndPoint().Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                    }
                    else
                    {
                        coordinates += ",0";
                    }
                }

                coordinatesNode.AppendChild(doc.CreateTextNode(coordinates));

            }
        }

        public void AddSqlGeographyPoint(SqlGeography sqlGeography, string name, string description, string recordId)
        {
            List<SqlGeography> sqlGeographies = new List<SqlGeography>();
            sqlGeographies.Add(sqlGeography);
            AddSqlGeographyPoints(sqlGeographies, name, null, description, recordId);
        }

        public void AddSqlGeographyPoint( SqlGeography  sqlGeography, string name, string styleName, string description, string recordId)
        {
            List<SqlGeography> sqlGeographies = new List<SqlGeography>();
            sqlGeographies.Add(sqlGeography);
            AddSqlGeographyPoints(sqlGeographies, name, styleName, description, recordId);
        }

        public void AddSqlGeographyPoints(List<SqlGeography> sqlGeographyPoints, string name, string description,string recordId)
        {
            AddSqlGeographyPoints(sqlGeographyPoints, name, null, description, recordId);
        }

        public void AddSqlGeographyPoints(List<SqlGeography> sqlGeographies, string name, string styleName, string description, string recordId)
        {
            XmlNode PlacemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(PlacemarkNode);

            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            PlacemarkNode.AppendChild(nameNode);

            if (!String.IsNullOrEmpty(description))
            {
                XmlNode dataNode = doc.CreateElement("ExtendedData");
                PlacemarkNode.AppendChild(dataNode);

                XmlElement dataRecordNode = doc.CreateElement("Data");
                dataRecordNode.SetAttribute("name", "recordId");

                XmlNode dataRecordElement = doc.CreateElement("value");

                dataRecordElement.AppendChild(doc.CreateTextNode(recordId));
                dataRecordNode.AppendChild(dataRecordElement);

                XmlElement dataSourceNode = doc.CreateElement("Data");
                dataSourceNode.SetAttribute("name", "source");

                XmlNode dataSourceElement = doc.CreateElement("value");
                dataSourceElement.AppendChild(doc.CreateTextNode(description));
                dataSourceNode.AppendChild(dataSourceElement);

                dataNode.AppendChild(dataRecordNode);
                dataNode.AppendChild(dataSourceNode);

                PlacemarkNode.AppendChild(dataNode);

            }

            if (!String.IsNullOrEmpty(styleName))
            {
                XmlNode styleNode = doc.CreateElement("styleUrl");
                styleNode.AppendChild(doc.CreateTextNode("#" + styleName));
                PlacemarkNode.AppendChild(styleNode);
            }

            XmlNode multiGeometryNode = null;

            if (sqlGeographies.Count > 1)
            {
                multiGeometryNode = doc.CreateElement("MultiGeometry");
                PlacemarkNode.AppendChild(multiGeometryNode);
            }
            else
            {
                multiGeometryNode = PlacemarkNode;
            }


            foreach (SqlGeography sqlGeography in sqlGeographies)
            {

                AppendPointNode(multiGeometryNode, sqlGeography);

                //XmlNode PointNode = doc.CreateElement("Point");
                //multiGeometryNode.AppendChild(PointNode);

                //XmlNode coordinateNode = doc.CreateElement("coordinates");
                //PointNode.AppendChild(coordinateNode);

                //string coordinates = "";

                //coordinates += sqlGeography.Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //coordinates += "," + sqlGeography.Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));

                //if (!sqlGeography.Z.IsNull)
                //{
                //    coordinates += "," + sqlGeography.Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //}
                //else
                //{
                //    coordinates += ",0";
                //}

                //coordinateNode.AppendChild(doc.CreateTextNode(coordinates));
            }

        }

        public void AppendPointNode(XmlNode parentNode, SqlGeography sqlGeographyPoint)
        {
            XmlNode PointNode = doc.CreateElement("Point");
            parentNode.AppendChild(PointNode);

            XmlNode coordinateNode = doc.CreateElement("coordinates");
            PointNode.AppendChild(coordinateNode);

            string coordinates = "";

            coordinates += sqlGeographyPoint.Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
            coordinates += "," + sqlGeographyPoint.Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));

            if (!sqlGeographyPoint.Z.IsNull)
            {
                coordinates += "," + sqlGeographyPoint.Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
            }
            else
            {
                coordinates += ",0";
            }

            coordinateNode.AppendChild(doc.CreateTextNode(coordinates));
        }

        public void AddSqlGeographyLine(SqlGeography sqlGeography, string name, string description, string recordId)
        {
            Color color = ColorUtils.RandomColor(RandomCount);
            string styleName = "style" + color.ToArgb();
            AddStyleLine(styleName, color, 3.0);

            List<SqlGeography> sqlGeographies = new List<SqlGeography>();
            sqlGeographies.Add(sqlGeography);
            AddSqlGeographyLines(sqlGeographies, name, styleName, description, recordId);
        }

        public void AddSqlGeographyLine(SqlGeography sqlGeography, string name, string styleName, string description, string recordId)
        {
            List<SqlGeography> sqlGeographies = new List<SqlGeography>();
            sqlGeographies.Add(sqlGeography);
            AddSqlGeographyLines(sqlGeographies, name, styleName, description, recordId);
        }

        public void AddSqlGeographyLines(List<SqlGeography> sqlGeographyLines, string name, string description, string recordId)
        {
            Color color = ColorUtils.RandomColor(RandomCount);
            string styleName = "style" + color.ToArgb();
            AddStyleLine(styleName, color, 3.0);

            AddSqlGeographyLines(sqlGeographyLines, name, styleName, description, recordId);
        }

        public void AddSqlGeographyLines(List<SqlGeography> sqlGeographies, string name, string styleName, string description,string recordId)
        {
            XmlNode PlacemarkNode = doc.CreateElement("Placemark");
            documentNode.AppendChild(PlacemarkNode);

            XmlNode nameNode = doc.CreateElement("name");
            nameNode.AppendChild(doc.CreateTextNode(name));
            PlacemarkNode.AppendChild(nameNode);

            if (!String.IsNullOrEmpty(description))
            {
                XmlNode dataNode = doc.CreateElement("ExtendedData");
                PlacemarkNode.AppendChild(dataNode);

                XmlElement dataRecordNode = doc.CreateElement("Data");
                dataRecordNode.SetAttribute("name", "recordId");

                XmlNode dataRecordElement = doc.CreateElement("value");

                dataRecordElement.AppendChild(doc.CreateTextNode(recordId));
                dataRecordNode.AppendChild(dataRecordElement);

                XmlElement dataSourceNode = doc.CreateElement("Data");
                dataSourceNode.SetAttribute("name", "source");

                XmlNode dataSourceElement = doc.CreateElement("value");

                dataSourceElement.AppendChild(doc.CreateTextNode(description));
                dataSourceNode.AppendChild(dataSourceElement);

                dataNode.AppendChild(dataRecordNode);
                dataNode.AppendChild(dataSourceNode);

                PlacemarkNode.AppendChild(dataNode);

            }

            if (!String.IsNullOrEmpty(styleName))
            {
                XmlNode styleNode = doc.CreateElement("styleUrl");
                styleNode.AppendChild(doc.CreateTextNode("#" + styleName));
                PlacemarkNode.AppendChild(styleNode);
            }

            XmlNode multiGeometryNode = null;

            if (sqlGeographies.Count > 1)
            {
                multiGeometryNode = doc.CreateElement("MultiGeometry");
                PlacemarkNode.AppendChild(multiGeometryNode);
            }
            else
            {
                multiGeometryNode = PlacemarkNode;
            }


            foreach (SqlGeography sqlGeography in sqlGeographies)
            {

                AppendLineNode(multiGeometryNode, sqlGeography);

                //XmlNode LineNode = doc.CreateElement("LineString");
                //multiGeometryNode.AppendChild(LineNode);

                //XmlNode extrudeNode = doc.CreateElement("extrude");
                //LineNode.AppendChild(extrudeNode);
                //extrudeNode.AppendChild(doc.CreateTextNode("1"));

                //XmlNode tessellateNode = doc.CreateElement("tessellate");
                //LineNode.AppendChild(tessellateNode);
                //tessellateNode.AppendChild(doc.CreateTextNode("1"));

                //XmlNode altitudeModeNode = doc.CreateElement("altitudeMode");
                //LineNode.AppendChild(altitudeModeNode);
                //altitudeModeNode.AppendChild(doc.CreateTextNode("absolute"));

                //string coordinates = "";
                //int numberOfPoints = sqlGeography.STNumPoints().Value;
                //for (int i = 1; i < numberOfPoints + 1; i++)
                //{

                //    SqlGeography currentPoint = sqlGeography.STPointN(i);

                //    if (i > 1)
                //    {
                //        coordinates += Environment.NewLine;
                //    }

                //    coordinates += currentPoint.Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //    coordinates += "," + currentPoint.Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));

                //    if (!currentPoint.Z.IsNull)
                //    {
                //        coordinates += "," + currentPoint.Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                //    }
                //    else
                //    {
                //        coordinates += ",0";
                //    }
                //}


                //XmlNode coordinateNode = doc.CreateElement("coordinates");
                //LineNode.AppendChild(coordinateNode);
                //coordinateNode.AppendChild(doc.CreateTextNode(coordinates));
            }
        }

        public void AppendLineNode(XmlNode parentNode, SqlGeography sqlGeographyLine)
        {
            XmlNode LineNode = doc.CreateElement("LineString");
            parentNode.AppendChild(LineNode);

            XmlNode extrudeNode = doc.CreateElement("extrude");
            LineNode.AppendChild(extrudeNode);
            extrudeNode.AppendChild(doc.CreateTextNode("2"));

            XmlNode tessellateNode = doc.CreateElement("tessellate");
            LineNode.AppendChild(tessellateNode);
            tessellateNode.AppendChild(doc.CreateTextNode("1"));

            XmlNode altitudeModeNode = doc.CreateElement("altitudeMode");
            LineNode.AppendChild(altitudeModeNode);
            altitudeModeNode.AppendChild(doc.CreateTextNode("relativeToGround"));

            string coordinates = "";
            int numberOfPoints = sqlGeographyLine.STNumPoints().Value;
            for (int i = 1; i < numberOfPoints + 1; i++)
            {

                SqlGeography currentPoint = sqlGeographyLine.STPointN(i);

                if (i > 1)
                {
                    coordinates += Environment.NewLine;
                }

                coordinates += currentPoint.Long.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                coordinates += "," + currentPoint.Lat.Value.ToString(new System.Globalization.CultureInfo("en-US"));

                if (!currentPoint.Z.IsNull)
                {
                    coordinates += "," + currentPoint.Z.Value.ToString(new System.Globalization.CultureInfo("en-US"));
                }
                else
                {
                    coordinates += ",0";
                }
            }


            XmlNode coordinateNode = doc.CreateElement("coordinates");
            LineNode.AppendChild(coordinateNode);
            coordinateNode.AppendChild(doc.CreateTextNode(coordinates));
        }
    }
}
