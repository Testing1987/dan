using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Alignment_mdi
{
    class Coord_conv
    {
        Point3d Convert_coordinate_to_ll84(double n, double e, string coord_sys_code)
        {
            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            int cscount1 = Dictionary1.GetSize();
            OSGeo.MapGuide.MgStringCollection col_names = Enum1.NextName(cscount1);

            string good_cs = "LL84";

            for (int i = 0; i < cscount1; ++i)
            {
                string nume1 = col_names.GetItem(i);
                if (nume1.ToLower() == coord_sys_code.ToLower())
                {
                    good_cs = nume1;
                }
            }

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Dictionary1.GetCoordinateSystem(good_cs);
            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem("LL84");
            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(e, n);
            return new Point3d(Coord1.X, Coord1.Y, 0);
        }

        Point3d Convert_coordinate_to_new_CS(Point3d Point1)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            Point3d Point2 = new Point3d();
            string Curent_system = Acmap.GetMapSRS();
            if (string.IsNullOrEmpty(Curent_system) == true)
            {
                MessageBox.Show("Please set your coordinate system");
                return Point2;
            }

            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem("LL84");

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);
            return Point2;
        }

        public static string get_ll84(Point2d pt1)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            string Curent_system = Acmap.GetMapSRS();

            if (string.IsNullOrEmpty(Curent_system) == true)
            {
                return "0,0,0";
            }


            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem("LL84");
            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(pt1.X, pt1.Y);
            return Convert.ToString(Coord1.X) + "," + Convert.ToString(Coord1.Y) + ",0";
        }

        public static string build_KML_string(string kml_name, string kml_line_style, System.Data.DataTable dt1)
        {
            string q = "\u0022";

            string line1 = "<?xml version=" + q + "1.0" + q + " encoding=" + q + "UTF-8" + q + "?>";
            string line2 = "<kml xmlns=" + q + "http://www.opengis.net/kml/2.2" + q + " xmlns:gx=" + q + "http://www.google.com/kml/ext/2.2" + q + " xmlns:kml=" + q + "http://www.opengis.net/kml/2.2" + q + " xmlns:atom=" + q + "http://www.w3.org/2005/Atom" + q + ">";
            string line3 = "<Document >";
            string line4 = "	<name>" + kml_name + ".kml</name>";
            string line5 = "	<open>1</open>";

            string line6 = kml_line_style;


            string line7 = "";


            for (int k = 0; k < dt1.Rows.Count; ++k)
            {
                if (Convert.ToInt32(dt1.Rows[k][3]) > 0)
                {
                    string l1 = "	<Folder>";

                    string l2 = "	<name>" + Convert.ToString(dt1.Rows[k][0]) + "</name>";
                    string l3 = "	<open>1</open>";

                    string l4 = Convert.ToString(dt1.Rows[k][2]);

                    string l5 = "	</Folder>";

                    if (line7 == "")
                    {
                        line7 = l1 + "\r\n" + l2 + "\r\n" + l3 + "\r\n" + l4 + "\r\n" + l5;
                    }
                    else
                    {
                        line7 = line7 + "\r\n" + l1 + "\r\n" + l2 + "\r\n" + l3 + "\r\n" + l4 + "\r\n" + l5;
                    }
                }
            }

            string line8 = "</Document>";
            string line9 = "</kml>";


            string string1 = line1 + "\r\n" + line2 + "\r\n" + line3 + "\r\n" + line4 + "\r\n" + line5 + "\r\n" + line6 + "\r\n" + line7 + "\r\n" + line8 + "\r\n" + line9;

            return string1;
        }

        public static string get_kml_style(string nume1, int colorindex, int width1)
        {
            #region color assignment
            string color1 = "";
            switch (colorindex)
            {
                case 1:
                    color1 = "ff1400fa";
                    break;
                case 2:
                    color1 = "ff14f0ff";
                    break;
                case 3:
                    color1 = "ff00ff00";
                    break;
                case 4:
                    color1 = "ffffff00";
                    break;
                case 5:
                    color1 = "ffffaa00";
                    break;
                case 6:
                    color1 = "ffff00ff";
                    break;
                case 7:
                    break;
                case 8:
                    color1 = "ff808080";
                    break;
                case 9:
                    color1 = "ffc0c0c0";
                    break;
                case 10:
                    color1 = "ff0000ff";
                    break;
                case 11:
                    color1 = "ff7f7fff";
                    break;
                case 12:
                    color1 = "ff0000cc";
                    break;
                case 13:
                    color1 = "ff6666cc";
                    break;
                case 14:
                    color1 = "ff000099";
                    break;
                case 15:
                    color1 = "ff4c4c99";
                    break;
                case 16:
                    color1 = "ff00007f";
                    break;
                case 17:
                    color1 = "ff3f3f7f";
                    break;
                case 18:
                    color1 = "ff00004c";
                    break;
                case 19:
                    color1 = "ff08264c";
                    break;
                case 20:
                    color1 = "ff003fff";
                    break;
                case 21:
                    color1 = "ff7f9fff";
                    break;
                case 22:
                    color1 = "ff0033cc";
                    break;
                case 23:
                    color1 = "ff667fcc";
                    break;
                case 24:
                    color1 = "ff002699";
                    break;
                case 25:
                    color1 = "ff4c5f99";
                    break;
                case 26:
                    color1 = "ff001f7f";
                    break;
                case 27:
                    color1 = "ff3f4c7f";
                    break;
                case 28:
                    color1 = "ff00134c";
                    break;
                case 29:
                    color1 = "ff26f4c";
                    break;
                case 30:
                    color1 = "ff007fff";
                    break;
                case 31:
                    color1 = "ff7fbfff";
                    break;
                case 32:
                    color1 = "ff0066cc";
                    break;
                case 33:
                    color1 = "ff7fbfff";
                    break;
                case 34:
                    color1 = "ff004c99";
                    break;
                case 35:
                    color1 = "ff4c7299";
                    break;
                case 36:
                    color1 = "ff003f7f";
                    break;
                case 37:
                    color1 = "ff3f5f7f";
                    break;
                case 38:
                    color1 = "ff00264c";
                    break;
                case 39:
                    color1 = "ff26394c";
                    break;
                case 40:
                    color1 = "ff00bfff";
                    break;
                case 41:
                    color1 = "ff7fdfff";
                    break;
                case 42:
                    color1 = "ff0099cc";
                    break;
                case 43:
                    color1 = "ff66b2cc";
                    break;
                case 44:
                    color1 = "ff007299";
                    break;
                case 45:
                    color1 = "ff4c8599";
                    break;
                case 46:
                    color1 = "ff005f7f";
                    break;
                case 47:
                    color1 = "ff246f7f";
                    break;
                case 48:
                    color1 = "ff0394c";
                    break;
                case 49:
                    color1 = "ff26424c";
                    break;
                case 50:
                    color1 = "ff00ffff";
                    break;
                case 51:
                    color1 = "ff7fffff";
                    break;
                case 52:
                    color1 = "ff00cccc";
                    break;
                case 53:
                    color1 = "ff66cccc";
                    break;
                case 54:
                    color1 = "ff009999";
                    break;
                case 55:
                    color1 = "ff4c9999";
                    break;
                case 56:
                    color1 = "ff007f7f";
                    break;
                case 57:
                    color1 = "ff3f7f7f";
                    break;
                case 58:
                    color1 = "ff004c4c";
                    break;
                case 59:
                    color1 = "ff264c4c";
                    break;
                case 60:
                    color1 = "ff00ffbf";
                    break;
                case 61:
                    color1 = "ff7fffdf";
                    break;
                case 62:
                    color1 = "ff00cc99";
                    break;
                case 63:
                    color1 = "ff6ccb2";
                    break;
                case 64:
                    color1 = "ff009972";
                    break;
                case 65:
                    color1 = "ff4c9985";
                    break;
                case 66:
                    color1 = "ff007f5f";
                    break;
                case 67:
                    color1 = "ff3f7f6f";
                    break;
                case 68:
                    color1 = "ff004c39";
                    break;
                case 69:
                    color1 = "ff264v42";
                    break;
                case 70:
                    color1 = "ff00ff7f";
                    break;
                case 71:
                    color1 = "ff7fffbf";
                    break;
                case 72:
                    color1 = "ff00cc66";
                    break;
                case 73:
                    color1 = "ff66cc99";
                    break;
                case 74:
                    color1 = "ff00994c";
                    break;
                case 75:
                    color1 = "ff4c9972";
                    break;
                case 76:
                    color1 = "ff007f3f";
                    break;
                case 77:
                    color1 = "ff3f7f5f";
                    break;
                case 78:
                    color1 = "ff004c26";
                    break;
                case 79:
                    color1 = "ff264c39";
                    break;
                case 80:
                    color1 = "ff00ff3f";
                    break;
                case 81:
                    color1 = "ff7fff9f";
                    break;
                case 82:
                    color1 = "ff00cc33";
                    break;
                case 83:
                    color1 = "ff66cc7f";
                    break;
                case 84:
                    color1 = "ff009926";
                    break;
                case 85:
                    color1 = "ff4c995f";
                    break;
                case 86:
                    color1 = "ff007f1f";
                    break;
                case 87:
                    color1 = "ff3f7f4c";
                    break;
                case 88:
                    color1 = "ff00fc13";
                    break;
                case 89:
                    color1 = "ff264c2f";
                    break;
                case 90:
                    color1 = "ff00ff00";
                    break;
                case 91:
                    color1 = "ff7fff7f";
                    break;
                case 92:
                    color1 = "ff00cc00";
                    break;
                case 93:
                    color1 = "ff66cc66";
                    break;
                case 94:
                    color1 = "ff009900";
                    break;
                case 95:
                    color1 = "ff4c994c";
                    break;
                case 96:
                    color1 = "ff007f00";
                    break;
                case 97:
                    color1 = "ff3f7f3f";
                    break;
                case 98:
                    color1 = "ff004c00";
                    break;
                case 99:
                    color1 = "ff264c26";
                    break;
                case 100:
                    color1 = "ff3fff00";
                    break;
                case 101:
                    color1 = "ff9fff00";
                    break;
                case 102:
                    color1 = "ff33cc00";
                    break;
                case 103:
                    color1 = "ff7fcc66";
                    break;
                case 104:
                    color1 = "ff269900";
                    break;
                case 105:
                    color1 = "ff5f994c";
                    break;
                case 106:
                    color1 = "ff1f7f00";
                    break;
                case 107:
                    color1 = "ff4f7f3f";
                    break;
                case 108:
                    color1 = "ff134c00";
                    break;
                case 109:
                    color1 = "ff2f4c26";
                    break;
                case 110:
                    color1 = "ff7fff00";
                    break;
                case 111:
                    color1 = "ffcfff7f";
                    break;
                case 112:
                    color1 = "ff66cc00";
                    break;
                case 113:
                    color1 = "ff99cc66";
                    break;
                case 114:
                    color1 = "ff4c9900";
                    break;
                case 115:
                    color1 = "ff72994c";
                    break;
                case 116:
                    color1 = "ff3f7f00";
                    break;
                case 117:
                    color1 = "ff5f7f3f";
                    break;
                case 118:
                    color1 = "ff264c00";
                    break;
                case 119:
                    color1 = "ff394c26";
                    break;
                case 120:
                    color1 = "ffbfff00";
                    break;
                case 121:
                    color1 = "ffdfff7f";
                    break;
                case 122:
                    color1 = "ff99cc00";
                    break;
                case 123:
                    color1 = "ffb2cc66";
                    break;
                case 124:
                    color1 = "ff729900";
                    break;
                case 125:
                    color1 = "ff85994c";
                    break;
                case 126:
                    color1 = "ff5f7f00";
                    break;
                case 127:
                    color1 = "ff6f7f3f";
                    break;
                case 128:
                    color1 = "ff394c00";
                    break;
                case 129:
                    color1 = "ff24c26";
                    break;
                case 130:
                    color1 = "ffffff00";
                    break;
                case 131:
                    color1 = "ffffff7f";
                    break;
                case 132:
                    color1 = "ffcccc00";
                    break;
                case 133:
                    color1 = "ffcccc66";
                    break;
                case 134:
                    color1 = "ff999900";
                    break;
                case 135:
                    color1 = "ff9994c";
                    break;
                case 136:
                    color1 = "ff7f7f00";
                    break;
                case 137:
                    color1 = "ff7f7f3f";
                    break;
                case 138:
                    color1 = "ff4c4c00";
                    break;
                case 139:
                    color1 = "ff4c4c26";
                    break;
                case 140:
                    color1 = "ffffbf00";
                    break;
                case 141:
                    color1 = "ffffdf7f";
                    break;
                case 142:
                    color1 = "ffcc9900";
                    break;
                case 143:
                    color1 = "ffccvv266";
                    break;
                case 144:
                    color1 = "ff997200";
                    break;
                case 145:
                    color1 = "ff9854c";
                    break;
                case 146:
                    color1 = "ff7f5f00";
                    break;
                case 147:
                    color1 = "ff7f6f3f";
                    break;
                case 148:
                    color1 = "ff4c3900";
                    break;
                case 149:
                    color1 = "ff4c4226";
                    break;
                case 150:
                    color1 = "ffff7f00";
                    break;
                case 151:
                    color1 = "ffffbf7f";
                    break;
                case 152:
                    color1 = "ffcc6600";
                    break;
                case 153:
                    color1 = "ffcc9966";
                    break;
                case 154:
                    color1 = "ff994c00";
                    break;
                case 155:
                    color1 = "ff99724c";
                    break;
                case 156:
                    color1 = "ff7f3f00";
                    break;
                case 157:
                    color1 = "ff7f5f3f";
                    break;
                case 158:
                    color1 = "ff4c2600";
                    break;
                case 159:
                    color1 = "ff4c3926";
                    break;
                case 160:
                    color1 = "ffff3f00";
                    break;
                case 161:
                    color1 = "ffff9f7f";
                    break;
                case 162:
                    color1 = "ffcc3300";
                    break;
                case 163:
                    color1 = "ffcc7f66";
                    break;
                case 164:
                    color1 = "ff992600";
                    break;
                case 165:
                    color1 = "ff995f4c";
                    break;
                case 166:
                    color1 = "ff7f1f00";
                    break;
                case 167:
                    color1 = "ff7f4f3f";
                    break;
                case 168:
                    color1 = "ff4c1300";
                    break;
                case 169:
                    color1 = "FF4C2F26";
                    break;
                case 170:
                    color1 = "ffff0000";
                    break;
                case 171:
                    color1 = "ffff7f7f";
                    break;
                case 172:
                    color1 = "ffcc0000";
                    break;
                case 173:
                    color1 = "ffcc6666";
                    break;
                case 174:
                    color1 = "ff990000";
                    break;
                case 175:
                    color1 = "FF994C4C";
                    break;
                case 176:
                    color1 = "ff7f0000";
                    break;
                case 177:
                    color1 = "ff7f3f3f";
                    break;
                case 178:
                    color1 = "ff4c0000";
                    break;
                case 179:
                    color1 = "ff4c2626";
                    break;
                case 180:
                    color1 = "ffff003f";
                    break;
                case 181:
                    color1 = "ffff7f9f";
                    break;
                case 182:
                    color1 = "ffcc0033";
                    break;
                case 183:
                    color1 = "ffcc667f";
                    break;
                case 184:
                    color1 = "ff990026";
                    break;
                case 185:
                    color1 = "ff994c5f";
                    break;
                case 186:
                    color1 = "ff7f001f";
                    break;
                case 187:
                    color1 = "ff7f3f4f";
                    break;
                case 188:
                    color1 = "ff13004c";
                    break;
                case 189:
                    color1 = "ff2c262f";
                    break;
                case 190:
                    color1 = "ffff07f";
                    break;
                case 191:
                    color1 = "ffff7fbf";
                    break;
                case 192:
                    color1 = "ffcc0066";
                    break;
                case 193:
                    color1 = "ffcc6699";
                    break;
                case 194:
                    color1 = "ff99004c";
                    break;
                case 195:
                    color1 = "ff994c72";
                    break;
                case 196:
                    color1 = "ff4c003f";
                    break;
                case 197:
                    color1 = "ff7f3f5f";
                    break;
                case 198:
                    color1 = "ff4c0026";
                    break;
                case 199:
                    color1 = "ff4c2639";
                    break;
                case 200:
                    color1 = "ffff00bf";
                    break;
                case 201:
                    color1 = "ffff7fdf";
                    break;
                case 202:
                    color1 = "ffcc0099";
                    break;
                case 203:
                    color1 = "ffcc66b2";
                    break;
                case 204:
                    color1 = "ff990072";
                    break;
                case 205:
                    color1 = "ff994c85";
                    break;
                case 206:
                    color1 = "ff7f005f";
                    break;
                case 207:
                    color1 = "ff7f005f";
                    break;
                case 208:
                    color1 = "ff4c0039";
                    break;
                case 209:
                    color1 = "ff4c2642";
                    break;
                case 210:
                    color1 = "ffff00ff";
                    break;
                case 211:
                    color1 = "ffff7fff";
                    break;
                case 212:
                    color1 = "ffcc00cc";
                    break;
                case 213:
                    color1 = "ffcc66cc";
                    break;
                case 214:
                    color1 = "ff990099";
                    break;
                case 215:
                    color1 = "ff994c99";
                    break;
                case 216:
                    color1 = "ff7f007f";
                    break;
                case 217:
                    color1 = "ff994c7f";
                    break;
                case 218:
                    color1 = "ff4c004c";
                    break;
                case 219:
                    color1 = "ff4c264c";
                    break;
                case 220:
                    color1 = "ffbf00ff";
                    break;
                case 221:
                    color1 = "ffdf7fff";
                    break;
                case 222:
                    color1 = "ff9900cc";
                    break;
                case 223:
                    color1 = "ffc266cc";
                    break;
                case 224:
                    color1 = "ff720099";
                    break;
                case 225:
                    color1 = "ff854c99";
                    break;
                case 226:
                    color1 = "ff5f007f";
                    break;
                case 227:
                    color1 = "ff6f3f7f";
                    break;
                case 228:
                    color1 = "ff39004c";
                    break;
                case 229:
                    color1 = "ff42264c";
                    break;
                case 230:
                    color1 = "ff7f00ff";
                    break;
                case 231:
                    color1 = "ffbf7fff";
                    break;
                case 232:
                    color1 = "ff6600cc";
                    break;
                case 233:
                    color1 = "ff9966cc";
                    break;
                case 234:
                    color1 = "ff4c0099";
                    break;
                case 235:
                    color1 = "ff724c99";
                    break;
                case 236:
                    color1 = "ff3f007f";
                    break;
                case 237:
                    color1 = "ff5f3f7f";
                    break;
                case 238:
                    color1 = "ff260056";
                    break;
                case 239:
                    color1 = "ff39264c";
                    break;
                case 240:
                    color1 = "ff3f00ff";
                    break;
                case 241:
                    color1 = "ff9f7fff";
                    break;
                case 242:
                    color1 = "ff3300cc";
                    break;
                case 243:
                    color1 = "ff7f66cc";
                    break;
                case 244:
                    color1 = "ff260099";
                    break;
                case 245:
                    color1 = "ff5f4c99";
                    break;
                case 246:
                    color1 = "ff1f007f";
                    break;
                case 247:
                    color1 = "ff4f3f7f";
                    break;
                case 248:
                    color1 = "ff13004c";
                    break;
                case 249:
                    color1 = "ff2f264c";
                    break;
                case 250:
                    color1 = "ff333333";
                    break;
                case 251:
                    color1 = "ff5b5b5b";
                    break;
                case 252:
                    color1 = "ff848484";
                    break;
                case 253:
                    color1 = "ffadadad";
                    break;
                case 254:
                    color1 = "ffd6d6d6";
                    break;
                case 255:
                    break;
                default:
                    break;
            }

            #endregion

            string q = "\u0022";

            string line1 = "	<Style id=" + q + nume1 + q + ">";
            string line2 = "		<LineStyle>";
            string line3 = "			<color>" + color1 + "</color>";
            string line4 = "			<width>" + width1.ToString() + "</width>";
            string line5 = "		</LineStyle>";
            string line6 = "		<PolyStyle>";
            string line7 = "			<color>" + color1 + "</color>";
            string line8 = "			<outline>0</outline>";
            string line9 = "		</PolyStyle>";
            string line91 = "		<IconStyle>";
            string line92 = "			<color>" + "ffe7ebef" + " </color>";
            string line93 = "			<colorMode>normal</colorMode>";
            string line94 = "		</IconStyle>";
            string line10 = "	</Style>";

            string line11 = "	<Style id=" + q + nume1 + "0" + q + ">";
            string line12 = "		<LineStyle>";
            string line13 = "			<color>" + color1 + "</color>";
            string line14 = "			<width>" + (2 * width1).ToString() + "</width>";
            string line15 = "		</LineStyle>";
            string line16 = "		<PolyStyle>";
            string line17 = "			<color>" + color1 + "</color>";
            string line18 = "			<outline>0</outline>";
            string line19 = "		</PolyStyle>";
            string line191 = "		<IconStyle>";
            string line192 = "			<color>" + "ffe7ebef" + " </color>";
            string line193 = "			<colorMode>normal</colorMode>";
            string line194 = "		</IconStyle>";
            string line20 = "	</Style>";

            string line21 = "	<StyleMap id=" + q + "site_" + nume1 + q + ">";
            string line22 = "		<Pair>";
            string line23 = "			<key>normal</key>";
            string line24 = "			<styleUrl>#" + nume1 + "</styleUrl>";
            string line25 = "		</Pair>";
            string line26 = "		<Pair>";
            string line27 = "			<key>highlight</key>";
            string line28 = "			<styleUrl>#" + nume1 + "0</styleUrl>";
            string line29 = "		</Pair>";
            string line30 = "	</StyleMap>";

            if (color1 == "")
            {
                line3 = "";
                line7 = "";
                line13 = "";
                line17 = "";
            }

            return line1 + "\r\n" + line2 + "\r\n" + line3 + "\r\n" + line4 + "\r\n" + line5 + "\r\n" + line6 + "\r\n" +
                    line7 + "\r\n" + line8 + "\r\n" + line9 + "\r\n" +
                    line91 + "\r\n" + line92 + "\r\n" + line93 + "\r\n" + line94 + "\r\n" + 
                    line10 + "\r\n" + line11 + "\r\n" +
                    line12 + "\r\n" + line13 + "\r\n" + line14 + "\r\n" + line15 + "\r\n" + line16 + "\r\n" +
                    line17 + "\r\n" + line18 + "\r\n" + line19 + "\r\n" +
                    line191 + "\r\n" + line192 + "\r\n" + line193 + "\r\n" + line194 + "\r\n" + 
                    line20 + "\r\n" + line21 + "\r\n" + line22 + "\r\n" + 
                    line23 + "\r\n" + line24 + "\r\n" + line25 + "\r\n" + line26 + "\r\n" +
                    line27 + "\r\n" + line28 + "\r\n" + line29 + "\r\n" + line30 + "\r\n";
        }

        public static string get_kml_line_placemark(string placemark_name, string style_map_name, string coords)
        {

            string line1 = "		<Placemark>";
            string line2 = "			<name>" + placemark_name + "</name>";
            string line3 = "			<styleUrl>#" + style_map_name + "</styleUrl>";
            string line4 = "			<LineString>";
            string line5 = "				<tessellate>1</tessellate>";
            string line6 = "				<coordinates>";
            string line7 = coords;
            string line8 = "				</coordinates>";
            string line9 = "			</LineString>";
            string line10 = "		</Placemark>";

            return line1 + "\r\n" + line2 + "\r\n" + line3 + "\r\n" + line4 + "\r\n" + line5 + "\r\n" +
                    line6 + "\r\n" + line7 + "\r\n" + line8 + "\r\n" + line9 + "\r\n" + line10 + "\r\n";
        }


        public static string get_kml_polygon_placemark(string placemark_name, string style_map_name, string coords)
        {

            string line1 = "		<Placemark>";
            string line2 = "			<name>" + placemark_name + "</name>";
            string line3 = "			<styleUrl>#" + style_map_name + "</styleUrl>";
            string line4 = "			<Polygon>";
            string line5 = "				<altitudeMode>clampToGround</altitudeMode>";
            string line6 = "				<outerBoundaryIs>";
            string line7 = "				 <LinearRing>";
            string line8 = "				 <coordinates>";
            string line9 = "  " + coords;
            string line10 = "				 </coordinates>";
            string line11 = "				 </LinearRing>";
            string line12 = "				</outerBoundaryIs>";
            string line13 = "			</Polygon>";
            string line14 = "		</Placemark>";

            return line1 + "\r\n" + line2 + "\r\n" + line3 + "\r\n" + line4 + "\r\n" + line5 + "\r\n" +
                    line6 + "\r\n" + line7 + "\r\n" + line8 + "\r\n" + line9 + "\r\n" + line10 + "\r\n" +
                    line11 + "\r\n" + line12 + "\r\n" + line13 + "\r\n" + line14 + "\r\n";
        }

        public static List<string> get_layers_from_dwg()
        {
            List<string> lista_layers = new List<string>();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));
                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false && Name_of_layer.Contains("$") == false && Layer1.IsFrozen == false && Layer1.IsOff == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;
                    }
                }
                System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    lista_layers.Add(dt2.Rows[i][0].ToString());
                }
                Trans1.Commit();
            }
            return lista_layers;
        }

    }
}
