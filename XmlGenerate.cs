using InPlanLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace SISOLVERTEST
{
    /// <summary>
    /// 自动生成子元数 
    /// 自动生成子元数的属性
    /// </summary>
   public class XmlGenerate
    {

       public  struct MyStruct
        {
            string modela;
        }

        public enum ImdependenceModel
        {
            broadside_coupled_stripline_2s = 15,
            coated_coplanar_strips_with_lower_gnd_1b = 13,
            coated_coplanar_waveguide_with_lower_gnd_1b = 13,
            coated_coplanar_waveguide_with_lower_gnd_2b = 13,
            coated_microstrip_1b = 3,
            coated_microstrip_2b = 3,
            diff_coated_coplanar_strips_with_lower_gnd_1b = 14,
            diff_coated_coplanar_waveguide_with_lower_gnd_1b = 14,
            diff_offset_coplanar_strips_1b1a = 12,
            diff_offset_coplanar_waveguide_1b1a = 12,
            diff_surface_coplanar_strips_1b = 12,
            diff_surface_coplanar_strips_with_lower_gnd_1b = 10,
            diff_surface_coplanar_waveguide_with_lower_gnd_1b = 10,
            edge_coupled_coated_microstrip_1b = 4,
            edge_coupled_coated_microstrip_2b = 4,
            edge_coupled_coated_without_ground_1b = 14,
            edge_coupled_coated_without_ground_2b = 14,
            edge_coupled_embedded_microstrip_1b1a = 8,
            edge_coupled_embedded_microstrip_1b2a = 8,
            edge_coupled_embedded_microstrip_1e1b1a = 8,
            edge_coupled_embedded_microstrip_1e1b2a = 8,
            edge_coupled_embedded_microstrip_2b1a = 8,
            edge_coupled_embedded_microstrip_2b2a = 8,
            edge_coupled_offset_stripline_1b1a = 2,
            edge_coupled_offset_stripline_1b1a1r = 2,
            edge_coupled_offset_stripline_1b2a = 2,
            edge_coupled_offset_stripline_2b1a = 2,
            edge_coupled_offset_stripline_2b1a1r = 2,
            edge_coupled_offset_stripline_2b2a = 2,
            edge_coupled_surface_microstrip_1b = 6,
            embedded_microstrip_1b1a = 7,
            embedded_microstrip_1b2a = 7,
            embedded_microstrip_1e1b1a = 7,
            embedded_microstrip_1e1b2a = 7,
            embedded_microstrip_2b1a = 7,
            embedded_microstrip_2b2a = 7,
            offset_coplanar_waveguide_1b1a = 11,
            offset_stripline_1b1a = 1,
            offset_stripline_1b2a = 1,
            offset_stripline_2b1a = 1,
            offset_stripline_2b2a = 1,
            surface_coplanar_strips_with_lower_gnd_1b = 9,
            surface_coplanar_waveguide_with_lower_gnd_1b = 9,
            surface_microstrip_1b = 5,
        }

        /// <summary>
        /// 自动添加属性
        /// 自动挂载到父节点
        /// </summary>
        /// <param name="parentnode"></param>
        /// <param name="xmlelem"></param>
        public static void AutoGerateAttr(XmlDocument rootNode, XmlElement xmlelem,string[] attrName, string[] values)
        {
            //自动设置属性
            //属性数组：
            //XmlElement nodeDocm = xmldoc.CreateElement(rootNodeName);
            for (int i = 0; i < attrName.Length; i++)
            {
                xmlelem.SetAttribute(attrName[i], values[i]);
            }
            rootNode.AppendChild(xmlelem);
            //return xmlelem;
            
        }

        internal static void AutoGerateCNodeAttr(XmlElement rootNode, XmlElement nodeChild, string[] attrName, string[] values)
        {
            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i], values[0]);
            }
        }

        internal static void AutoGerateCNodeAttr(XmlElement rootNode, XmlElement nodeChild, string[] attrName, string values)
        {
            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i], values);
            }
        }

        internal static void AutoGerateCNodeAttr_CORE(XmlElement rootNode, XmlElement nodeChild, string[] attrName, string[] values)
        {
            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i], values[i]);
            }
            rootNode.AppendChild(nodeChild);
        }

        internal static void AutoGerateCNodeAttr_CORE(XmlElement rootNode, XmlElement nodeChild, string[] attrName, object[] values)
        {
            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i], values[i].ToString());
            }
            rootNode.AppendChild(nodeChild);
        }

        internal static void AutoGerateCNodeAttr_CORE(XmlElement rootNode, XmlElement nodeChild, string[] attrName,List<string[]> list)
        {
            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i],"0");
            }
            rootNode.AppendChild(nodeChild);
        }

        internal static void AutoGerateNodeForCommon(XmlElement rootNode, XmlElement nodeChild,string[] attrName,string[] values)
        {

            for (int i = 0; i < attrName.Length; i++)
            {
                nodeChild.SetAttribute(attrName[i], values[0]);
            }
            rootNode.AppendChild(nodeChild);
        }

        internal static void DictNotContainKeys(Dictionary<string, string>[] dictattrarray, Array array)
        {
            //判断DIC 中是否有数组中的 key
            //有 不处理
            //没有 添加并设置值为空
            foreach (var dictitem in dictattrarray)
            {
                foreach (var arritem in array)
                {
                    if (!dictitem.ContainsKey(arritem.ToString()))
                    {
                        dictitem.Add(arritem.ToString(), "Inplan中无对应字段");
                    }
                }
            }
        }

        public static void InterfaceOfSiSolver(string filepath)
        {
            //添加SPF_PROXY   的COM组件

            //todo: 这里要打开
            
            
            int result = 0;
            var svc = new SPF_PROXYLib.PROXY(); // 新建COM对象
            string xml = filepath;    //文件的绝对路径
            result = svc.FUN(xml, 1, 1); // 参数1：XML文件路径；参数2：是否生成EXCEL报告，1=生成，0=不生成；参数3：是否生成PDF报告，1=生成，0=不生成
            if (result == 0)
            {
                MessageBox.Show("阻抗报告已导出"); //运行完成,没有错误
                openSisolverReport(filepath); //系统自动打开生成的阻抗计算报告PDF和EXCEL  
            }
            else
            {
                MessageBox.Show("运行完成,发现错误,错误内容请查看Log文件");
            }
        }

        private static void openSisolverReport(string filepath)
        {
            //自动打开已生成文件
            string path = filepath;
            string xml = ".xml";
            string pdf = ".pdf";
            string xls = ".xls";

            //filepath.Replace(".xml", xml);
            //filepath.Replace(".xml", pdf);
            //filepath.Replace(".xml", xls);

            string xmlpath = filepath.Replace(".xml", xml);
            string pdfpath = filepath.Replace(".xml", pdf); ;
            string xlspath = filepath.Replace(".xml", xls); ;
            //System.Diagnostics.Process.Start(xmlpath); //打开此文件。
            System.Diagnostics.Process.Start(pdfpath); //打开此文件。
            System.Diagnostics.Process.Start(xlspath); //打开此文件。
        }


        /// <summary>
        /// 给子节点组中的子节点属性批量赋值
        /// </summary>
        /// <param name="xmldoc"></param>
        /// <param name="parentElemt"></param>
        /// <param name="xmlElements"></param>
        /// <param name="attrs"></param>
        /// <param name="list"></param>
        /// <param name="attrname"></param>
        internal static void SetAttrsByArrayToJobAttrs2(XmlDocument xmldoc, XmlElement parentElemt, XmlElement[] xmlElements, string[] attrs, List<Dictionary<string, string>> list, string attrname)
        {
            foreach (var xmlelement in xmlElements)
            {
                XmlElement childElem = xmldoc.CreateElement(attrname);
                int i = 0;
                foreach (var listelement in list)
                {
                    XmlGenerate.AutoGerateCNodeAttr_CORE(parentElemt, childElem, attrs, listelement.Values.ToArray());
                    i++;
                }
                string count = parentElemt.OuterXml;
            }
            //循环给每一个节点赋值，并挂载到父节点

            //加载父节点;{rootNode.Name}  设置子节点 {nodeChild.Name}  属性：{attrName[i]}   属性值： {value[i]}
        }

        /// <summary>
        /// 给子节点组中的子节点属性批量赋值
        /// </summary>
        /// <param name="xmldoc"></param>
        /// <param name="parentElemt"></param>
        /// <param name="xmlElements"></param>
        /// <param name="attrs"></param>
        /// <param name="list"></param>
        /// <param name="attrname"></param>
        internal static void SetAttrsByArrayToJobAttrs3(XmlDocument xmldoc, XmlElement parentElemt, XmlElement[] xmlElements, List<Dictionary<string, string>> list, string attrname)
        {
            //int i = 0;
            //最原始的想法  再引入数组批量赋值 xmlelement.SetAttribute(list[i].Keys,list[i].Values);
            List<string[]> keyslist = new List<string[]>();
            List<string[]> valueslist = new List<string[]>();
            try
            {
                foreach (var li in list)
                {
                    int j = 0;
                    int k = 0;
                    string[] strkey = new string[li.Keys.Count];
                    string[] strvalue = new string[li.Keys.Count];

                    foreach (var key in li.Keys)
                    {
                        strkey[j] = key.ToString();
                        j++;
                    }
                    keyslist.Add(strkey);
                    foreach (var value in li.Values)
                    {
                        strvalue[k] = value.ToString();
                        k++;
                    }
                    valueslist.Add(strvalue);

                }

            }
            catch (Exception ex)
            {
                string exstr = ex.ToString();
                //throw;
            }



            int i = 0;
            foreach (var xmlelement in xmlElements)
            {
                
                XmlElement childElem = xmldoc.CreateElement(attrname);
                XmlGenerate.AutoGerateCNodeAttr_CORE(parentElemt, childElem, keyslist[i].ToArray(), valueslist[i].ToArray());
                i++;
            }



            //foreach (var xmlelement in xmlElements)
            //{
            //    int q = 0;
            //     //xmlelement = xmldoc.CreateElement(attrname);
            //    for (int m = 0; m < 12; m++)
            //    {
            //        xmlelement.SetAttribute(list[q].[m], strvalue[m]);
            //    }
            //    parentElemt.AppendChild(xmlelement);
               
            //}
            //循环给每一个节点赋值，并挂载到父节点
            string count = parentElemt.OuterXml;
            //加载父节点;{rootNode.Name}  设置子节点 {nodeChild.Name}  属性：{attrName[i]}   属性值： {value[i]}
        }


        internal static void SetAttrsByArrayToJobAttrs(XmlDocument xmldoc, XmlElement parentElem, XmlElement[] xmlElements, string[] attrs, List<string[]> list, string attrname)
        {

            //foreach (var xmlelement in xmlElements)
            //{

            //循环给每一个节点赋值，并挂载到父节点
            foreach (var listElem in list)
            {
                int i = 0;
                XmlElement childElem = xmldoc.CreateElement(attrname);
                XmlGenerate.AutoGerateCNodeAttr_CORE(parentElem, childElem, attrs, listElem.ToArray());
                i++;

            }
            //}
            //循环给每一个节点赋值，并挂载到父节点
            //int i = 0;
            //foreach (var item in list)
            //{
            //    XmlElement element = xmldoc.CreateElement(attrname);
            //    XmlGenerate.AutoGerateCNodeAttr_CORE(element, parentElem, attrs, item.ToArray());
            //    i++;
            //}
        }

        private void SetAttrsByArraylistToJobAttrs(XmlDocument xmldoc, XmlElement parentElem, string[] attrs, ArrayList[] arraylist, string attrname)
        {
            //循环给每一个节点赋值，并挂载到父节点
            int i = 0;
            foreach (var item in arraylist)
            {
                XmlElement element = xmldoc.CreateElement(attrname);
                object[] obj = item.ToArray();
                XmlGenerate.AutoGerateCNodeAttr_CORE(element, parentElem, attrs, obj);
                i++;
            }
        }

        public  static string fileExists(string v, string p1, int times,string jobName)
        {
            string newfilename = v;
            if (File.Exists(v))
            {
                string copy = "_副本";
                string filename = string.Format("{0}{1}{2}{3}.xml", p1, jobName, copy, times);
                newfilename = filename;
                times++;
                newfilename = fileExists(filename, p1, times, jobName);
            }
            return newfilename;
        }

        internal static void DictToAddArrayList(Dictionary<string, string>[] dict, ArrayList arraylist)
        {

            int i = 0;
            foreach (var item in dict)
            {
                arraylist.Add(dict[i].Values.ToList<string>());
                i++;
            }
        }

        /// <summary>
        /// 获取属性和值 循环获取元素属性和值
        /// </summary>
        internal static Dictionary<string, string> AllAttrGetKeyValues(dynamic obj,Dictionary<string ,string> dictNew)
        {
            foreach (IAttribute attr in obj.AttrsEx1(true, true))
            {
                switch (attr.ExtentionContentType())
                {
                    case ExtContentType.EXT_TYPE_INT:
                        dictNew.Add(attr.Name(),attr.IntVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_DOUBLE:
                        dictNew.Add(attr.Name(), attr.DoubleVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_BOOL:
                        dictNew.Add(attr.Name(), attr.BoolVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_STRING:

                        dictNew.Add(attr.Name(), attr.StrVal());
                        break;
                    /*
                     # 1、材料Faimily 是MIX
                     # 2、料号是钻孔
                    */


                    //if (attr.Name() == "FAMILY")
                    //{
                    //    string str = "";
                    //}
                    //else if (attr.Name()== "FAMILY"&& attr.StrVal()== "Mixed")
                    //{
                    //    dictNew.Add(attr.Name(), "NY2150");
                    //}
                    //else
                    //{
                    //    dictNew.Add(attr.Name(), attr.StrVal());
                    //}

                    case ExtContentType.EXT_TYPE_DATE:
                        dictNew.Add(attr.Name(), attr.Date().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_ENUM:
         
                        //if (attr.Name() == "FAMILY" && attr.StrVal() == "Mixed")
                        //{
                        //    dictNew.Add(attr.Name(), "NY2150");
                        //}
                        //else
                        //{
                        //    dictNew.Add(attr.Name(), attr.StrVal());
                        //}

                        dictNew.Add(attr.Name(), attr.StrVal());
                        break;
                    case ExtContentType.EXT_TYPE_LENGTH:
                        dictNew.Add(attr.Name(), attr.Length(AvailableUnits.MIL).ToString());
                        break;
                    case ExtContentType.EXT_TYPE_AREA:
                        //dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.MIL).ToString());
                        break;
                    case ExtContentType.EXT_TYPE_WEIGHT:
                        dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.OZ).ToString());
                        break;
                    case ExtContentType.EXT_TYPE_CURRENCY:
                        //
                        break;
                    case ExtContentType.EXT_TYPE_TEMPERATURE:
                        dictNew.Add(attr.Name(), attr.Temperature(AvailableUnits.CELCIUS).ToString());
                        break;
                    case ExtContentType.EXT_TYPE_BLOB:
                        //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_BLOB_REF:
                        //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_RESISTANCE:

                        break;
                    case ExtContentType.EXT_TYPE_FREQUENCY:

                        break;
                    case ExtContentType.EXT_TYPE_CLOB:

                        break;
                    case ExtContentType.EXT_TYPE_CURRENT:

                        break;
                    case ExtContentType.EXT_TYPE_VOLTAGE:

                        break;
                    case ExtContentType.LAST_EXT_CONT_TYPE:

                        break;
                    default:
                        break;
                }

            }
            return dictNew;
        }

        internal static Dictionary<string, double> AllAttrGetKeyValues_double(dynamic obj, Dictionary<string, double> dictNew)
        {
            foreach (IAttribute attr in obj.AttrsEx1(true, true))
            {
                switch (attr.ExtentionContentType())
                {
                    case ExtContentType.EXT_TYPE_INT:
                        dictNew.Add(attr.Name(), attr.IntVal());
                        break;
                    case ExtContentType.EXT_TYPE_DOUBLE:
                        dictNew.Add(attr.Name(), attr.DoubleVal());
                        break;
                    case ExtContentType.EXT_TYPE_BOOL:
                        break;
                    case ExtContentType.EXT_TYPE_STRING:
                        break;
                    case ExtContentType.EXT_TYPE_DATE:
                        break;
                    case ExtContentType.EXT_TYPE_ENUM:
                        break;
                    case ExtContentType.EXT_TYPE_LENGTH:
                        dictNew.Add(attr.Name(), attr.Length(AvailableUnits.MIL));
                        break;
                    case ExtContentType.EXT_TYPE_AREA:
                        //dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.MIL).ToString());
                        break;
                    case ExtContentType.EXT_TYPE_WEIGHT:
                        dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.OZ));
                        break;
                    case ExtContentType.EXT_TYPE_CURRENCY:
                        //
                        break;
                    case ExtContentType.EXT_TYPE_TEMPERATURE:
                        dictNew.Add(attr.Name(), attr.Temperature(AvailableUnits.CELCIUS));
                        break;
                    case ExtContentType.EXT_TYPE_BLOB:
                        //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_BLOB_REF:
                        //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                        break;
                    case ExtContentType.EXT_TYPE_RESISTANCE:

                        break;
                    case ExtContentType.EXT_TYPE_FREQUENCY:

                        break;
                    case ExtContentType.EXT_TYPE_CLOB:

                        break;
                    case ExtContentType.EXT_TYPE_CURRENT:

                        break;
                    case ExtContentType.EXT_TYPE_VOLTAGE:

                        break;
                    case ExtContentType.LAST_EXT_CONT_TYPE:

                        break;
                    default:
                        break;
                }

            }
            return dictNew;
        }

        internal static void DictClone(string[] attrs, string[] attrs_Inplan, Dictionary<string, string> dictionary, Dictionary<string, string> dictionaryclone)
        {
            int i = 0;
            foreach (var key in attrs)
            {
                dictionaryclone.Add(attrs[i], dictionary[attrs_Inplan[i]]);
                i++;
            }
        }

        /// <summary>
        /// 创建真实的子集键值对
        /// </summary>
        /// <param name="attrsMaterialsFoil_Inplan"></param>
        /// <param name="dictionary"></param>
        internal static void ChildDictCreate(string[] attrsMaterialsFoil_Inplan, Dictionary<string, string> dictionary, Dictionary<string, string> childdict)
        {
            //childdict = new Dictionary<string, string>();
            int i = 0;
            foreach (var key in attrsMaterialsFoil_Inplan)
            {
                if (dictionary.Keys.Contains(attrsMaterialsFoil_Inplan[i]))
                {
                    childdict.Add(attrsMaterialsFoil_Inplan[i],dictionary[key]);
                }
                i++;
            }
        }

        internal static void FilterKeyValues(Dictionary<string, string> dictionary1, Dictionary<string, string> dictionary2, string[] attrsMaterialsFoil)
        {
            //A:创建真实的子集键值对 
            //B:Sisolver 数组 
            //前提是： 对应关系 一一对应 i 才会一一赋值
            int i = 0;
            foreach (var key in dictionary1.Keys)
            {
                dictionary2.Add(attrsMaterialsFoil[i], dictionary1[key]);
                i++;
            }
        }

        /// <summary>
        /// 获取指定属性和值 循环获取元素属性和值
        /// </summary>
        internal static Dictionary<string, string> SpecialAttrGetKeyValues(dynamic obj, Dictionary<string, string> dictNew, Array arr)
        {
            foreach (IAttribute attr in obj.AttrsEx1(true, true))
            {
                //如果Job属性中包含数组变量 就赋值元数属性值
                if (((IList)arr).Contains(attr.Name()))
                {
                    switch (attr.ExtentionContentType())
                    {
                        case ExtContentType.EXT_TYPE_INT:
                            dictNew.Add(attr.Name(), attr.IntVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_DOUBLE:
                            dictNew.Add(attr.Name(), attr.DoubleVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_BOOL:
                            dictNew.Add(attr.Name(), attr.BoolVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_STRING:
                            dictNew.Add(attr.Name(), attr.StrVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_DATE:
                            dictNew.Add(attr.Name(), attr.Date().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_ENUM:
                            dictNew.Add(attr.Name(), attr.EnumStrings().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_LENGTH:
                            dictNew.Add(attr.Name(), attr.Length(AvailableUnits.MIL).ToString());
                            break;
                        case ExtContentType.EXT_TYPE_AREA:
                            //dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.MIL).ToString());
                            break;
                        case ExtContentType.EXT_TYPE_WEIGHT:
                            dictNew.Add(attr.Name(), attr.Weight(AvailableUnits.OZ).ToString());
                            break;
                        case ExtContentType.EXT_TYPE_CURRENCY:
                            //
                            break;
                        case ExtContentType.EXT_TYPE_TEMPERATURE:
                            dictNew.Add(attr.Name(), attr.Temperature(AvailableUnits.CELCIUS).ToString());
                            break;
                        case ExtContentType.EXT_TYPE_BLOB:
                            //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_BLOB_REF:
                            //dictNew.Add(attr.Name(), attr.BlobVal().ToString());
                            break;
                        case ExtContentType.EXT_TYPE_RESISTANCE:

                            break;
                        case ExtContentType.EXT_TYPE_FREQUENCY:

                            break;
                        case ExtContentType.EXT_TYPE_CLOB:

                            break;
                        case ExtContentType.EXT_TYPE_CURRENT:

                            break;
                        case ExtContentType.EXT_TYPE_VOLTAGE:

                            break;
                        case ExtContentType.LAST_EXT_CONT_TYPE:

                            break;
                        default:
                            break;
                    }
                }
            }
            return dictNew;
        }
        public static object getAttr(Dictionary<string,dynamic> uda_Dic,string key)
        {
            object a = null;
            IAttribute attr = uda_Dic[key];
            switch (attr.ExtentionContentType())
            {
                case ExtContentType.EXT_TYPE_BLOB:
                    break;
                case ExtContentType.EXT_TYPE_BLOB_REF:
                    break;
                case ExtContentType.EXT_TYPE_BOOL:
                    a = attr.BoolVal();
                    break;
                case ExtContentType.EXT_TYPE_CLOB:
                    break;
                case ExtContentType.EXT_TYPE_CURRENCY:
                    break;
                case ExtContentType.EXT_TYPE_CURRENT:
                    break;
                case ExtContentType.EXT_TYPE_DATE:
                    break;
                case ExtContentType.EXT_TYPE_DOUBLE:
                    a = attr.DoubleVal();
                    break;
                case ExtContentType.EXT_TYPE_FREQUENCY:
                    break;
                case ExtContentType.EXT_TYPE_INT:
                    a = attr.IntVal();
                    break;
                case ExtContentType.EXT_TYPE_RESISTANCE:
                    break;
                case ExtContentType.EXT_TYPE_STRING:
                    a = attr.StrVal();
                    break;
                case ExtContentType.EXT_TYPE_ENUM:
                    a = attr.StrVal();
                    break;
                case ExtContentType.EXT_TYPE_TEMPERATURE:
                    break;
                case ExtContentType.EXT_TYPE_VOLTAGE:
                    break;
                case ExtContentType.LAST_EXT_CONT_TYPE:
                    break;
                default:
                    break;
            }
            return a;
        }

        internal static double UmToMil(double t1)
        {
            return t1 * 0.03937;
        }

        public static object getAttr(Dictionary<string, dynamic> uda_Dic, string key, AvailableUnits unit)
        {
            object a = null;
            IAttribute attr = uda_Dic[key];
            switch (attr.ExtentionContentType())
            {
                case ExtContentType.EXT_TYPE_AREA:
                    a = attr.Area(unit);
                    break;
                case ExtContentType.EXT_TYPE_BLOB:
                    break;
                case ExtContentType.EXT_TYPE_BLOB_REF:
                    break;
                case ExtContentType.EXT_TYPE_BOOL:
                    a = attr.BoolVal();
                    break;
                case ExtContentType.EXT_TYPE_CLOB:
                    break;
                case ExtContentType.EXT_TYPE_CURRENCY:
                    break;
                case ExtContentType.EXT_TYPE_CURRENT:
                    break;
                case ExtContentType.EXT_TYPE_DATE:
                    break;
                case ExtContentType.EXT_TYPE_DOUBLE:
                    a = attr.DoubleVal();
                    break;
                case ExtContentType.EXT_TYPE_FREQUENCY:
                    break;
                case ExtContentType.EXT_TYPE_INT:
                    a = attr.IntVal();
                    break;
                case ExtContentType.EXT_TYPE_LENGTH:
                    a = attr.Length(unit);
                    break;
                case ExtContentType.EXT_TYPE_RESISTANCE:
                    break;
                case ExtContentType.EXT_TYPE_STRING:
                    a = attr.StrVal();
                    break;
                case ExtContentType.EXT_TYPE_TEMPERATURE:
                    break;
                case ExtContentType.EXT_TYPE_VOLTAGE:
                    break;
                case ExtContentType.EXT_TYPE_WEIGHT:
                    a = attr.Weight(unit);
                    break;
                case ExtContentType.LAST_EXT_CONT_TYPE:
                    break;
                default:
                    break;
            }
            return a;
        }

        public static object getAttr(Dictionary<string, dynamic> uda_Dic,string key, ComboBox cmb)
        {
            object a = null;
            IAttribute attr = uda_Dic[key];
            switch (attr.ExtentionContentType())
            {
                case ExtContentType.EXT_TYPE_BLOB:
                    break;
                case ExtContentType.EXT_TYPE_BLOB_REF:
                    break;
                case ExtContentType.EXT_TYPE_BOOL:
                    a = attr.BoolVal();
                    break;
                case ExtContentType.EXT_TYPE_CLOB:
                    break;
                case ExtContentType.EXT_TYPE_CURRENCY:
                    break;
                case ExtContentType.EXT_TYPE_CURRENT:
                    break;
                case ExtContentType.EXT_TYPE_DATE:
                    break;
                case ExtContentType.EXT_TYPE_DOUBLE:
                    a = attr.DoubleVal();
                    break;
                case ExtContentType.EXT_TYPE_ENUM:
                    a = attr.StrVal();
                    Array enum_strings = attr.EnumStrings();
                    foreach (var e in enum_strings)
                    {
                        cmb.Items.Add(e);
                        //cmb.Properties.Items.Add(e);
                    }
                    break;
                case ExtContentType.EXT_TYPE_FREQUENCY:
                    break;
                case ExtContentType.EXT_TYPE_INT:
                    a = attr.IntVal();
                    break;
                case ExtContentType.EXT_TYPE_RESISTANCE:
                    break;
                case ExtContentType.EXT_TYPE_STRING:
                    a = attr.StrVal();
                    break;
                case ExtContentType.EXT_TYPE_TEMPERATURE:
                    break;
                case ExtContentType.EXT_TYPE_VOLTAGE:
                    break;
                case ExtContentType.LAST_EXT_CONT_TYPE:
                    break;
                default:
                    break;
            }
            return a;
        }

        public static int jobcomponslength(Array array)
        {
            int i = 0;
            foreach (var item in array)
            {
                i++;
            }
            return i;
        }

        //internal static ImdependenceModel ImdependenceModelGetEnum(ImdependenceModel imdep)
        //{
        //    int a = 0;
        //    switch (imdep)
        //    {
        //        default:
        //            break;
        //    }
        //    return imdep;
        //}
    }
}
