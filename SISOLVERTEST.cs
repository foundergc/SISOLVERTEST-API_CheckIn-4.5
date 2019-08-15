using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using InPlanLib;

namespace SISOLVERTEST
{
    public partial class SISOLVERTEST : Form
    {
        public static IApplicationManager appManager;
        public static IJob theJob;
        public static IJobManager jobManager;
        public static IStackup stackup;
        public static string jobName;
        int times = 1;
        int logintimes = 1;

        //一级
        string[] attrDocument = { "Dev_Owner", "Fab", "Version", "Project_Name", "Project_Ver", "Project_Designer", "Date", "Layer_Count", "Base_Material", "Customer", "Stackup_Ready", "Imp_Calc_Fixed", "Width_Min" };
        //string[] attrDocument_values = { "FOUNDERPCB F5", "KB", "1.0", "TEST", "A0", "FOUNDERPCB DEMO", "2019/7/15", "6", "IS415","HUAWEI","YES","YES","3"};
        string[] attr_values = { "0" };

        //二级
        string[] attrMaterials = { "MAT_TYPE" };
        string[] attrProcess_Parameters = { "PARA_TYPE" };

        //三级
        /*string[] attrsMaterialsCore = { "TYPE", "MRP_CODE", "VENDOR", "FAMILY", "TOP_FOIL_CU_WEIGHT", "BOT_FOIL_CU_WEIGHT", "LAMINATE_RAW_THICKNESS",
            "LAMINATE_PERMITTIVITY_" ,"LAMINATE_DISSIPATION_FACTOR_", "RESIN_PERCENTAGE_", "TOP_FOIL_RAW_THICKNESS_", "BOT_FOIL_RAW_THICKNESS_",
            "PREPREG_COUNT_", "CORE_STRUCTURE_","GLASS_TYPE_", "FOIL_TYPE_"};*/
        string[] attrsMaterialsCore = { "TYPE", "MRP_CODE", "VENDOR", "FAMILY", "TOP_FOIL_CU_WEIGHT", "BOT_FOIL_CU_WEIGHT", "LAMINATE_RAW_THICKNESS" };
        string[] attrsMaterialsCore_Inplan = { "TYPE", "MRP_NAME", "VENDOR", "FAMILY", "TOP_FOIL_CU_WEIGHT", "BOT_FOIL_CU_WEIGHT", "LAMINATE_THICKNESS" };

        //string[] attrsMaterialsPrepreg = { "TYPE", "MRP_CODE", "VENDOR", "FAMILY", "PP_TYPE", "RESIN_PERCENTAGE", "PP_RAW_THICKNESS", "PP_RAW_DK_" };
        string[] attrsMaterialsPrepreg = { "TYPE", "MRP_CODE", "VENDOR", "FAMILY", "PP_TYPE", "RESIN_PERCENTAGE", "PP_RAW_THICKNESS" };
        string[] attrsMaterialsPrepreg_Inplan = { "TYPE", "MRP_NAME", "VENDOR", "FAMILY", "PP_TYPE_", "RESIN_PERCENTAGE", "RAW_THICKNESS" };

        //string[] attrsMaterialsFoil = { "TYPE", "MRP_CODE", "VENDOR", "CU_WEIGHT", "FOIL_RAW_THICKNESS_", "FOIL_TYPE_" };
        string[] attrsMaterialsFoil = { "TYPE", "MRP_CODE", "VENDOR", "CU_WEIGHT" };
        string[] attrsMaterialsFoil_Inplan = { "TYPE", "MRP_NAME", "VENDOR", "CU_WEIGHT" };

        string[] attrsMaterialsStackupSegms = { "SEG_INDEX", "SEG_TYPE", "MRP_CODE" };

        /*string[] attrsProParaCU_LAYER = { "CU_LAYER_INDEX", "CU_RAW_THICKNESS_", "CU_FINISH_THICKNESS" , "COPPER_USAGE"
                 , "CU_RAW_WEIGHT" , "CU_UNDERCUT", "LAYER_SIGNAL_TYPE_" };*/
        string[] attrsProParaCU_LAYER = { "CU_LAYER_INDEX", "CU_RAW_THICKNESS_", "CU_FINISH_THICKNESS"
                , "COPPER_USAGE"        };

        string[] attrsProParaCU_LAYER_Inplan = { "OVERALL_THICKNESS", "COPPER_TYPE" , "COPPER_USAGE"
                , "REQUIRED_CU_WEIGHT"  };
        /*
         string[] attrsProParaCU_LAYER_Inplan = { "CU_RAW_THICKNESS_", "CU_FINISH_THICKNESS" , "COPPER_USAGE"
                , "CU_RAW_WEIGHT" , "CU_UNDERCUT" };
             */


        string[] attrProParaSM_DIE_LAYER = { "DIE_LAYER_INDEX", "DIE_RAW_THICKNESS", "DIE_FINISH_THICKNESS",
            "RESIN_FILLED_CU_THK", "RESIN_FLOW_THK", "RESIN_FILLED_VIA_THK" };

        //string[] attrsProParaSM_LAYER = { "SM_DK", "SM_C1", "SM_C2", "SM_C3" };
        string[] attrsProParaSM_LAYER = { "SM_DK", "SM_C1", "SM_C2", "SM_C3" };
        //string[] attrsProParaSM_LAYER_Inplan = { "SM_DK", "SM_C1", "SM_C2", "SM_C3" };

        //string[] attrsProParaDRILL_LIST = { "VIA_TYPE", "START_LAYER", "END_LAYER" };
        string[] attrsProParaDRILL_LIST = { "START_LAYER", "END_LAYER" };
        string[] attrsProParaDRILL_LIST_Inplan = { "START_INDEX", "END_INDEX" };
        string[] attrsProParaDRILL_LIST_New = { "VIA_TYPE", "START_INDEX", "END_INDEX" };

        /*string[] attrsImpeListsIMP_ITEM = { "ITEM_INDEX", "IMP_MODEL_NO", "SIGNAL_LAYER", "REF_LAYER_1", "REF_LAYER_2",
            "NUM_SIGNAL", "ZO_TARGET", "ZO_TOL", "WIDTH_DESIGN", "SPACE_DESIGN", "COPLANAR_SPACE_DESIGN", "zo_calc", "zo_optimal",
            "width_optimal", "space_optimal","coplanar_space_optimal" ,"parameter_tags","parameter_values_design","parameter_values_optimal"};*/
        string[] attrsImpeListsIMP_ITEM = { "ITEM_INDEX", "IMP_MODEL_NO", "SIGNAL_LAYER", "REF_LAYER_1" };
        string[] attrsImpeListsIMP_ITEM_Inplan = { "ARTWORK_TRACE_WIDTH", "CALC_REQ_TRACE_WIDTH_TOL_PLUS", "DESIGN_TRACE_TRACE_SPACING", "MODEL_NAME" };
        string[] attrsImpeListsIMP_ITEM_New = { "ARTWORK_TRACE_WIDTH", "CALC_REQ_TRACE_WIDTH_TOL_PLUS", "DESIGN_TRACE_TRACE_SPACING", "MODEL_NAME", "REF_LAYER_2" };
        int jobNum_CORE;
        int jobNum_PREPREG;
        int jobNum_FOIL;
        int jobNum_STACKUP_SEG;

        int jobNum_CU_LAYER;
        int jobNum_DIE_LAYER;
        int jobNum_SM_LAYER;
        int jobNum_DRILL_LIST;
        int jobNum_IMP_ITEM;

        string s1 = "chenxing";
        string s2 = "founder@666";
        List<string[]> listStackupSeg;
        //List<string[]> listDIE_LAYER;
        List<string[]> listSM_LAYER;

        List<Dictionary<string, string>> listFoilT;
        List<Dictionary<string, string>> listCoreT;
        List<Dictionary<string, string>> listPrepregT;
        List<Dictionary<string, string>> listCU_LAYERT;
        List<Dictionary<string, string>> listDieLayerT;
        List<Dictionary<string, string>> listDRILL_LISTT;
        List<Dictionary<string, string>> listIMP_ITEMT;

        Dictionary<string, string>[] dictFoilT;
        Dictionary<string, string>[] dictCoreT;
        Dictionary<string, string>[] dictPrepregT;

        string base_Material;
        string customerName;
        double job_CU_FINISH_THICKNESS;

        public SISOLVERTEST()
        {
            InitializeComponent();
        }

        public SISOLVERTEST(string _jobname)
        {
            InitializeComponent();
            jobName = _jobname;
        }

        private void SISOLVERTESTLoad(object sender, EventArgs e)
        {
            if (jobName is null)
            {
                jobName = "HE08N84023D0";
                jobOutStep();
            }
            else
            {
                JobInStep();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(string.Format("当前打开的Job为： {0}", jobName));
            //creatJob();
            GetJobData();
            CreateXmlDoc();
            //string filepath = @"C:\SiSolverCalculate\test\API接口简单示例.xml";
            //InterfaceOfSiSolver(filepath);
        }

        private void GetJobData()
        {
            GetMaterial();  
            GetStackupSeg();
            GetCopperLayer();
            GetDrillLayer();
            GetDieLayer();
            GetMaskLayer();
            GetImpedance_Lists2();
        }

        /// <summary>
        /// 获取Materials
        /// 已获取，保存数据即可
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetMaterial()
        {
            //根据MRP_NAME 属性 去掉重复的 Material  FOIL CORE PREPREG  
            Dictionary<string, string> dictFoilMrpName = new Dictionary<string, string>();
            Dictionary<string, string> dictPrepregMrpName = new Dictionary<string, string>();
            Dictionary<string, string> dictCoreMrpName = new Dictionary<string, string>();

            //根据MRP_NAME 属性 去掉重复的 Material  FOIL CORE PREPREG  
            Dictionary<string, string> dictFoilMrpName_N = new Dictionary<string, string>();
            Dictionary<string, string> dictPrepregMrpName_N = new Dictionary<string, string>();
            Dictionary<string, string> dictCoreMrpName_N = new Dictionary<string, string>();

            int Total_Core = 0;
            int Total_Foil = 0;
            int Total_Prepreg = 0;
            stackup = theJob.Stackup();//job->stackup

            //1、获取数组的长度  {CORE}[X]  [FOIL][X] <PREPREG>[X]
            foreach (IStackupSegment seg in stackup.BuildUpSegments()) //stackup->stksegements->stksegement
            {
                //***************生成stakupsegment 
                seg.SegmentIndex();
                seg.SegmentType();
                //seg.JobMaterials();    //stackupsegement
                foreach (IJobMaterial jobmat in seg.JobMaterials())//stksegement->jobmaterials
                {
                    IMaterial mat = jobmat.Material();//jobmaterials->jobmaterial

                    //***************生成stakupsegment 
                    mat.MrpName();
                    if (mat.MtrType() == MtrType.CORE_MTR)//material->分类  material(CORE)
                    {
                        if (!dictCoreMrpName.Keys.Contains(mat.MrpName()))
                        {
                            dictCoreMrpName.Add(mat.MrpName(), "第" + (Total_Core + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + mat.MrpName());
                            Total_Core++;
                        }

                    }
                    if (mat.MtrType() == MtrType.FOIL_MTR)//material->分类  material(CORE)
                    {
                        if (!dictFoilMrpName.Keys.Contains(mat.MrpName()))
                        {
                            dictFoilMrpName.Add(mat.MrpName(), "第" + (Total_Foil + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + mat.MrpName());
                            Total_Foil++;
                        }
                    }
                    if (mat.MtrType() == MtrType.PREPREG_MTR)//material->分类  material(CORE)
                    {
                        if (!dictPrepregMrpName.Keys.Contains(mat.MrpName()))
                        {
                            dictPrepregMrpName.Add(mat.MrpName(), "第" + (Total_Prepreg + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + mat.MrpName());
                            Total_Prepreg++;
                        }

                    }
                }
            }

            //2 给键值对数组建立固定数组赋初值
            Dictionary<string, dynamic> udaObj = new Dictionary<string, dynamic>();

            Dictionary<string, string>[] dictFoil = new Dictionary<string, string>[Total_Foil];
            Dictionary<string, string>[] dictCore = new Dictionary<string, string>[Total_Core];
            Dictionary<string, string>[] dictPrepreg = new Dictionary<string, string>[Total_Prepreg];

            for (int arrayLenth = 0; arrayLenth < dictFoil.Length; arrayLenth++)
            {
                dictFoil[arrayLenth] = new Dictionary<string, string>();
            }
            for (int arrayLenth = 0; arrayLenth < dictCore.Length; arrayLenth++)
            {
                dictCore[arrayLenth] = new Dictionary<string, string>();
            }
            for (int arrayLenth = 0; arrayLenth < dictPrepreg.Length; arrayLenth++)
            {
                dictPrepreg[arrayLenth] = new Dictionary<string, string>();
            }

            //给键值对数组  添加指定属性的元素
            //数组下标

            int i, j, k;
            i = 0;
            j = 0;
            k = 0;



            try
            {
                foreach (IStackupSegment seg in stackup.BuildUpSegments()) //stackup->stksegements->stksegement
                {


                    foreach (IJobMaterial jobmat in seg.JobMaterials())//stksegement->jobmaterials
                    {
                        IMaterial mat = jobmat.Material();//jobmaterials->jobmaterial
                        //base_Material = mat.Family();
                        if (mat.MtrType() == MtrType.FOIL_MTR)//material->分类  material(CORE)
                        {
                            //针对CORE属性  1、CORE所有属性  2、属性筛选 3、添加键值对
                            IFoil foil = mat.Foil();
                            Dictionary<string, string> dictAllFoil = new Dictionary<string, string>();
                            XmlGenerate.AllAttrGetKeyValues(foil, dictAllFoil);
                            XmlGenerate.AllAttrGetKeyValues(mat, dictAllFoil);
                            //Base_Material = dictAllFoil["FAMILY"];
                            string[] a = attrsMaterialsFoil;
                            string[] b = attrsMaterialsFoil_Inplan;
                            int length = 0;
                            if (a.Length == b.Length)
                            {
                                length = a.Length;
                            }
                            else
                            {
                                string str = "请检查是否数组长度定义一一对应";
                            }

                            if (!dictFoilMrpName_N.Keys.Contains(dictAllFoil["MRP_NAME"]))
                            {
                                dictFoilMrpName_N.Add(dictAllFoil["MRP_NAME"], "第" + (i + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + dictAllFoil["MRP_NAME"]);
                                for (int l = 0; l < length; l++)
                                {
                                    dictFoil[i].Add(a[l], dictAllFoil[b[l]]);
                                }
                                i++;
                            }
                            string debug = "";

                        }
                        if (mat.MtrType() == MtrType.PREPREG_MTR)//material->分类  material(CORE)
                        {
                            IPrepreg prepreg = mat.Prepreg();
                            Dictionary<string, string> dictAllPrepreg = new Dictionary<string, string>();
                            XmlGenerate.AllAttrGetKeyValues(prepreg, dictAllPrepreg);
                            XmlGenerate.AllAttrGetKeyValues(mat, dictAllPrepreg);
                            base_Material = dictAllPrepreg["FAMILY"];
                            string[] a = attrsMaterialsPrepreg;
                            string[] b = attrsMaterialsPrepreg_Inplan;
                            int length = 0;
                            if (a.Length == b.Length)
                            {
                                length = a.Length;
                            }
                            else
                            {
                                string str = "请检查是否数组长度定义一一对应";
                            }
                            if (!dictPrepregMrpName_N.Keys.Contains(dictAllPrepreg["MRP_NAME"]))
                            {
                                dictPrepregMrpName_N.Add(dictAllPrepreg["MRP_NAME"], "第" + (j + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + dictAllPrepreg["MRP_NAME"]);
                                for (int l = 0; l < length; l++)
                                {
                                    dictPrepreg[j].Add(a[l], dictAllPrepreg[b[l]]);
                                }
                                j++;
                            }
                            string debug = "";

                        }
                        if (mat.MtrType() == MtrType.CORE_MTR)//material->分类  material(CORE)
                        {
                            ICore core = mat.Core();
                            Dictionary<string, string> dictAllCore = new Dictionary<string, string>();
                            XmlGenerate.AllAttrGetKeyValues(core, dictAllCore);
                            XmlGenerate.AllAttrGetKeyValues(mat, dictAllCore);
                            base_Material = dictAllCore["FAMILY"];
                            string[] a = attrsMaterialsCore;
                            string[] b = attrsMaterialsCore_Inplan;
                            int length = 0;
                            if (a.Length == b.Length)
                            {
                                length = a.Length;
                            }
                            else
                            {
                                string str = "请检查是否数组长度定义一一对应";
                            }
                            if (!dictCoreMrpName_N.Keys.Contains(dictAllCore["MRP_NAME"]))
                            {
                                dictCoreMrpName_N.Add(dictAllCore["MRP_NAME"], "第" + (k + 1).ToString() + "种类别的" + mat.MtrType() + ":  " + dictAllCore["MRP_NAME"]);
                                for (int l = 0; l < length; l++)
                                {
                                    dictCore[k].Add(a[l], dictAllCore[b[l]]);
                                }
                                k++;
                            }
                            string debug = "";

                        }
                    }
                }

                //对应XML子节点个数
                jobNum_FOIL = Total_Foil;
                jobNum_PREPREG = Total_Prepreg;
                jobNum_CORE = Total_Core;
                jobNum_STACKUP_SEG = Total_Foil + Total_Prepreg + Total_Core;

                listFoilT = new List<Dictionary<string, string>>();
                listPrepregT = new List<Dictionary<string, string>>();
                listCoreT = new List<Dictionary<string, string>>();
                foreach (var item in dictFoil)
                {

                    listFoilT.Add(item);
                }
                foreach (var item in dictPrepreg)
                {
                    listPrepregT.Add(item);
                }
                foreach (var item in dictCore)
                {
                    listCoreT.Add(item);
                }
                dictFoilT = dictFoil;
                dictPrepregT = dictPrepreg;
                dictCoreT = dictCore;
            }
            catch (Exception ex)
            {
                string exmess = ex.ToString();
                //throw;
            }
        }

        /// <summary>
        /// 获取STACKUP_Segments
        /// </summary>
        private void GetStackupSeg()
        {
            List<string[]> list = new List<string[]>();
            // "SEG_INDEX", "SEG_TYPE", "MRP_CODE" 
            int j = 1;
            foreach (IStackupSegment seg in stackup.BuildUpSegments()) //stackup->stksegements->stksegement
            {
                foreach (IJobMaterial jobmat in seg.JobMaterials())//stksegement->jobmaterials
                {
                    string[] str = new string[3];
                    IMaterial mat = jobmat.Material();//jobmaterials->jobmaterial
                                                      //str[0] = seg.SegmentIndex().ToString();
                    str[0] = j.ToString();

                    if (seg.SegmentType().ToString() == StkSegmentType.FOIL_SEG.ToString())
                    {
                        str[1] = "Foil";
                    }
                    else if (seg.SegmentType().ToString() == StkSegmentType.ISOLATOR_SEG.ToString())
                    {
                        str[1] = "Prepreg";
                    }
                    else if (seg.SegmentType().ToString() == StkSegmentType.CORE_SEG.ToString())
                    {
                        str[1] = "Core";
                    }

                    str[2] = mat.MrpName();
                    list.Add(str);
                    j++;
                }

            }
            listStackupSeg = list;
        }

        /// <summary>
        /// 获取Process_Parameters  CopperLayer 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetCopperLayer()
        {
            int totalCoperLayer;
            totalCoperLayer = 0;

            //先对coperlayer取值
            //theJob.CopperLayers();
            foreach (ICopperLayer copper in theJob.CopperLayers())
            {
                totalCoperLayer = totalCoperLayer + 1;
            }

            jobNum_CU_LAYER = totalCoperLayer;

            Dictionary<string, string>[] dictCU_LAYER = new Dictionary<string, string>[totalCoperLayer];

            int i;
            i = 0;

            for (int dictLength = 0; dictLength < dictCU_LAYER.Length; dictLength++)
            {
                dictCU_LAYER[dictLength] = new Dictionary<string, string>();
            }

            int len_process = XmlGenerate.jobcomponslength(theJob.AllProcesses(false));
            string[] AllCuFinish_Thickness = new string[len_process];
            //获取：  //每一层的原始铜厚

            int j = 0;
            foreach (IProcess process in theJob.AllProcesses(false))
            {
                //Dictionary<string, string> dictAllProcess = new Dictionary<string, string>();
                //XmlGenerate.AllAttrGetKeyValues(process, dictAllProcess);
                Dictionary<string, double> dictAllProcessDoubleData = new Dictionary<string, double>();
                XmlGenerate.AllAttrGetKeyValues_double(process, dictAllProcessDoubleData);
                double d1 = dictAllProcessDoubleData["PLATING_THK_MAX_"];
                double d2 = dictAllProcessDoubleData["PLATING_THK_MIN_"];
                double a3 = (d1 + d2) / 2;
                AllCuFinish_Thickness[j] = a3.ToString();
                //线路完成面铜厚度取Procee（Final Assembly）下PLATING_THK_MIN_和PLATING_THK_MAX_的平均值
                job_CU_FINISH_THICKNESS = a3;
                j++;
            }


            try
            {
                foreach (ICopperLayer copper in theJob.CopperLayers())
                {
                    string CU_LAYER_INDEX = i.ToString();//  添加到最后生成的键值对中  
                    Dictionary<string, string> dictAllCU_LAYER = new Dictionary<string, string>();
                    XmlGenerate.AllAttrGetKeyValues(copper, dictAllCU_LAYER);
                    //string str =dictAllCU_LAYER["d "];
                    dictCU_LAYER[i].Add("CU_LAYER_INDEX", (i + 1).ToString());
                    if (i == 0)
                    {
                        dictCU_LAYER[i].Add("CU_FINISH_THICKNESS", AllCuFinish_Thickness[0]);
                    }
                    else
                    {
                        //dictCU_LAYER[i].Add("CU_FINISH_THICKNESS", dictAllCU_LAYER["OVERALL_THICKNESS"]);
                        dictCU_LAYER[i].Add("CU_FINISH_THICKNESS", dictAllCU_LAYER["LAYER_FINISH_CU_THK_"]);
                        //dictCU_LAYER[i].Add("CU_FINISH_THICKNESS", copper.RequiredThickness(AvailableUnits.MIL).ToString());

                    }
                    //下划线可不输入
                    //dictCU_LAYER[i].Add("CU_RAW_THICKNESS_", dictAllCU_LAYER["COPPER_USAGE"]);
                    dictCU_LAYER[i].Add("COPPER_USAGE", dictAllCU_LAYER["COPPER_USAGE"]);
                    dictCU_LAYER[i].Add("CU_RAW_WEIGHT", dictAllCU_LAYER["REQUIRED_CU_WEIGHT"]);
                    //字段待确定
                    dictCU_LAYER[i].Add("CU_UNDERCUT", ""); //待确定
                    //dictCU_LAYER[i].Add("LAYER_SIGNAL_TYPE_", dictAllCU_LAYER["COPPER_USAGE"]);
                    i++;
                }
                listCU_LAYERT = new List<Dictionary<string, string>>();
                foreach (var item in dictCU_LAYER)
                {
                    listCU_LAYERT.Add(item);
                }
            }
            catch (Exception ex)
            {
                string exmess = ex.ToString();
                //throw;
            }
        }

        /// <summary>
        /// 获取MaskLayer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetDieLayer()
        {
            //CopperLayresCount
            int boardlayers = theJob.CopperLayresCount();

            Dictionary<string, string>[] dictDieLayer = new Dictionary<string, string>[boardlayers - 1];
            jobNum_DIE_LAYER = dictDieLayer.Length;
            for (int dictLength = 0; dictLength < dictDieLayer.Length; dictLength++)
            {
                dictDieLayer[dictLength] = new Dictionary<string, string>();
            }
            for (int i = 0; i < dictDieLayer.Count(); i++)
            {
                dictDieLayer[i].Add("DIE_LAYER_INDEX", (i + 1).ToString());
                dictDieLayer[i].Add("DIE_RAW_THICKNESS", "0");
                dictDieLayer[i].Add("DIE_FINISH_THICKNESS", "0");
                dictDieLayer[i].Add("RESIN_FILLED_CU_THK", "0");
                dictDieLayer[i].Add("RESIN_FLOW_THK", "0");
                dictDieLayer[i].Add("RESIN_FILLED_VIA_THK", "0");
            }
            listDieLayerT = new List<Dictionary<string, string>>();
            foreach (var item in dictDieLayer)
            {
                listDieLayerT.Add(item);
            }


        }

        /// <summary>
        /// 获取MaskLayer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetMaskLayer()
        {
            List<string[]> list = new List<string[]>();
            string[] mask = new string[4];
            mask[0] = "4";//默认4.0

            double CU_FINISH_THICKNESS = job_CU_FINISH_THICKNESS;
            double T1;
            double T2;
            double T3;
            double T4;

            T1 = XmlGenerate.UmToMil(46);
            T2 = XmlGenerate.UmToMil(63.5);
            T3 = XmlGenerate.UmToMil(89);
            T4 = XmlGenerate.UmToMil(114);

            double c1; //基材防焊厚度
            double c2; //线路防焊厚度
            double c3; //基材防焊厚度

            if (CU_FINISH_THICKNESS < T1)
            {
                c1 = XmlGenerate.UmToMil(38);
                c2 = XmlGenerate.UmToMil(18);
                c3 = XmlGenerate.UmToMil(18);
                mask[1] = c1.ToString();
                mask[2] = c2.ToString();
                mask[3] = c3.ToString();
            }
            else if (CU_FINISH_THICKNESS >= T1 && CU_FINISH_THICKNESS < T2)
            {
                c1 = XmlGenerate.UmToMil(48);
                c2 = XmlGenerate.UmToMil(18);
                c3 = XmlGenerate.UmToMil(18);
                mask[1] = c1.ToString();
                mask[2] = c2.ToString();
                mask[3] = c3.ToString();
            }
            else if (CU_FINISH_THICKNESS >= T2 && CU_FINISH_THICKNESS < T3)
            {
                c1 = XmlGenerate.UmToMil(58);
                c2 = XmlGenerate.UmToMil(30);
                c3 = XmlGenerate.UmToMil(30);
                mask[1] = c1.ToString();
                mask[2] = c2.ToString();
                mask[3] = c3.ToString();
            }
            else if (CU_FINISH_THICKNESS >= T3 && CU_FINISH_THICKNESS < T4)
            {
                c1 = XmlGenerate.UmToMil(90);
                c2 = XmlGenerate.UmToMil(30);
                c3 = XmlGenerate.UmToMil(30);
                mask[1] = c1.ToString();
                mask[2] = c2.ToString();
                mask[3] = c3.ToString();
            }


            list.Add(mask);
            listSM_LAYER = list;
        }

        /// <summary>
        /// 获取DrillLayer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetDrillLayer()
        {
            int totalDrllPrgs;
            totalDrllPrgs = 0;
            try
            {
                foreach (IDrillProgram lDrllPrgs in theJob.DrillPrgs())
                {
                    totalDrllPrgs = totalDrllPrgs + 1;
                }

                jobNum_DRILL_LIST = totalDrllPrgs;

                Dictionary<string, string>[] dictDRILL_LIST = new Dictionary<string, string>[totalDrllPrgs];

                int i;
                i = 0;

                for (int dictLength = 0; dictLength < dictDRILL_LIST.Length; dictLength++)
                {
                    dictDRILL_LIST[dictLength] = new Dictionary<string, string>();
                }

                foreach (IDrillProgram DrllPrgs in theJob.DrillPrgs())
                {

                    Dictionary<string, string> dictAllDRILL_LIST = new Dictionary<string, string>();
                    XmlGenerate.AllAttrGetKeyValues(DrllPrgs, dictAllDRILL_LIST);
                    dictDRILL_LIST[i].Add("VIA_TYPE", dictAllDRILL_LIST["DRILL_TYPE"]);
                    dictDRILL_LIST[i].Add("START_LAYER", dictAllDRILL_LIST["START_INDEX"]);
                    dictDRILL_LIST[i].Add("END_LAYER", dictAllDRILL_LIST["END_INDEX"]);
                    i++;
                }
                listDRILL_LISTT = new List<Dictionary<string, string>>();
                foreach (var item in dictDRILL_LIST)
                {
                    listDRILL_LISTT.Add(item);
                }
            }
            catch (Exception ex)
            {
                string str = ex.ToString();
                //throw;
            }
        }

        /// <summary>
        /// 获取阻抗 Impedance_Lists
        /// path: 由于获取阻抗是 STACKUP 中取
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetImpedance_Lists2()
        {
            //stackup 获取阻抗
            stackup = theJob.Stackup();
            int totalImpCon;
            int totalImpLin;
            totalImpCon = 0;
            totalImpLin = 0;
            //IImpedanceConstraint
            foreach (IImpedanceConstraint impCon in stackup.ImpedanceConstraints())
            {
                totalImpCon = totalImpCon + 1;
            }
            jobNum_IMP_ITEM = totalImpCon;
            Dictionary<string, string>[] dictImpCon = new Dictionary<string, string>[totalImpCon];
            Dictionary<string, string>[] dictImpLin = new Dictionary<string, string>[totalImpLin];

            for (int dictImpConLength = 0; dictImpConLength < dictImpCon.Length; dictImpConLength++)
            {
                dictImpCon[dictImpConLength] = new Dictionary<string, string>();
            }
            for (int dictImpLinLength = 0; dictImpLinLength < dictImpLin.Length; dictImpLinLength++)
            {
                dictImpLin[dictImpLinLength] = new Dictionary<string, string>();
            }
            int i;
            i = 0;
            try
            {
                foreach (IImpedanceConstraint impCon in stackup.ImpedanceConstraints())
                {
                    Dictionary<string, string> dictAllImpCon = new Dictionary<string, string>();
                    XmlGenerate.AllAttrGetKeyValues(impCon, dictAllImpCon);

                    //阻抗模型 取值  OK
                    //编号取层 OK
                    //字段取值 NOT   NUM_SIGNAL="待确定"

                    int modecode = (int)Enum.Parse(typeof(XmlGenerate.ImdependenceModel), impCon.ModelName(), true);

                  

                    //字段取值
                    //string modelname=impCon.ModelName();
                    //信号层 
                    ICopperLayer trace_layer;
                    trace_layer = impCon.ControlledTraceLayer();
                    string tracelayer = trace_layer.Name();
                    if (tracelayer == "COMP")
                    {
                        tracelayer = "1";
                    }
                    else if (tracelayer == "SOLD")
                    {
                        tracelayer = jobName.Substring(2, 2);
                    }

                    //参考层   TopModelLayer  BottomModelLayer
                    ICopperLayer copperLayertop;
                    ICopperLayer copperLayerbot;
                    copperLayertop = impCon.TopModelLayer();
                    copperLayerbot = impCon.BottomModelLayer();
                    string reftop;
                    string refbot;
                    if (impCon.TopModelLayer() is null)
                    {
                        reftop = "0";
                    }
                    else
                    {
                        reftop = copperLayertop.Name();
                    }

                    if (impCon.BottomModelLayer() is null)
                    {
                        refbot = "0";
                    }
                    else
                    {
                        refbot = copperLayerbot.Name();
                    }

                    //7和8这个模型，单独只取BOTTOM_MODEL_LAYER可行  时间：2019年7月22日22:26:26
                    if (modecode == 7 || modecode == 8)
                    {
                        reftop = "0";
                        refbot = copperLayerbot.Name();
                    }


                    if (reftop == "COMP")
                    {
                        reftop = "1";// L去掉
                    }
                    else if (reftop == "SOLD")
                    {
                        reftop = "" + jobName.Substring(2, 2);
                    }

                    if (refbot == "COMP")
                    {
                        refbot = "1";
                    }
                    else if (refbot == "SOLD")
                    {
                        refbot = "" + jobName.Substring(2, 2);
                    }



                    // 字段： 区分单根双根阻抗  NUM_SIGNAL  待确定
                    // 字段：目标阻抗值   ZO_TARGET CalculationRequiredImpedance()

                    //发现一个问题： 客户要求阻抗不在JOB中
                    double omCustomer=impCon.CustomerRequiredImpedance(AvailableUnits.OHMS);
                    double LWCustomer = impCon.OriginalTraceWidth(AvailableUnits.MIL);
                    double LSCopCustomer = impCon.CustomerRequiredCoplanarSpacing(AvailableUnits.MIL);
                    double LSDifCustomer = impCon.CustomerRequiredDifferentialSpacing(AvailableUnits.MIL);



                    double omdouble = impCon.CalculationRequiredImpedance(AvailableUnits.OHMS);
                    double omtol = omdouble * 0.1;
                    string omstring = impCon.CalculationRequiredImpedance(AvailableUnits.OHMS).ToString();
                    // 字段： 目标阻抗值公差  ZO_TOL  （10%）  CalculationRequiredImpedancePlusTol CalculationRequiredImpedanceMinusTol
                    string tolp = impCon.CalculationRequiredImpedancePlusTol(AvailableUnits.OHMS).ToString();
                    string tolm = impCon.CalculationRequiredImpedanceMinusTol(AvailableUnits.OHMS).ToString();
                    
                    //字段： 设计线宽  WIDTH_DESIGN CALCULATION_REQ_TRACE_WIDTH
                    //已更改
                    double width_double = impCon.CalculationReqLineWidth(AvailableUnits.MIL);
                    string width = impCon.CalculationReqLineWidth(AvailableUnits.MIL).ToString();
                    //设计间距(差分间距) SPACE_DESIGN   DIFF_SPACE_
                    //共面阻抗间距  COPLANAR_SPACE_DESIGN  COPLANARITY_SPACE_
                    bool cop = impCon.IsNullCustomerRequiredCoplanarSpacing();
                    bool sx = impCon.IsNullCustomerRequiredDifferentialSpacing();
                    //返回值即可  UDAS属性 值  
                    //设计间距(差分间距) SPACE_DESIGN   DIFF_SPACE_
                    //共面阻抗间距  COPLANAR_SPACE_DESIGN  COPLANARITY_SPACE_

                    dictImpCon[i].Add("ITEM_INDEX", (i + 1).ToString());
                    dictImpCon[i].Add("IMP_MODEL_NO", modecode.ToString());
                    dictImpCon[i].Add("SIGNAL_LAYER", tracelayer.Replace("L", ""));//删除L
                    dictImpCon[i].Add("REF_LAYER_1", reftop.Replace("L", ""));
                    dictImpCon[i].Add("REF_LAYER_2", refbot.Replace("L", ""));

                    if (LSDifCustomer > 0)
                    {
                        dictImpCon[i].Add("NUM_SIGNAL", "2");
                    }
                    else if (LSDifCustomer == 0)
                    {
                        dictImpCon[i].Add("NUM_SIGNAL", "1");
                    }

                    //dictImpCon[i].Add("NUM_SIGNAL", (width_double > 0 ? "2" : "1"));
                    //dictImpCon[i].Add("NUM_SIGNAL", " "); //待确定 omCustomer
                    dictImpCon[i].Add("ZO_TARGET", omCustomer.ToString());
                    //客户要求阻抗
                    dictImpCon[i].Add("ZO_TOL", (omCustomer*0.1).ToString());
                    //dictImpCon[i].Add("ZO_TOL", omtol.ToString());
                    dictImpCon[i].Add("WIDTH_DESIGN", LWCustomer.ToString());
                    //LSCopCustomer
                    dictImpCon[i].Add("SPACE_DESIGN", LSDifCustomer.ToString());
                    //dictImpCon[i].Add("SPACE_DESIGN", dictAllImpCon["DIFF_SPACE_"]);
                    dictImpCon[i].Add("COPLANAR_SPACE_DESIGN", LSCopCustomer.ToString());
                    //dictImpCon[i].Add("COPLANAR_SPACE_DESIGN", dictAllImpCon["COPLANARITY_SPACE_"]);

                    i++;
                }
                listIMP_ITEMT = new List<Dictionary<string, string>>();
                foreach (var item in dictImpCon)
                {
                    listIMP_ITEMT.Add(item);
                }


            }
            catch (Exception ex)
            {
                string str = ex.ToString();
                //throw;
            }

        }

        /// <summary>
        /// 创建XML到文档并调用SiSolver生成阻抗报告
        /// </summary>
        private void CreateXmlDoc()
        {
            //1、创建XML文档实例
            XmlDocument xmldoc = new XmlDocument();
            string version = "1.0";
            string encoding = "UTF-8";
            string standalone = "yes";
            //2、创建文档声明
            XmlDeclaration dec = xmldoc.CreateXmlDeclaration(version, encoding, standalone);
            xmldoc.AppendChild(dec);
            //3、创建文档元素
            //Doc节点
            XmlElement rootNode = xmldoc.CreateElement("Document");
            //一级   节点设置属性
            try
            {
                int length = attrDocument.Length;
                Dictionary<string, string> dictDocument = new Dictionary<string, string>();
                string[] attrDocument_values = new string[length];
                string[] attrDocument_keys = new string[length];

                GetCustomer();

                dictDocument.Add("Dev_Owner", "SISOLVER");
                dictDocument.Add("Fab", "FOUNDERPCBF5");
                dictDocument.Add("Version", "1.0");
                dictDocument.Add("Project_Name", "Test");
                dictDocument.Add("Project_Ver", "A0");
                dictDocument.Add("Project_Designer", "SISOLVER");
                dictDocument.Add("Date", DateTime.Now.ToString());
                dictDocument.Add("Layer_Count", jobName.Substring(2, 2));
                dictDocument.Add("Base_Material", base_Material); //base_Material
                dictDocument.Add("Customer", "FOUNDERPCBF5");
                dictDocument.Add("Stackup_Ready", "YES");
                dictDocument.Add("Imp_Calc_Fixed", "NO");
                dictDocument.Add("Width_Min", "2"); //Width_Min="待确认"  这个预设为2

                int k = 0;
                foreach (var item in dictDocument.Keys)
                {
                    attrDocument_keys[k] = item.ToString();
                    k++;
                }

                int v = 0;
                foreach (var item in dictDocument.Values)
                {
                    attrDocument_values[v] = item.ToString();
                    v++;
                }

                XmlGenerate.AutoGerateAttr(xmldoc, rootNode, attrDocument_keys, attrDocument_values);
            }
            catch (Exception)
            {

                throw;
            }
            //二级   节点设置属性
            XmlElement nodeChild0 = xmldoc.CreateElement("Materials");
            XmlElement nodeChild1 = xmldoc.CreateElement("Materials");
            XmlElement nodeChild2 = xmldoc.CreateElement("Materials");
            XmlElement nodeChild3 = xmldoc.CreateElement("STACKUP_Segments");
            XmlElement nodeChild4 = xmldoc.CreateElement("Process_Parameters");
            XmlElement nodeChild5 = xmldoc.CreateElement("Process_Parameters");
            XmlElement nodeChild6 = xmldoc.CreateElement("Process_Parameters");
            XmlElement nodeChild7 = xmldoc.CreateElement("Process_Parameters");
            XmlElement nodeChild8 = xmldoc.CreateElement("Impedance_Lists");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild0, attrMaterials, "CORE");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild1, attrMaterials, "PREPREG");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild2, attrMaterials, "FOIL");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild4, attrProcess_Parameters, "CU_LAYER");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild5, attrProcess_Parameters, "DIE_LAYER");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild6, attrProcess_Parameters, "SM_LAYER");
            XmlGenerate.AutoGerateCNodeAttr(rootNode, nodeChild7, attrProcess_Parameters, "DRILL_LIST");

            //三级  节点设置属性
            //Job对象： CORE、PREPREG、FOIL、STACKUP_SEG、CU_LAYER、DIE_LAYER、SM_LAYER、DRILL_LIST、IMP_ITEM
            // 三级节点个数   由Job决定

            XmlElement[] xmlTEle_CORE = new XmlElement[jobNum_CORE];          //Job对应的所有CORE
            XmlElement[] xmlTEle_PREPREG = new XmlElement[jobNum_PREPREG];       //Job对应的所有PREPREG
            XmlElement[] xmlTEle_FOIL = new XmlElement[jobNum_FOIL];          //Job对应的所有FOIL
            XmlElement[] xmlTEle_STACKUP_SEG = new XmlElement[jobNum_STACKUP_SEG];   //Job对应的所有STACKUP_SEG
            XmlElement[] xmlTEle_CU_LAYER = new XmlElement[jobNum_CU_LAYER];      //Job对应的所有CU_LAYER
            XmlElement[] xmlTEle_DIE_LAYER = new XmlElement[jobNum_DIE_LAYER];     //Job对应的所有DIE_LAYER
            XmlElement[] xmlTEle_SM_LAYER = new XmlElement[1];      //Job对应的所有SM_LAYER
            XmlElement[] xmlTEle_DRILL_LIST = new XmlElement[jobNum_DRILL_LIST];    //Job对应的所有DRILL_LIST
            XmlElement[] xmlTEle_IMP_ITEM = new XmlElement[jobNum_IMP_ITEM];      //Job对应的所有IMP_ITEM

            string stop = "";

            try
            {
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild0, xmlTEle_CORE, listCoreT, "CORE");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild1, xmlTEle_PREPREG, listPrepregT, "PREPREG");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild2, xmlTEle_FOIL, listFoilT, "FOIL");
                XmlGenerate.SetAttrsByArrayToJobAttrs(xmldoc, nodeChild3, xmlTEle_STACKUP_SEG, attrsMaterialsStackupSegms, listStackupSeg, "STACKUP_SEG");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild4, xmlTEle_CU_LAYER, listCU_LAYERT, "CU_LAYER");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild5, xmlTEle_DIE_LAYER, listDieLayerT, "DIE_LAYER");
                XmlGenerate.SetAttrsByArrayToJobAttrs(xmldoc, nodeChild6, xmlTEle_SM_LAYER, attrsProParaSM_LAYER, listSM_LAYER, "SM_LAYER");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild7, xmlTEle_DRILL_LIST, listDRILL_LISTT, "DRILL_LIST");
                XmlGenerate.SetAttrsByArrayToJobAttrs3(xmldoc, nodeChild8, xmlTEle_IMP_ITEM, listIMP_ITEMT, "IMP_ITEM");
            }
            catch (Exception)
            {
                //throw;
            }

            //二级节点 一级节点
            XmlNode[] node = { nodeChild0, nodeChild1, nodeChild2, nodeChild3, nodeChild4, nodeChild5, nodeChild6, nodeChild7, nodeChild8 };
            foreach (var item in node)
            {
                rootNode.AppendChild(item);
            }
            //一级节点 顶级节点
            //jobName = "ME10N20JG6A3";//"GE02N20FQ4A0";
            string folderpath = string.Format("C:\\SiSolverCalculate\\{0}\\", jobName);
            string filepath = string.Format("C:\\SiSolverCalculate\\{0}\\{1}.xml", jobName, jobName);
            string p1 = string.Format("C:\\SiSolverCalculate\\{0}\\", jobName);
            string filename = string.Format("{0}.xml", jobName);
            if (!Directory.Exists(folderpath))
            {
                Directory.CreateDirectory(folderpath);
                //创建文件
                xmldoc.Save(filepath);
                XmlGenerate.InterfaceOfSiSolver(filepath);
            }
            else
            {
                //文件是否存在，创建文件
                if (File.Exists(filepath))
                {
                    //判断是否存在文件
                    //MessageBox.Show("阻抗已成功生成，如继续将生成的文件");
                    string str = XmlGenerate.fileExists(filepath, p1, times, jobName);
                    xmldoc.Save(str);
                    Thread.Sleep(10000);
                    XmlGenerate.InterfaceOfSiSolver(str);//p1 + str
                    //times = 1;
                }

                else
                {
                    xmldoc.Save(filepath);
                    Thread.Sleep(10000);
                    XmlGenerate.InterfaceOfSiSolver(filepath);
                }
            }
        }

        private void GetCustomer()
        {
            ICustomer customer = theJob.Customer();
            customerName = customer.Name();
        }

        private void creatJob()
        {
            //appManager = new ApplicationManager();
            jobManager = appManager.JobManager();
            if (jobManager.ErrorStatus() != 0)
            {
                MessageBox.Show(jobManager.ErrorMessage());
            }
            theJob = jobManager.OpenJob(jobName);
            if (theJob is null)
            {
                MessageBox.Show("获取的Job对象为空");
            }
            //MessageBox.Show(string.Format(" {0}",? (theJob.JobTypeStr().Length>0): "已成功获取Job","获取的Job对象为空"));
            if (appManager.ErrorStatus() != 0)
            {
                string logoex = "";
                //消息框弹窗内容、弹窗标题、确定按钮
                MessageBox.Show(appManager.ErrorMessage(), "Error", MessageBoxButtons.OK);
                Application.Exit();
            }
        }

        private void LoginInplan()
        {
            appManager.Login(s1, s2);
            theJob = jobManager.OpenJob(jobName);
            if (appManager.ErrorStatus() != 0)
            {
                string logoex = "";
                //消息框弹窗内容、弹窗标题、确定按钮
                MessageBox.Show("Inplan外部登录：  "+appManager.ErrorMessage(), "Error", MessageBoxButtons.OK);
                Application.Exit();
            }

            if (theJob is null && logintimes <= 2)
            {
                logintimes = logintimes + 1;
                LoginInplan();

                //MessageBox.Show("未能成功登录");
            }
            else if (theJob.ErrorStatus() != 0)
            {
                MessageBox.Show(theJob.ErrorMessage());
            }
            else
            {
                int loveu10000 = 0;
                //MessageBox.Show("登录成功");
            }
            logintimes += 1;
        }

        private void jobOutStep()
        {
            try
            {
                appManager = new ApplicationManager();
                jobManager = appManager.JobManager();
                LoginInplan();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                //throw;
            }
            
        }

        private void JobInStep()
        {
            appManager = new ApplicationManager();
            jobManager = appManager.JobManager();
            if (jobManager.ErrorStatus() != 0)
            {
                MessageBox.Show(jobManager.ErrorMessage());
            }
            theJob = jobManager.OpenJob(jobName);
            if (appManager.ErrorStatus() != 0)
            {
                string logoex = "";
                //消息框弹窗内容、弹窗标题、确定按钮
                MessageBox.Show("Inplan内部登录  "+appManager.ErrorMessage(), "Error", MessageBoxButtons.OK);
                Application.Exit();
            }
            else if (theJob.ErrorStatus() != 0)
            {
                MessageBox.Show(theJob.ErrorMessage());
            }
            else if (!(theJob is null))
            {
                //MessageBox.Show("获取的Job对象不为空");
            }
            else if ((theJob is null))
            {
                MessageBox.Show("获取的Job对象为空");
            }
        }

    }
}
