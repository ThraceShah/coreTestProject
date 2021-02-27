using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Connect;
using System.IO;
using System.Text;
using Filedeal;

namespace SolidworksModel
{
    public class Solution
    {
        public static void IputTD(string item, string filePaths, double t, char type)
        {
            string sDwgFileName = filePaths + item + ".dxf";
            if (type == 'F')
                sDwgFileName = filePaths + item + ".dxf";
            else
                sDwgFileName = filePaths + item + ".dwg";
            if (File.Exists(sDwgFileName) == false)
                return;
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                //string msg = "This message from C#. solidworks version is " + swApp.RevisionNumber();
                //swApp.SendMsgToUser(msg);//显示solidworks的版本
                //通过GetDocumentTemplate 获取默认模板的路径 ,第一个参数可以指定类型
                string partDefaultTemplate = swApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);
                //也可以直接指定slddot asmdot drwdot
                //partDefaultTemplate = @"xxx\..prtdot";
                var newDoc = swApp.NewDocument(partDefaultTemplate, 0, 0, 0);
                //var newDoc = swApp.NewDocument(partDefaultTemplate, 0, 0, 0);
                if (newDoc != null)
                {
                    //创建完成
                    //swApp.SendMsgToUser("Create done.");
                    //下面获取当前文件
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                    //选择对应的草图基准面
                    bool boolstatus = swModel.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, false, 0, null, 0);
                    var importData = (ImportDxfDwgData)swApp.GetImportFileData(sDwgFileName);
                    // Unit
                    importData.set_LengthUnit("", (int)swLengthUnit_e.swMM);
                    // Position
                    var bRet = importData.SetPosition("", (int)swDwgImportEntitiesPositioning_e.swDwgEntitiesCentered, 0, 0);
                    // Sheet scale
                    bRet = importData.SetSheetScale("", 1.0, 2.0);
                    // Paper size
                    bRet = importData.SetPaperSize("", (int)swDwgPaperSizes_e.swDwgPaperAsize, 0.0, 0.0);
                    //Import method
                    importData.set_ImportMethod("", (int)swImportDxfDwg_ImportMethod_e.swImportDxfDwg_ImportToExistingPart);
                    // Import file with importData
                    //swFeat = swFeatMgr.InsertDwgOrDxfFile2(sDwgFileName, importData);
                    Feature myFeature = swModel.FeatureManager.InsertDwgOrDxfFile2(sDwgFileName, importData);
                    //Feature myFeature = swModel.FeatureManager.InsertDwgOrDxfFile(sDwgFileName);
                    swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, t, t, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, 0, 0, false);
                    if (myFeature != null)
                    {
                        string myNewPartPath = sDwgFileName.Substring(0, sDwgFileName.Length - 4) + ".SLDPRT";
                        int longstatus = swModel.SaveAs3(myNewPartPath, 0, 1);
                        swModel.SaveAs3(myNewPartPath, 0, 1);
                        swApp.CloseDoc(myNewPartPath);
                    }
                    else
                    {
                        swApp.CloseDoc(newDoc.ToString());
                        string myNewPartPath = "C:\\temp.SLDPRT";
                        int longstatus = swModel.SaveAs3(myNewPartPath, 0, 1);
                        swModel.SaveAs3(myNewPartPath, 0, 1);
                        swApp.CloseDoc(myNewPartPath);
                    }
                }
            }
        }

        internal static void ToRectangle(double[] dimensions, string filePaths, double t, string item)
        {
            string sDwgFileName = filePaths + item + ".dwg";
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                double swSheetWidth = 0.4;
                double swSheetHeight = 0.3;
                string programPath = System.Environment.CurrentDirectory;
                var newDoc = swApp.NewDocument(@"C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\空白工程图模板.drwdot", 12, swSheetWidth, swSheetHeight);//创建一个空白草图，模板放在根目录
                if (newDoc != null)
                {
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                    swModel.SketchManager.CreateCornerRectangle(0, 0, 0, dimensions[1], dimensions[0], 0);//创建一个矩形
                    swModel.ClearSelection2(true);//确认草图
                    ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;
                    int error = 0;
                    int warnings = 0;
                    //设置dxf 导出版本 R14
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, 2);
                    //是否显示 草图
                    swModel.SetUserPreferenceToggle(196, false);
                    swModExt.SaveAs(sDwgFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref error, ref warnings);
                    swApp.CloseAllDocuments(true);
                }
            }
        }

        public static void GetBodyMaxfaceTodwg(string sldprtName)
        {
            ModelDoc2 swModel;
            PartDoc swPart;
            string sModelName;
            string sPathName;
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                int errors=0;
                int warnings=0;
                swModel = swApp.OpenDoc6(sldprtName,1, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                sModelName = sldprtName;
                sPathName = sldprtName;
                sPathName = sPathName.Substring(0, sPathName.Length - 7);
                if(!Directory.Exists(sPathName))
                Directory.CreateDirectory(sPathName);
                string sPartName=Exceldeal.Getfilename(sModelName,7);
                swPart = (PartDoc)swModel;
                object[] swBodys = (object[])swPart.GetBodies2(-1, false);
                HideBodys(ref swBodys, true);
                foreach (object body in swBodys)
                {
                    Body2 swBody = (Body2)body;
                    string bodyName=swBody.Name;
                    if(bodyName.Contains("镜向")||bodyName.Contains("阵列")||bodyName.Contains("镜像"))
                        continue;
                    swBody.HideBody(false);
                    GetMaxAreaFace(ref swBody);
                    swModel.SketchManager.InsertSketch(true);
                    swModel.SketchManager.InsertSketch(true);
                    //swModel.ClearSelection2(true);
                    GetFaceDwg(ref swPart, sModelName, sPathName+"//"+sPartName+"-"+bodyName+".dwg");
                    swBody.HideBody(true);
                }
                //swApp.CloseAllDocuments(true);
                swApp.CloseDoc(sPartName);
            }
        }
        static void GetFaceDwg(ref PartDoc swPart, string sModelName, string sPathName)
        {
            object varAlignment;
            double[] dataAlignment = new double[12];
            object varViews;
            string[] dataViews = new string[1];
            int options;
            for (int i = 0; i < 12; i++)
                dataAlignment[i] = 0.0;
            varAlignment = dataAlignment;
            dataViews[0] = "*Current";
            varViews = dataViews;
            //Export each annotation view to a separate drawing file
            swPart.ExportToDWG2(sPathName, sModelName, (int)swExportToDWG_e.swExportToDWG_ExportAnnotationViews, true, varAlignment, true, true, 0, varViews);
            //Export sheet metal to a single drawing file
            options = 1;  //include flat-pattern geometry
            swPart.ExportToDWG2(sPathName, sModelName, (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal, true, varAlignment, false, false, options, null);
        }
        public static void HideBodys(ref object[] swBodys, bool status)
        {
            foreach (Body2 body in swBodys)
                body.HideBody(status);
        }
        public static void GetMaxAreaFace(ref Body2 swBody)
        {
            int faceCount = swBody.GetFaceCount();
            Face2 swFace = swBody.IGetFirstFace();
            Face2 maxFace = swFace;
            double temArea = maxFace.GetArea();
            for (int i = 1; i < faceCount; i++)
            {
                swFace = swFace.IGetNextFace();
                if (temArea < swFace.GetArea())
                {
                    maxFace = swFace;
                    temArea = swFace.GetArea();
                }
            }
            Entity swEn = (Entity)maxFace;
            swEn.Select(false);//到這一步就可以選中
            var edges=maxFace.IGetEdges();
            var vertex=edges.IGetStartVertex();
            vertex.IGetPoint();
        }
        public static void ToDWG(string sldprtName)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                double swSheetWidth = 0.4;
                double swSheetHeight = 0.3;
                string programPath = System.Environment.CurrentDirectory;
                var newDoc = swApp.NewDocument(@"C:\ProgramData\SolidWorks\SOLIDWORKS 2018\templates\空白工程图模板.drwdot", 12, swSheetWidth, swSheetHeight);//创建一个空白草图，模板放在根目录
                if (newDoc != null)
                {
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                    DrawingDoc swDrawing = (DrawingDoc)swModel;
                    View swView = (View)swDrawing.CreateDrawViewFromModelView3(sldprtName, "*前视", 0, 0, 0);
                    swModel.ClearSelection2(true);//确认草图
                    ModelDocExtension swModExt = (ModelDocExtension)swModel.Extension;
                    int error = 0;
                    int warnings = 0;
                    //设置dxf 导出版本 R14
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, 2);
                    //是否显示 草图
                    swModel.SetUserPreferenceToggle(196, false);
                    string sDwgFileName = sldprtName.Substring(0, sldprtName.Length - 6) + "dwg";
                    swModExt.SaveAs(sDwgFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref error, ref warnings);
                    swApp.CloseAllDocuments(true);
                }
            }
        }


        public static double Getarea(string sldprtName)
        {
            ISldWorks swApp = Utility.ConnectToSolidWorks();
            if (swApp != null)
            {
                string partDefaultTemplate = swApp.GetDocumentTemplate((int)swDocumentTypes_e.swDocPART, "", 0, 0, 0);
                var newDoc = swApp.NewDocument(partDefaultTemplate, 0, 0, 0);
                if (newDoc != null)
                {
                    ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;
                    bool boolstatus = swModel.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, false, 0, null, 0);

                    var importData = (ImportDxfDwgData)swApp.GetImportFileData(sldprtName);
                    // Unit
                    importData.set_LengthUnit("", (int)swLengthUnit_e.swMM);
                    // Position
                    var bRet = importData.SetPosition("", (int)swDwgImportEntitiesPositioning_e.swDwgEntitiesCentered, 0, 0);
                    // Sheet scale
                    bRet = importData.SetSheetScale("", 1.0, 2.0);
                    // Paper size
                    bRet = importData.SetPaperSize("", (int)swDwgPaperSizes_e.swDwgPaperAsize, 0.0, 0.0);
                    //Import method
                    importData.set_ImportMethod("", (int)swImportDxfDwg_ImportMethod_e.swImportDxfDwg_ImportToExistingPart);
                    // Import file with importData
                    //swFeat = swFeatMgr.InsertDwgOrDxfFile2(sDwgFileName, importData);
                    Feature myFeature = swModel.FeatureManager.InsertDwgOrDxfFile2(sldprtName, importData);
                    //Feature myFeature = swModel.FeatureManager.InsertDwgOrDxfFile(sDwgFileName);

                    swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 1, 1, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, 0, 0, false);

                    swModel = (ModelDoc2)swApp.ActiveDoc;
                    ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
                    swModelDocExt.IncludeMassPropertiesOfHiddenBodies = false;
                    int massStatus = 0;
                    double[] massProperties = (double[])swModelDocExt.GetMassProperties(1, ref massStatus);
                    if ((massProperties != null))
                    {
                        swApp.CloseAllDocuments(true);
                        return massProperties[3];
                    }
                }
            }
            swApp.CloseAllDocuments(true);
            return -1;
        }

        public static void CopyRow(int i, int j, NPOI.SS.UserModel.IRow oldRow,
        ref NPOI.SS.UserModel.IWorkbook newWorkBook, ref NPOI.SS.UserModel.ISheet newSheet)
        {
            IRow newRow = newSheet.CreateRow(j);
            for (int m = 0; m < 26; m++)
            {
                ICell oldCell = oldRow.GetCell(m);
                ICell newCell = newRow.CreateCell(m);
                try
                {
                    switch (oldCell.CellType)
                    {
                        case CellType.Blank:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            newCell.SetCellValue(oldCell.BooleanCellValue);
                            break;
                        case CellType.Error:
                            newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                            break;
                        case CellType.Formula:
                            try
                            {
                                newCell.SetCellValue(oldCell.NumericCellValue);
                            }
                            catch (System.Exception)
                            {

                                newCell.SetCellValue(oldCell.RichStringCellValue);
                            }
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue(oldCell.NumericCellValue);
                            break;
                        case CellType.String:
                            newCell.SetCellValue(oldCell.RichStringCellValue);
                            break;
                        case CellType.Unknown:
                            newCell.SetCellValue(oldCell.StringCellValue);
                            break;
                    }
                }
                catch (System.NullReferenceException)
                {
                    newCell.SetCellValue("");
                    continue;
                }
                var newCellStyle = newWorkBook.CreateCellStyle();
                newCellStyle.CloneStyleFrom(oldCell.CellStyle);
                newCell.CellStyle = newCellStyle;
            }
        }

        public static void AutoColumnWidth(NPOI.SS.UserModel.ISheet newSheet, int cols, NPOI.SS.UserModel.ISheet oldSheet)
        {
            for (int col = 0; col <= cols; col++)
            {
                newSheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
                if (oldSheet.IsColumnHidden(col))
                    newSheet.SetColumnHidden(col, true);
            }
        }
    }
}
