using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Connect;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Filedeal;
using Microsoft.Win32;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SolidworksModel;

namespace TestProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public string[,] publicBomExcel = new string[3, 500];
        public string bomFilePath = "";//BOM的路径
        public string bomFilefolderPath = "";//BOM所在文件夹的路径
        int rowsint;
        private void ImportexcelBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Files|*.xls;*.xlsx";              // 设定打开的文件类型
            //openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;                       // 打开app对应的路径
            openFileDialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);  // 打开桌面

            // 如果选定了文件
            if (openFileDialog.ShowDialog() == true)
            {
                // 取得文件路径及文件名
                bomFilePath = openFileDialog.FileName;
                bomFilefolderPath = Exceldeal.GetPath(bomFilePath);//获取文件夹路径
                //MessageBox.Show(filePaths);
                IWorkbook workBook;
                using (FileStream file = new FileStream(bomFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (bomFilePath.Substring(bomFilePath.Length - 1, 1) == "s")
                        workBook = new HSSFWorkbook(file);
                    else
                        workBook = new XSSFWorkbook(file);
                    file.Close();
                }
                NPOI.SS.UserModel.ISheet xTest = workBook.GetSheet("Sheet1");
                rowsint = xTest.LastRowNum;
                if (xTest.GetRow(0).LastCellNum > 1)
                {

                    for (int i = 0; i <= rowsint; i++)
                    {
                        IRow tempRow = xTest.GetRow(i);
                        ICell cellLeft = tempRow.GetCell(0);
                        ICell cellRight = tempRow.GetCell(1);
                        publicBomExcel[0, i] = cellLeft.ToString();
                        publicBomExcel[1, i] = cellRight.ToString();
                        //thickList.Add(cellLeft.ToString(), cellRight.ToString());
                    }
                }//把BOM表的第一列和第二列字符串数组中，方便查找
            }
        }

        private void StretchBtn_Click(object sender, RoutedEventArgs e)
        {
            if (publicBomExcel[0, 0] == null)
            {
                MessageBox.Show("请导入BOM表");
                ImportexcelBtn_Click(sender, e);
            }
            if (publicBomExcel[0, 0] == null)//要是实在不看提示我也没办法了
                return;
            for (int i = 0; i <= rowsint; i++)
            {
                double t = Exceldeal.GetTh(publicBomExcel[1, i]);
                Solution.IputTD(publicBomExcel[0, i], bomFilefolderPath, t, 'F');
            }
            MessageBox.Show("运行结束");
        }

        private void StretchBtn_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (publicBomExcel[0, 0] == null)
            {
                MessageBox.Show("请导入BOM表");
                ImportexcelBtn_Click(sender, e);
            }
            if (publicBomExcel[0, 0] == null)//要是实在不看提示我也没办法了
                return;
            for (int i = 0; i <= rowsint; i++)
            {
                double t = Exceldeal.GetTh(publicBomExcel[1, i]);
                Solution.IputTD(publicBomExcel[0, i], bomFilefolderPath, t, 'G');
            }
            MessageBox.Show("运行结束");
        }

        private void Gen_Reg_Click(object sender, RoutedEventArgs e)
        {
            if (publicBomExcel[0, 0] == null)
            {
                MessageBox.Show("请导入BOM表");
                ImportexcelBtn_Click(sender, e);
            }
            if (publicBomExcel[0, 0] == null)//要是实在不看提示我也没办法了
                return;
            for (int i = 0; i <= rowsint; i++)
            {
                double t = Exceldeal.GetTh(publicBomExcel[1, i]);//获取板厚
                double[] dimensions = new double[3] { 0, 0, 0 };
                dimensions = Exceldeal.GetDimen(publicBomExcel, i);//获取尺寸和外形信息
                Solution.ToRectangle(dimensions, bomFilefolderPath, t, publicBomExcel[0, i]);
            }
            MessageBox.Show("运行结束");
        }

        private void outDWG_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            // Set validate names and check file exists to false otherwise windows will
            // not let you select "Folder Selection."
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            // Always default to Folder Selection.
            folderBrowser.FileName = "Folder Selection.";
            if (folderBrowser.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                Fileusing Filelist = new Fileusing();
                Filelist.Director(folderPath);
                foreach (String sldprtName in Filelist.list)
                {
                    int start = sldprtName.Length - 6;
                    //MessageBox.Show(sldprtName.Substring(start, 6));
                    if (sldprtName.Substring(start, 6) == "sldprt" || sldprtName.Substring(start, 6) == "SLDPRT")
                        Solution.ToDWG(sldprtName);
                }
            }
            MessageBox.Show("运行结束");
        }

        private void GetArea_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderBrowser = new OpenFileDialog();
            folderBrowser.ValidateNames = false;
            folderBrowser.CheckFileExists = false;
            folderBrowser.CheckPathExists = true;
            folderBrowser.FileName = "Folder Selection.";
            IWorkbook workBook;
            workBook = new HSSFWorkbook();
            var sheet1 = workBook.CreateSheet("Sheet1");
            //var itemRow=sheet1.CreateRow(0);
            //var areaRow=sheet1.CreateRow(1);
            if (folderBrowser.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(folderBrowser.FileName);
                Fileusing Filelist = new Fileusing();
                Filelist.Director(folderPath);
                int i = 0;
                foreach (String sldprtName in Filelist.list)
                {
                    var row = sheet1.CreateRow(i);
                    int start = sldprtName.Length - 3;
                    //MessageBox.Show(sldprtName.Substring(start, 6));
                    if (sldprtName.Substring(start, 3) == "dwg" || sldprtName.Substring(start, 3) == "DWG" || sldprtName.Substring(start, 3) == "dxf" || sldprtName.Substring(start, 3) == "DXF")
                    {
                        row.CreateCell(1).SetCellValue(Solution.Getarea(sldprtName) * 1000000);
                        row.CreateCell(0).SetCellValue(Exceldeal.Getfilename(sldprtName));
                        i++;
                    }
                }
                FileStream stream = File.OpenWrite(folderPath + "\\面积统计.xls"); ;
                workBook.Write(stream);
                stream.Close();
            }
            MessageBox.Show("面积计算完成");
        }

        private void Split_fig_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Files|*.xls;*.xlsx";              // 设定打开的文件类型
            openFileDialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);  // 打开桌面
            if (openFileDialog.ShowDialog() == true)
            {
                // 取得文件路径及文件名
                string filePath2 = openFileDialog.FileName;//获取包含文件名在内的文件路径
                string filePath3 = openFileDialog.SafeFileName;//仅获取当前文件文件名
                string[] filePathm = filePath3.Split('.');
                bomFilefolderPath = Exceldeal.GetPath(filePath2);//获取文件夹路径
                IWorkbook oriworkBook;
                using (FileStream file = new FileStream(filePath2, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (filePath2.Substring(filePath2.Length - 1, 1) == "s")
                        oriworkBook = new HSSFWorkbook(file);
                    else
                        oriworkBook = new XSSFWorkbook(file);
                    file.Close();
                }
                NPOI.SS.UserModel.ISheet oriSheet = oriworkBook.GetSheet("Sheet1");
                int cellcount = oriSheet.GetRow(0).LastCellNum;
                rowsint = 0;
                try
                {
                    for(;true;rowsint++)
                    oriSheet.GetRow(rowsint).GetCell(5).ToString();
                }
                catch (System.Exception)
                {}
                int i = 0;
                while (i+1<rowsint&&oriSheet.GetRow(i + 1).GetCell(5).ToString()!="")
                {
                    int j = 0;
                    IWorkbook newWorkBook = new HSSFWorkbook();
                    NPOI.SS.UserModel.ISheet newSheet = newWorkBook.CreateSheet("Sheet1");
                    Solution.CopyRow(0,0,oriSheet.GetRow(0),ref newWorkBook,ref newSheet);
                    do
                    {
                        j++; i++;
                        Solution.CopyRow(i,j,oriSheet.GetRow(i),ref newWorkBook,ref newSheet);
                    } while (i+1<rowsint&&oriSheet.GetRow(i).GetCell(5).ToString() ==
                    oriSheet.GetRow(i + 1).GetCell(5).ToString());
                    string path=bomFilefolderPath +oriSheet.GetRow(i).GetCell(1);
                    if (false == System.IO.Directory.Exists(path))
                    {
                        //创建文件夹
                        Directory.CreateDirectory(path);
                    }
                    Solution.AutoColumnWidth(newSheet,25,oriSheet);
                    SaveWorkbook(newWorkBook,path+"\\"+ oriSheet.GetRow(i).GetCell(5) + ".xls" );
                }
                MessageBox.Show("拆分完成！");
            }
        }
        void SaveWorkbook(IWorkbook workbook, string path)
        {
            using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
        }

        private void DealexcelBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Files|*.xls;*.xlsx";              // 设定打开的文件类型
            //openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;                       // 打开app对应的路径
            openFileDialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);  // 打开桌面

            // 如果选定了文件
            if (openFileDialog.ShowDialog() == true)
            {
                // 取得文件路径及文件名
                string filePath2 = openFileDialog.FileName;//获取包含文件名在内的文件路径
                string filePath3 = openFileDialog.SafeFileName;//仅获取当前文件文件名
                string[] filePathm = filePath3.Split('.');
                bomFilefolderPath = Exceldeal.GetPath(filePath2);//获取文件夹路径
                IWorkbook workBook;
                using (FileStream file = new FileStream(filePath2, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (filePath2.Substring(filePath2.Length - 1, 1) == "s")
                        workBook = new HSSFWorkbook(file);
                    else
                        workBook = new XSSFWorkbook(file);
                    file.Close();
                }
                NPOI.SS.UserModel.ISheet xTest = workBook.GetSheet("Sheet1");
                rowsint = xTest.LastRowNum;
                NPOI.SS.UserModel.ISheet bomSheet;
                if (workBook.GetSheet("BOM") == null)
                {
                    bomSheet = workBook.CreateSheet("BOM");
                }
                else
                {
                    bomSheet = workBook.GetSheetAt(2);
                }
                if (xTest.GetRow(0).LastCellNum > 1)
                {
                    IRow row1 = bomSheet.CreateRow(0);
                    row1.CreateCell(0).SetCellValue("件号");
                    row1.CreateCell(1).SetCellValue("数量");
                    row1.CreateCell(2).SetCellValue("类型");
                    row1.CreateCell(3).SetCellValue("规格");
                    row1.CreateCell(4).SetCellValue("材质");
                    row1.CreateCell(5).SetCellValue("重量");
                    row1.CreateCell(6).SetCellValue("备注");
                    row1.CreateCell(7).SetCellValue("备注2");
                    row1.CreateCell(8).SetCellValue("规格2");
                    row1.CreateCell(9).SetCellValue("规格3");
                    for (int i = 0; i <= rowsint / 3; i++)
                    {
                        IRow row = bomSheet.CreateRow(i + 1);//index代表多少行
                        row.HeightInPoints = 15;//行高
                        row.CreateCell(0).SetCellValue(xTest.GetRow(3 * i).GetCell(0).ToString());//获得件号
                        row.CreateCell(1).SetCellValue(xTest.GetRow(3 * i).GetCell(2).ToString());//获得数量
                        string[] type = xTest.GetRow(3 * i + 2).GetCell(3).ToString().Replace("PLATE", "钢板").Replace("PIPE", "钢管").Replace("ROUND", "圆钢").Split(' ');
                        string type2 = "";
                        if (type[0][0] != 'W')
                        {
                            row.CreateCell(2).SetCellValue(type[0]);//获得类型
                            if (type.Length > 1)
                            {
                                for (int k = 1; k < type.Length; k++)
                                    type2 = type2 + type[k];
                            }
                        }
                        else
                        {
                            row.CreateCell(2).SetCellValue(type[0] + " " + type[1]);//获得类型
                            if (type.Length > 2)
                            {
                                for (int k = 2; k < type.Length; k++)
                                    type2 = type2 + type[k];
                            }
                        }
                        row.CreateCell(8).SetCellValue(type2);//获得规格2
                        if (row.GetCell(2).ToString() == "钢板")
                            row.CreateCell(3).SetCellValue(xTest.GetRow(3 * i + 1).GetCell(3).ToString().Replace('x', '*').Replace(',', '.').Replace("...", "-").Replace(" ", string.Empty).Insert(0, "t"));//获得规格
                        else
                            row.CreateCell(3).SetCellValue(xTest.GetRow(3 * i + 1).GetCell(3).ToString().Replace('x', '*').Replace(',', '.').Replace("...", "-").Replace(" ", string.Empty));//获得规格
                        row.CreateCell(9).SetCellValue(row.GetCell(3).ToString() + " " + type2);
                        string[] str = xTest.GetRow(3 * i).GetCell(5).ToString().Split(' ', '(', ')');
                        string marks = "";
                        string mark2 = "";
                        for (int k = 0; k < str.Length; k++)
                        {
                            if (str[k].Length >= 2)
                                mark2 = str[k].Substring(0, 2);
                            if (mark2 == "DI" || mark2 == "EN" || mark2 == "IS" || mark2 == "GB")
                            {
                                for (int j = k; j < str.Length; j++)
                                    marks = marks + str[j] + " ";
                                break;
                            }
                        }
                        row.CreateCell(4).SetCellValue(str[0]);//获得材质
                        row.CreateCell(6).SetCellValue(xTest.GetRow(3 * i).GetCell(6).ToString());//获得备注
                        row.CreateCell(7).SetCellValue(marks);//获得备注2
                    }
                }
                FileStream stream = File.OpenWrite(bomFilefolderPath + filePathm[filePathm.Length - 2] + "已转化." + filePathm[filePathm.Length - 1]);
                workBook.Write(stream);
                stream.Close();
                MessageBox.Show("清单转化完成");
            }
        }


        private void btn_GetMass_Click(object sender, EventArgs e)
        {
            // 获取质量属性可参考 Get Mass Properties of Visible and Hidden Components Example (C#)

            ISldWorks swApp = Utility.ConnectToSolidWorks();

            ModelDoc2 swModel = (ModelDoc2)swApp.ActiveDoc;

            ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;

            swModelDocExt.IncludeMassPropertiesOfHiddenBodies = false;
            int massStatus = 0;

            double[] massProperties = (double[])swModelDocExt.GetMassProperties(1, ref massStatus);
            if ((massProperties != null))
            {
                MessageBox.Show(" CenterOfMassX = " + massProperties[0]);
                MessageBox.Show(" CenterOfMassY = " + massProperties[1]);
                MessageBox.Show(" CenterOfMassZ = " + massProperties[2]);
                MessageBox.Show(" Volume = " + massProperties[3]);
                MessageBox.Show(" Area = " + massProperties[4]);
                MessageBox.Show(" Mass = " + massProperties[5]);
                MessageBox.Show(" MomXX = " + massProperties[6]);
                MessageBox.Show(" MomYY = " + massProperties[7]);
                MessageBox.Show(" MomZZ = " + massProperties[8]);
                MessageBox.Show(" MomXY = " + massProperties[9]);
                MessageBox.Show(" MomZX = " + massProperties[10]);
                MessageBox.Show(" MomYZ = " + massProperties[11]);
            }
            MessageBox.Show("-------------------------------");
        }

    }
}
