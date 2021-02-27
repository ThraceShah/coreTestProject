using System;
using System.Collections.Generic;
using System.IO;

namespace Filedeal
{
    public class Fileusing
    {
        public List<String> list = new List<String>();

        public void Director(string dirs)
        {
            //绑定到指定的文件夹目录
            DirectoryInfo dir = new DirectoryInfo(dirs);
            //检索表示当前目录的文件和子目录
            FileSystemInfo[] fsinfos = dir.GetFileSystemInfos();
            //遍历检索的文件和子目录
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                //判断是否为空文件夹　　
                if (fsinfo is DirectoryInfo)
                {
                    //递归调用
                    Director(fsinfo.FullName);
                }
                else
                {
                    Console.WriteLine(fsinfo.FullName);
                    //将得到的文件全路径放入到集合中
                    list.Add(fsinfo.FullName);
                }
            }
        }
    }

    public class Exceldeal
    {
        public static double GetTh(string d)
        {
            double t = 0;
            int temp = 0;
            string[] tem;
            tem = d.Split('t', '*', 'x');
            temp = int.Parse(tem[1]);
            t = temp / 1000.0;
            return t;
        }
        public static string GetPath(string pathfile)
        {
            string[] tem;
            string temp = "";
            tem = pathfile.Split('\\');
            foreach (string str in tem)
                temp = str;
            pathfile = pathfile.Substring(0, pathfile.Length - temp.Length);
            return pathfile;
        }

        public static string Getfilename(string pathfile,int len)
        {
            string[] tem;
            string temp = "";
            tem = pathfile.Split('\\');
            foreach (string str in tem)
                temp = str;
            string filename = temp.Substring(0, temp.Length - len);
            return filename;
        }

        internal static double[] GetDimen(string[,] bomExcel, int i)
        {
            double[] size = new double[] { 0, 0, 0 };
            string[] tem;
            tem = bomExcel[1, i].Split('t', '*', 'x');
            size[0] = int.Parse(tem[2]) / 1000.0;
            size[1] = int.Parse(tem[3]) / 1000.0;
            return size;
        }
    }
}