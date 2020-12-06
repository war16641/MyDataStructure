using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyDataStructure;
using System;
using System.Collections.Generic;
using System.Text;

namespace MyDataStructure.Tests
{
    [TestClass()]
    public class FlatDataModelTests
    {


        [TestMethod()]
        public void load_from_excelTest()
        {
            FlatDataModel fdm = FlatDataModel.load_from_excel(Environment.CurrentDirectory + "\\测试需要\\工作簿1.xlsx");
            //string str2 = Environment.CurrentDirectory;//获取和设置当前目录（即该进程从中启动的目录）的完全限定路径
            //System.Diagnostics.Debug.Print(str2);
            //System.Diagnostics.Debug.Print(System.IO.Directory.GetCurrentDirectory());
            Assert.AreEqual(300.0, (double)fdm.units[1].data["年龄"], 0.001);

        }

        [TestMethod()]
        public void find_oneTest()
        {
            FlatDataModel fdm = FlatDataModel.load_from_excel(Environment.CurrentDirectory + "\\测试需要\\工作簿1.xlsx");
            int iint;
            DataUnit u = fdm.find_one(delegate (DataUnit u)
            {
                if ((double)u.data["年龄"] == 20.0)
                {
                    return true;
                }
                return false;
            },
            out iint);
            Assert.AreEqual("爱丽丝", u.data["姓名"].ToString());
            Assert.AreEqual(2, iint);
        }

        [TestMethod()]
        public void sortTest()
        {
            FlatDataModel fdm = FlatDataModel.load_from_excel(
                Environment.CurrentDirectory + "\\测试需要\\工作簿1.xlsx");
            fdm.sort(delegate (DataUnit a, DataUnit b)
            {
                if ((double)a.data["年龄"] > (double)b.data["年龄"]) return 1;
                return -1;
            });
            Assert.AreEqual("奥特曼", fdm.units[2].data["姓名"]);


            fdm = FlatDataModel.load_from_excel(
                Environment.CurrentDirectory + "\\测试需要\\flhz.xlsx");
            List<string> lst = new List<string>();
            lst.Add("身高");
            lst.Add("年龄");
            fdm.sort(lst);
            //fdm.show_in_excel();
            Assert.AreEqual("普京", fdm.units[fdm.Count - 2].data["姓名"]);

        }

        //[TestMethod()]
        //public void make_bunchTest()
        //{
        //    FlatDataModel fdm = FlatDataModel.load_from_excel(
        //        Environment.CurrentDirectory + "\\测试需要\\flhz.xlsx");
        //    string a = "sd";
        //    double s = 12.2;
        //}

        [TestMethod()]
        public void avgTest()
        {
            List<object> lst = new List<object>();
            lst.Add(1.0);
            lst.Add(2.0);
            lst.Add(3.0);
            Assert.AreEqual(2.0, (double)MyStatistic.avg(lst));
            Assert.AreEqual(1.0, (double)MyStatistic.min(lst));
        }

        [TestMethod()]
        public void flhzTest()
        {
            FlatDataModel fdm = FlatDataModel.load_from_excel(
                Environment.CurrentDirectory + "\\测试需要\\flhz.xlsx");

            FLHZ_OPERATION f1 = new FLHZ_OPERATION();
            f1.fieldname = "身高";
            f1.func = MyStatistic.sum;
            FLHZ_OPERATION f2 = new FLHZ_OPERATION();
            f2.fieldname = "身高";
            f2.newname = "平均身高";
            f2.func = MyStatistic.avg;
            FLHZ_OPERATION f3 = new FLHZ_OPERATION();
            f3.fieldname = "宣言";
            f3.func = MyStatistic.sum;
            var lo = new List<string>() { "性别", "种族" };
            FlatDataModel rt = fdm.flhz(lo, f1, f2, f3);
            //rt.show_in_excel();
            Assert.AreEqual(4, rt.Count);
            Assert.AreEqual(174.0, rt.units[0].data["平均身高"]);
            //Assert.Fail();
        }

        [TestMethod()]
        public void add_field_from_other_modelTest()
        {
            FlatDataModel fdm = FlatDataModel.load_from_excel(
                Environment.CurrentDirectory + "\\测试需要\\flhz.xlsx");
            FlatDataModel fdmo = FlatDataModel.load_from_excel(
                Environment.CurrentDirectory + "\\测试需要\\添加字段.xlsx");
            string[] s = new string[] { "声望" ,"归属地"};
            fdm.add_field_from_other_model("姓名",s, fdmo, "sd");
            //fdm.show_in_excel();
            
            Assert.AreEqual(1000.0,(double)fdm.units[0].data["声望"],0.001);
            Assert.AreEqual("美国", (string)fdm.units[2].data["归属地"]);
            Assert.AreEqual("sd", (string)fdm.units[fdm.Count-1].data["归属地"]);
        }
    }
}