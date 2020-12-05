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
            FlatDataModel fdm = FlatDataModel.load_from_excel(Environment.CurrentDirectory+"\\测试需要\\工作簿1.xlsx");
            string str2 = Environment.CurrentDirectory;//获取和设置当前目录（即该进程从中启动的目录）的完全限定路径
            System.Diagnostics.Debug.Print(str2);
            System.Diagnostics.Debug.Print(System.IO.Directory.GetCurrentDirectory());
            Assert.Fail();
        }


    }
}