using LixinFastReadExcel07;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace TestProject
{
    
    
    /// <summary>
    ///这是 LixinFastReadExcelTest 的测试类，旨在
    ///包含所有 LixinFastReadExcelTest 单元测试
    ///</summary>
    [TestClass()]
    public class LixinFastReadExcelTest
    {
        string file = @"d:\website\test.xlsx";
        private TestContext testContextInstance;

        /// <summary>
        ///获取或设置测试上下文，上下文提供
        ///有关当前测试运行及其功能的信息。
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region 附加测试特性
        // 
        //编写测试时，还可使用以下特性:
        //
        //使用 ClassInitialize 在运行类中的第一个测试前先运行代码
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //使用 ClassCleanup 在运行完类中的所有测试后再运行代码
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //使用 TestInitialize 在运行每个测试前先运行代码
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //使用 TestCleanup 在运行完每个测试后运行代码
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///LixinFastReadExcel 构造函数 的测试
        ///</summary>
        [TestMethod()]
        public void LixinFastReadExcelConstructorTest()
        {
            LixinFastReadExcel target = new LixinFastReadExcel(file);
            Assert.AreEqual(file, target.filePath);
        }



        /// <summary>
        ///getSheetName 的测试
        ///</summary>
        [TestMethod()]
        public void getSheetNameTest()
        {

            LixinFastReadExcel target = new LixinFastReadExcel(file); // TODO: 初始化为适当的值
            string[] expected = new string[] { "Sheet1", "Sheet2", "9月份和10月份办理二期(12月份的存量GPRS激励酬金" }; // TODO: 初始化为适当的值
            string[] actual;
            actual = target.getSheetName();
            Assert.AreEqual(expected.Length, actual.Length);
            for (int i = 0; i < actual.Length; i++)
            {
                Assert.AreEqual(expected[i], actual[i]);
            }
        }

        /// <summary>
        ///letter2Num 的测试
        ///</summary>
        [TestMethod()]
        public void letter2NumTest()
        {
            LixinFastReadExcel target = new LixinFastReadExcel(file); // TODO: 初始化为适当的值
            string letter ="AA"; // TODO: 初始化为适当的值
            int expected = 27; // TODO: 初始化为适当的值
            int actual;
            actual = target.letter2Num(letter);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///loadShareString 的测试
        ///</summary>
        [TestMethod()]
        public void loadShareStringTest()
        {
            LixinFastReadExcel target = new LixinFastReadExcel(file); // TODO: 初始化为适当的值
            target.loadShareString();
            Assert.AreEqual(target.myShareString[0], "渠道标识");
            Assert.AreEqual(target.myShareString[target.myShareString.Length-1], "VEJY405HC");
        }
    }
}
