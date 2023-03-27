using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using PM02_B7;

namespace UnitTestProject_B7
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string width = "0";
            string height = "0";
            string expected = "null";
            string actual = PM02_B7.Form1.ActiveForm(width, height);
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod_calculation_meter2()
        {
            string width = "";
            string height = "";
            string expected = "error";
            string actual = PM02_B7.Form1.calculation_meter(width, height);
            Assert.AreEqual(expected, actual);
        }
        [TestMethod]
        public void TestMethod_calculation_meter3()
        {
            string width = "-10";
            string height = "-5";
            string expected = "null";
            string actual = PM02_B7.Form1.calculation_meter(width, height);
            Assert.AreEqual(expected, actual);
        }
    }
}
