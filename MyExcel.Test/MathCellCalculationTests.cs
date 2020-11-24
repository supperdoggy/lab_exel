using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MyExcel.Test
{
    [TestClass]
    public class MathCellCalculationTests
    {
        [TestMethod]
        public void Evaluate_10plus20_30return()
        {
            // arrange
            const double x = 10;
            const double y = 20;
            string expr = $"={x} + {y}";
            const double expected = 30;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Evaluate_10minus20_Neg10return()
        {
            // arrange
            const double x = 10;
            const double y = 20;
            string expr = $"={x} - {y}";
            const double expected = -10;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Evaluate_10multiply20_200return()
        {
            // arrange
            const double x = 10;
            const double y = 20;
            string expr = $"={x} * {y}";
            const double expected = 200;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Div_IntAndInt_DoubleReturn()
        {
            // arrange
            const int x = 3;
            const int y = 2;
            string expr = $"={x} / {y}";
            const double expected = 1.5;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void SetUnary_plus10_10return()
        {
            // arrange
            const double x = 10;
            string expr = $"=+{x}";
            const double expected = 10;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void SetUnary_minus10_Neg10return()
        {
            // arrange
            const double x = 10;
            string expr = $"=-{x}";
            const double expected = -10;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Power_3and3_27return()
        {
            // arrange
            const double x = 3;
            const double y = 3;

            string expr = $"={x} ^ {y}";
            const double expected = 27;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Power_DoubleAndDouble_DoubleReturn()
        {
            // arrange
            const double x = 3.5;
            const double y = 2.2;

            string expr = $"={x.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)} " +
                          $"^ {y.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)}";
            const double expected = 15.7380056747621;

            // act
            object result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Max_7and3_7Return()
        {
            // arrange
            const double x = 7;
            const double y = 3;

            string expr = $"=max({x}, {y})";
            const double expected = 7;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Max_DoubleAndDouble_DoubleReturn()
        {
            // arrange
            const double x = 7.7778;
            const double y = 7.7777;

            string expr = $"=max({x.ToString("0.0000", System.Globalization.CultureInfo.InvariantCulture)}" +
                          $", {y.ToString("0.0000", System.Globalization.CultureInfo.InvariantCulture)})";
            const double expected = 7.7778;

            // act
            object result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Div_TwoDivSigns_InvalidOperationReturn()
        {
            // arrange
            const double x = 7.7778;
            const double y = 7.7777;

            string expr = $"={x} // {y}";
            const string expected = "InvalidOperation";

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Div_1and0_DivByZeroReturn()
        {
            // arrange
            const string expr = "=1/0";
            const string expected = "DivByZero";

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void Div_DoubleAnd0_DivByZeroReturn()
        {
            // arrange
            const double x = 2.55215211245215;
            string expr = $"={x.ToString("0.00000000000000", System.Globalization.CultureInfo.InvariantCulture)} " +
                          "/ 0";
            const string expected = "DivByZero";

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void FalseMethod_InvalidOperationReturn()
        {
            // arrange
            const string expr = "=son(21)";
            const string expected = "InvalidOperation";

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void DoubleMinus_minusMinus10_10Return()
        {
            // arrange
            const string expr = "=--10";
            const double expected = 10;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        public void JoinedSigns_plusMinus10_Neg10Return()
        {
            // arrange
            const string expr = "=+-10";
            const double expected = -10;

            // act
            var result = Calculate(expr);

            // assert
            Assert.AreEqual(expected, result);
        }

        private static object Calculate(string expr)
        {
            var testCell = new MathCell(expr);
            testCell.EvaluateFormula();

            return testCell.Value;
        }
    }
}
