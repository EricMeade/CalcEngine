using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CalcEngine.Functions
{
    internal static class MathTrig
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ABS", 1, Abs);
            ce.RegisterFunction("ACOS", 1, Acos);
            //ce.RegisterFunction("ACOSH", Acosh, 1);
            ce.RegisterFunction("ASIN", 1, Asin);
            //ce.RegisterFunction("ASINH", Asinh, 1);
            ce.RegisterFunction("ATAN", 1, Atan);
            ce.RegisterFunction("ATAN2", 2, Atan2);
            //ce.RegisterFunction("ATANH", Atanh, 1);
            ce.RegisterFunction("CEILING", 1, Ceiling);
            //ce.RegisterFunction("COMBIN", Combin, 1);
            ce.RegisterFunction("COS", 1, Cos);
            ce.RegisterFunction("COSH", 1, Cosh);
            //ce.RegisterFunction("DEGREES", Degrees, 1);
            //ce.RegisterFunction("EVEN", Even, 1);
            ce.RegisterFunction("EXP", 1, Exp);
            //ce.RegisterFunction("FACT", Fact, 1);
            //ce.RegisterFunction("FACTDOUBLE", FactDouble, 1);
            ce.RegisterFunction("FLOOR", 1, Floor);
            //ce.RegisterFunction("GCD", Gcd, 1);
            ce.RegisterFunction("INT", 1, Int);
            //ce.RegisterFunction("LCM", Lcm, 1);
            ce.RegisterFunction("LN", 1, Ln);
            ce.RegisterFunction("LOG", 1, 2, Log);
            ce.RegisterFunction("LOG10", 1, Log10);
            //ce.RegisterFunction("MDETERM", MDeterm, 1);
            //ce.RegisterFunction("MINVERSE", MInverse, 1);
            //ce.RegisterFunction("MMULT", MMult, 1);
            //ce.RegisterFunction("MOD", Mod, 2);
            //ce.RegisterFunction("MROUND", MRound, 1);
            //ce.RegisterFunction("MULTINOMIAL", Multinomial, 1);
            ce.RegisterFunction("NULLVALUE", 0, NullValue);
            //ce.RegisterFunction("ODD", Odd, 1);
            ce.RegisterFunction("PI", 0, Pi);
            ce.RegisterFunction("POWER", 2, Power);
            //ce.RegisterFunction("PRODUCT", Product, 1);
            //ce.RegisterFunction("QUOTIENT", Quotient, 1);
            //ce.RegisterFunction("RADIANS", Radians, 1);
            ce.RegisterFunction("RAND", 0, Rand);
            ce.RegisterFunction("RANDBETWEEN", 2, RandBetween);
            //ce.RegisterFunction("ROMAN", Roman, 1);
            ce.RegisterFunction("ROUND", 1, Round);
            //ce.RegisterFunction("ROUNDDOWN", RoundDown, 1);
            //ce.RegisterFunction("ROUNDUP", RoundUp, 1);
            //ce.RegisterFunction("SERIESSUM", SeriesSum, 1);
            ce.RegisterFunction("SIGN", 1, Sign);
            ce.RegisterFunction("SIN", 1, Sin);
            ce.RegisterFunction("SINH", 1, Sinh);
            ce.RegisterFunction("SQRT", 1, Sqrt);
            //ce.RegisterFunction("SQRTPI", SqrtPi, 1);
            //ce.RegisterFunction("SUBTOTAL", Subtotal, 1);
            ce.RegisterFunction("SUM", 1, int.MaxValue, Sum);
            ce.RegisterFunction("SUMEXPRIF", 3, 3, SumExprIf);
            ce.RegisterFunction("SUMIF", 2, 3, SumIf);
            //ce.RegisterFunction("SUMPRODUCT", SumProduct, 1);
            //ce.RegisterFunction("SUMSQ", SumSq, 1);
            //ce.RegisterFunction("SUMX2MY2", SumX2MY2, 1);
            //ce.RegisterFunction("SUMX2PY2", SumX2PY2, 1);
            //ce.RegisterFunction("SUMXMY2", SumXMY2, 1);
            ce.RegisterFunction("TAN", 1, Tan);
            ce.RegisterFunction("TANH", 1, Tanh);
            ce.RegisterFunction("TRUNC", 1, Trunc);
        }

#if DEBUG

        public static void Test(CalcEngine ce)
        {
            ce.Test("ABS(-12)", 12.0);
            ce.Test("ABS(+12)", 12.0);
            ce.Test("ACOS(.23)", System.Math.Acos(.23));
            ce.Test("ASIN(.23)", System.Math.Asin(.23));
            ce.Test("ATAN(.23)", System.Math.Atan(.23));
            ce.Test("ATAN2(1,2)", System.Math.Atan2(1, 2));
            ce.Test("CEILING(1.8)", System.Math.Ceiling(1.8));
            ce.Test("COS(1.23)", System.Math.Cos(1.23));
            ce.Test("COSH(1.23)", System.Math.Cosh(1.23));
            ce.Test("EXP(1)", System.Math.Exp(1));
            ce.Test("FLOOR(1.8)", System.Math.Floor(1.8));
            ce.Test("INT(1.8)", 1);
            ce.Test("LOG(1.8)", System.Math.Log(1.8, 10)); // default base is 10
            ce.Test("LOG(1.8, 4)", System.Math.Log(1.8, 4)); // custom base
            ce.Test("LN(1.8)", System.Math.Log(1.8)); // real log
            ce.Test("LOG10(1.8)", System.Math.Log10(1.8)); // same as Log(1.8)
            ce.Test("NULLVALUE()", null);
            ce.Test("PI()", System.Math.PI);
            ce.Test("POWER(2,4)", System.Math.Pow(2, 4));
            //ce.Test("RAND") <= 1.0);
            //ce.Test("RANDBETWEEN(4,5)") <= 5);
            ce.Test("SIGN(-5)", -1);
            ce.Test("SIGN(+5)", +1);
            ce.Test("SIGN(0)", 0);
            ce.Test("SIN(1.23)", System.Math.Sin(1.23));
            ce.Test("SINH(1.23)", System.Math.Sinh(1.23));
            ce.Test("SQRT(144)", System.Math.Sqrt(144));
            ce.Test("SUM(1, 2, 3, 4)", 1 + 2 + 3 + 4.0);
            ce.Test("TAN(1.23)", System.Math.Tan(1.23));
            ce.Test("TANH(1.23)", System.Math.Tanh(1.23));
            ce.Test("TRUNC(1.23)", 1.0);
            ce.Test("PI()", System.Math.PI);
            ce.Test("PI", System.Math.PI);
            ce.Test("LN(10)", System.Math.Log(10));
            ce.Test("LOG(10)", System.Math.Log10(10));
            ce.Test("EXP(10)", System.Math.Exp(10));
            ce.Test("SIN(PI()/4)", System.Math.Sin(System.Math.PI / 4));
            ce.Test("ASIN(PI()/4)", System.Math.Asin(System.Math.PI / 4));
            ce.Test("SINH(PI()/4)", System.Math.Sinh(System.Math.PI / 4));
            ce.Test("COS(PI()/4)", System.Math.Cos(System.Math.PI / 4));
            ce.Test("ACOS(PI()/4)", System.Math.Acos(System.Math.PI / 4));
            ce.Test("COSH(PI()/4)", System.Math.Cosh(System.Math.PI / 4));
            ce.Test("TAN(PI()/4)", System.Math.Tan(System.Math.PI / 4));
            ce.Test("ATAN(PI()/4)", System.Math.Atan(System.Math.PI / 4));
            ce.Test("ATAN2(1,2)", System.Math.Atan2(1, 2));
            ce.Test("TANH(PI()/4)", System.Math.Tanh(System.Math.PI / 4));
        }

#endif

        private static object Abs(List<Expression> p)
        {
            return System.Math.Abs((double)p[0]);
        }

        private static object Acos(List<Expression> p)
        {
            return System.Math.Acos((double)p[0]);
        }

        private static object Asin(List<Expression> p)
        {
            return System.Math.Asin((double)p[0]);
        }

        private static object Atan(List<Expression> p)
        {
            return System.Math.Atan((double)p[0]);
        }

        private static object Atan2(List<Expression> p)
        {
            return System.Math.Atan2((double)p[0], (double)p[1]);
        }

        private static object Ceiling(List<Expression> p)
        {
            return System.Math.Ceiling((double)p[0]);
        }

        private static object Cos(List<Expression> p)
        {
            return System.Math.Cos((double)p[0]);
        }

        private static object Cosh(List<Expression> p)
        {
            return System.Math.Cosh((double)p[0]);
        }

        private static object Exp(List<Expression> p)
        {
            return System.Math.Exp((double)p[0]);
        }

        private static object Floor(List<Expression> p)
        {
            return System.Math.Floor((double)p[0]);
        }

        private static object Int(List<Expression> p)
        {
            return (int)((double)p[0]);
        }

        private static object Ln(List<Expression> p)
        {
            return System.Math.Log((double)p[0]);
        }

        private static object Log(List<Expression> p)
        {
            var lbase = p.Count > 1 ? (double)p[1] : 10;
            return System.Math.Log((double)p[0], lbase);
        }

        private static object Log10(List<Expression> p)
        {
            return System.Math.Log10((double)p[0]);
        }

        private static object NullValue(List<Expression> p)
        {
            return null;
        }

        private static object Pi(List<Expression> p)
        {
            return System.Math.PI;
        }

        private static object Power(List<Expression> p)
        {
            return System.Math.Pow((double)p[0], (double)p[1]);
        }

        private static Random _rnd = new Random();

        private static object Rand(List<Expression> p)
        {
            return _rnd.NextDouble();
        }

        private static object RandBetween(List<Expression> p)
        {
            return _rnd.Next((int)(double)p[0], (int)(double)p[1]);
        }

        private static object Round(List<Expression> p)
        {
            return System.Math.Round((double)p[0]);
        }

        private static object Sign(List<Expression> p)
        {
            return System.Math.Sign((double)p[0]);
        }

        private static object Sin(List<Expression> p)
        {
            return System.Math.Sin((double)p[0]);
        }

        private static object Sinh(List<Expression> p)
        {
            return System.Math.Sinh((double)p[0]);
        }

        private static object Sqrt(List<Expression> p)
        {
            return System.Math.Sqrt((double)p[0]);
        }

        private static object Sum(List<Expression> p)
        {
            var tally = new Tally();
            foreach (Expression e in p)
            {
                tally.Add(e);
            }
            return tally.Sum();
        }

        private static object SumIf(List<Expression> p)
        {
            // get parameters
            IEnumerable range = p[0] as IEnumerable;
            IEnumerable sumRange = p.Count < 3 ? range : p[2] as IEnumerable;
            var criteria = p[1].Evaluate();

            // build list of values in range and sumRange
            var rangeValues = new List<object>();
            foreach (var value in range)
            {
                rangeValues.Add(value);
            }
            var sumRangeValues = new List<object>();
            foreach (var value in sumRange)
            {
                sumRangeValues.Add(value);
            }

            //// compute total
            var ce = p[0].CreateSubContextCalcEngine();
            var tally = new Tally();
            for (int i = 0; i < System.Math.Min(rangeValues.Count, sumRangeValues.Count); i++)
            {
                if (ValueSatisfiesCriteria(rangeValues[i], criteria, ce))
                {
                    tally.AddValue(sumRangeValues[i]);
                }
            }

            // done
            return tally.Sum();
        }

        private static bool ValueSatisfiesCriteria(object value, object criteria, CalcEngine ce)
        {
            // safety...
            if (value == null)
            {
                return false;
            }

            // if criteria is a number, straight comparison
            if (criteria is double)
            {
                return (double)value == (double)criteria;
            }

            // convert criteria to string
            var cs = criteria as string;
            if (!string.IsNullOrEmpty(cs))
            {
                // if criteria is an expression (e.g. ">20"), use calc engine
                if (cs[0] == '=' || cs[0] == '<' || cs[0] == '>')
                {
                    // build expression
                    var expression = string.Format("{0}{1}", value, cs);

                    // add quotes if necessary
                    var pattern = @"(\w+)(\W+)(\w+)";
                    var m = Regex.Match(expression, pattern);
                    if (m.Groups.Count == 4)
                    {
                        double d;
                        if (!double.TryParse(m.Groups[1].Value, out d) || !double.TryParse(m.Groups[3].Value, out d))
                        {
                            expression = string.Format("\"{0}\"{1}\"{2}\"", m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
                        }
                    }

                    // evaluate
                    return (bool)ce.Evaluate(expression);
                }

                // if criteria is a regular expression, use regex
                if (cs.IndexOf('*') > -1)
                {
                    var pattern = cs.Replace(@"\", @"\\");
                    pattern = pattern.Replace(".", @"\");
                    pattern = pattern.Replace("*", ".*");
                    return Regex.IsMatch(value.ToString(), pattern, RegexOptions.IgnoreCase);
                }

                // straight string comparison
                return string.Equals(value.ToString(), cs, StringComparison.OrdinalIgnoreCase);
            }

            // should never get here?
            System.Diagnostics.Debug.Assert(false, "failed to evaluate criteria in SumIf");
            return false;
        }

        /// <summary>
        /// Sums the expression if the IF condition evaluates to true. The if condition and expressions 
        /// are evaluated against the DataContext for each element of the collection, and only when the if 
        /// condition result is true does the expression get evaluated. 
        /// <example> "sum = SUMEXPRIF(Entity.SomeCollection, \"IF(AND(IsSelected, HASVALUE(Premium), HASVALUE(CCM)), TRUE, FALSE)\", \"Premium\")" </example>
        /// </summary>
        /// <param name="p">The list of expressions.</param>
        /// <returns></returns>
        private static object SumExprIf(List<Expression> p)
        {
            // get parameters
            IEnumerable range = p[0].Evaluate() as IEnumerable;
            var criteria = p[1].Evaluate() as string;
            var propertyExpression = p[2].Evaluate() as string;

            // build list of values in range and sumRange
            var rangeValues = new List<object>();
            foreach (var value in range)
            {
                rangeValues.Add(value);
            }
            // compute total
            var tally = new Tally();
            // initialize variables of new CalcEngine with parent CalcEngine variables
            var ce = p[0].CreateSubContextCalcEngine();
            for (int i = 0; i < rangeValues.Count; i++)
            {
                if (rangeValues[i] != null)
                {
                    ce.DataContext = rangeValues[i];
                    if ((bool)ce.Evaluate(criteria) == true)
                    {
                        tally.AddValue(ce.Evaluate(propertyExpression));
                    }
                }
            }

            // done
            return tally.Sum();
        }

        private static object Tan(List<Expression> p)
        {
            return System.Math.Tan((double)p[0]);
        }

        private static object Tanh(List<Expression> p)
        {
            return System.Math.Tanh((double)p[0]);
        }

        private static object Trunc(List<Expression> p)
        {
            return (double)(int)((double)p[0]);
        }
    }
}