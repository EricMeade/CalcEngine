using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace CalcEngine.Functions
{
    internal static class Logical
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("AND", 1, int.MaxValue, And);
            ce.RegisterFunction("OR", 1, int.MaxValue, Or);
            ce.RegisterFunction("NOT", 1, Not);
            ce.RegisterFunction("IF", 3, If);
            ce.RegisterFunction("TRUE", 0, True);
            ce.RegisterFunction("FALSE", 0, False);
            ce.RegisterFunction("HASVALUE", 1, HasValue);
            ce.RegisterFunction("CONTAINS", 2, Contains);
            ce.RegisterFunction("LISTCOUNT", 1, ListCount);
            ce.RegisterFunction("ISNULL", 2, IsNull);
            ce.RegisterFunction("ISNUMERIC", 1, IsNumeric);
            ce.RegisterFunction("EVALEXPRIF", 3, 3, EvalExprIf);
            ce.RegisterFunction("EXECUTEEXPR", 1, 2, ExecuteExpr);
            ce.RegisterFunction("LSWITCH", 3, int.MaxValue, LSwitch);
            ce.RegisterFunction("SWITCH", 3, int.MaxValue, Switch);
            ce.RegisterFunction("GETPROPERTY", 2, GetProperty);
            ce.RegisterFunction("SETPROPERTY", 3, SetProperty);
            ce.RegisterFunction("ELEMENTAT", 2, 5, ElementAt);
            ce.RegisterFunction("DEEPCOPY", 1, DeepCopy);
            ce.RegisterFunction("NEW", 1, int.MaxValue, New);
            ce.RegisterFunction("THROWEX", 1, int.MaxValue, ThrowEx);
            ce.RegisterFunction("REGISTERFUNCTION", 2, int.MaxValue, RegisterFunction);
        }

#if DEBUG

        public static void Test(CalcEngine ce)
        {
            ce.Test("AND(true, true)", true);
            ce.Test("AND(true, false)", false);
            ce.Test("AND(false, true)", false);
            ce.Test("AND(false, false)", false);
            ce.Test("OR(true, true)", true);
            ce.Test("OR(true, false)", true);
            ce.Test("OR(false, true)", true);
            ce.Test("OR(false, false)", false);
            ce.Test("NOT(false)", true);
            ce.Test("NOT(true)", false);
            ce.Test("IF(5 > 4, true, false)", true);
            ce.Test("IF(5 > 14, true, false)", false);
            ce.Test("TRUE()", true);
            ce.Test("FALSE()", false);
            ce.Test("ISNUMERIC(23)", true);
            ce.Test("ISNUMERIC(3%)", true);

            #region //Test HASVALUE function

            double? nullableValue = null;
            ce.Variables["_nullTestVariable"] = nullableValue;
            ce.Test("HASVALUE(_nullTestVariable)", false);
            nullableValue = 105.0;
            ce.Variables["_nullTestVariable"] = nullableValue;
            ce.Test("HASVALUE(_nullTestVariable)", true);
            ce.Variables.Remove("_nullTestVariable");

            ce.Variables.Add("null", null);
            ce.Test("ISNULL(null, 1)", 1);
            ce.Variables.Remove("null");

            #endregion //Test HASVALUE function
            List<object> list = new List<object>() { "Item 1", "Item 2" };
            ce.Variables["_listCountTest"] = list;
            ce.Test("LISTCOUNT(_listCountTest)", 2);
            ce.Variables.Remove("_listCountTest");
        }

#endif

        private static object And(List<Expression> p)
        {
            var b = true;
            foreach (var v in p)
            {
                b = b && (bool)v;
            }
            return b;
        }

        private static object Or(List<Expression> p)
        {
            var b = false;
            foreach (var v in p)
            {
                b = b || (bool)v;
            }
            return b;
        }

        private static object Not(List<Expression> p)
        {
            return !(bool)p[0];
        }

        private static object If(List<Expression> p)
        {
            if (p[0] == null)
                return p[2].Evaluate();
            else
                return (bool)p[0] ? p[1].Evaluate() : p[2].Evaluate();
        }

        private static object True(List<Expression> p)
        {
            return true;
        }

        private static object False(List<Expression> p)
        {
            return false;
        }

        private static object HasValue(List<Expression> p)
        {
            var b = false;
            if (p[0] != null)
            {
                b = (p[0].Evaluate() == null) ? false : true;
                //Type t = p[0].GetType();
                //if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                //    b = false;
            }
            return b;
        }

        private static object Contains(List<Expression> p)
        {
            if (p[0] != null)
            {
                System.Collections.IEnumerable list = p[0].Evaluate() as System.Collections.IEnumerable;
                if (list != null)
                {
                    IEnumerable<object> genericList = list.Cast<object>();
                    object value = p[1].Evaluate();

                    if (genericList != null && value != null)
                    {
                        return genericList.Contains(value);
                    }
                }
            }
            return false;
        }

        private static object ListCount(List<Expression> p)
        {
            if (p[0] != null)
            {
                IEnumerable<object> list = p[0].Evaluate() as IEnumerable<object>;
                if (list != null)
                {
                    return list.Count();
                }
            }
            return 0;
        }
        private static object IsNull(List<Expression> p)
        {
            object p0 = p[0].Evaluate();

            // Check for empty strings
            if (p0 is string && string.IsNullOrEmpty((string)p0))
            {
                p0 = null;
            }

            return p0 ?? p[1].Evaluate();
        }

        private static object IsNumeric(List<Expression> p)
        {
            string strValue = (string)p[0];
            return (string.IsNullOrEmpty(strValue) == false && strValue.All(c => char.IsDigit(c) || c == '.' || c == ',') && strValue.Where(c => c == '.').Count() <= 1) ? true : false;
        }

        /// <summary>
        /// Registers a dynamic function which can be referenced in the CalcEngine.
        /// <example>
        ///        "RegisterFunction("AddNumbers", "P1", "P2", { @*_a = P1 + P2; } );"
        ///        *total = AddNumbers(15, 8);
        /// </example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters. 
        ///     p[0]=Function Name,
        ///     p[1]=Parameter 1 (Optional),
        ///     p[.]=Parameter 2 (Optional),
        ///     p[N-1]=Parameter N (Optional),
        ///     p[N]=Function code to execute
        /// </param>
        /// <returns></returns>
        private static object RegisterFunction(List<Expression> p)
        {
            // Get Parameters
            var functionName = p.First().Evaluate() as string;
            var functionExpression = p.Last() as FunctionExpression; // Automatically appended by CalcEngine to enable function to run in current CalcEngine context
            var functionCode = p[p.Count - 2].Evaluate() as string;
            var parameters = new List<string>();
            if (p.Count > 3)
                parameters.AddRange(p.GetRange(1, p.Count - 3).Select(e => e.Evaluate() as string));

            functionExpression._ce.RegisterDynamicFunction(functionName, parameters, functionCode, ExecuteFunc);
            return null;
        }

        private static object ExecuteFunc(List<Expression> p)
        {
            // Get Parameters
            var expression = p[0];
            var criteria = "true";
            var calculations = ((p.Count == 1) ? p[0].Evaluate() : p[1].Evaluate()) as string;

            return EvaluateConditionalExpression(expression, criteria, calculations);
        }


        /// <summary>
        /// Evaluates the expression if the IF condition evaluates to true. The if condition and expressions 
        /// are evaluated against the DataContext or each element of the collection, and only when the if 
        /// condition result is true does the expression get evaluated. The expression has to be in the form 
        /// of a calculation method, meaning, "A = B + C".
        /// 
        /// Creates variable _this as a 
        /// <example> "*_totalPremium = EVALEXPRIF(Entity.SomeCollection, \"IF(AND(IsSelected, HASVALUE(Premium)), TRUE, FALSE)\", \"*_tmpSum = _tmpSum + Premium [ | ... ]\")" </example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters. 
        ///     p[0]=IEnumerable or Context,
        ///     p[1]=Conditional expression to be evaluated for each item in the collection,
        ///     p[2]=String containing a list of calculation methods seperated by 
        ///         the <see cref="CALCULATION_SEPERATOR"/> character, 
        /// </param>
        /// <returns></returns>
        private static object EvalExprIf(List<Expression> p)
        {
            // Get Parameters
            var expression = p[0];
            var criteria = p[1].Evaluate();
            var calculations = p[2].Evaluate() as string;

            return EvaluateConditionalExpression(expression, (criteria ?? "").ToString(), calculations);
        }

        /// <summary>
        /// Execute expressions against the DataContext. The expression has to be in the form 
        /// of a calculation method, meaning, "A = B + C".
        /// <example> "*_totalPremium = EXECUTEEXPR(DataContext, \"*_tmpSum = _tmpSum + Premium [ ; ... ]\")" </example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters (2 Parameters)
        ///     p[0]=DataContext
        ///     p[1]=String containing a list of calculation methods seperated by the <see cref="CALCULATION_SEPERATOR"/> character
        ///     
        /// The list of expressions/parameters (1 Parameter)
        ///     p[0]=String containing a list of calculation methods seperated by the <see cref="CALCULATION_SEPERATOR"/> character
        /// </param>
        /// <returns></returns>
        private static object ExecuteExpr(List<Expression> p)
        {
            // Get Parameters
            var expression = p[0];
            var criteria = "true";
            var calculations = ((p.Count == 1) ? p[0].Evaluate() : p[1].Evaluate()) as string;

            return EvaluateConditionalExpression(expression, criteria, calculations);
        }

        private static object EvaluateConditionalExpression(Expression expression, string criteria, string calculations)
        {
            bool criteriaMatched;
            return EvaluateConditionalExpression(expression, criteria, calculations, out criteriaMatched);
        }

        private static object EvaluateConditionalExpression(Expression expression, string criteria, string calculations, out bool criteriaMatched)
        {
            criteriaMatched = false;
            var evaluatedExpression = expression.Evaluate();
            if (evaluatedExpression is IDictionary)
            {
                // Convert Dictionary to IList<KeyValuePair>
                evaluatedExpression = (evaluatedExpression as IDictionary).OfType<object>();
            }
            else if (expression is VariableExpression)
            {
                // If dynamic code, set evaluatedExpression to context of caller (which is CalcEngine.THIS)
                var vxce = (expression as VariableExpression)._ce;
                if (vxce.Variables.Keys.Any(v => v.StartsWith(Expression.DYNAMIC_CODE_DEFAULT_NAME) && vxce.Variables[v] == evaluatedExpression))
                {
                    evaluatedExpression = vxce.Variables[CalcEngine.THIS];
                }
            }
            IEnumerable<object> range = evaluatedExpression as IEnumerable<object>;

            // Build list of values in range and sumRange
            var rangeValues = new List<object>();
            if (range != null)
            {
                rangeValues.AddRange(range);
            }
            else
            {
                rangeValues.Add(evaluatedExpression);
            }

            // Initialize new calculation engine, from CalcEngine assiciated with the first parameter, to evaluate expressions
            var ce = expression.CreateSubContextCalcEngine();

            // Initialize variable containing reference to the enumerator data context
            object previousThis = ce.Variables.ContainsKey(CalcEngine.THIS) ? ce.Variables[CalcEngine.THIS] : null;
            object previousChanged = ce.Variables.ContainsKey(CalcEngine.CHANGED) ? ce.Variables[CalcEngine.CHANGED] : null;

            // Add variable to hold list of collection items that have changed
            ce.Variables[CalcEngine.CHANGED] = new List<object>();

            List<string> existingVariables = ce.Variables.Keys.ToList();

            bool expressionLoaded = false;
            for (int i = 0; i < rangeValues.Count; i++)
            {
                if (rangeValues[i] != null)
                {
                    ce.DataContext = rangeValues[i];
                    ce.Variables[CalcEngine.THIS] = rangeValues[i];

                    // Parsing the calculations can only be done after the DataContext has been set; therefore, 
                    // putting a flag around this so that the expressions are only loaded once
                    if (!expressionLoaded)
                    {
                        foreach (string calculation in SplitCalculations(calculations))
                        {
                            if (string.IsNullOrWhiteSpace(calculation) == true)
                                continue;
                            ce.AddCalculationMethod(calculation);
                        }

                        // If not set, default return property to lValue of last calculation method
                        if (ce.ReturnPropertyName == string.Empty)
                        {
                            if (ce.ReturnChangedCollection == true)
                                ce.ReturnPropertyName = CalcEngine.CHANGED;
                            else if (ce.CalculationMethodMappings.Count() > 0)
                                ce.ReturnPropertyName = ce.CalculationMethodMappings.Keys.Last();
                            else
                                ce.ReturnPropertyName = ce.Variables.Keys.Last();
                        }

                        expressionLoaded = true;
                    }

                    if ((bool)ce.Evaluate(criteria) == true)
                    {
                        criteriaMatched = true;
                        foreach (KeyValuePair<string, Expression> item in ce.CalculationMethodMappings)
                        {
                            var currentValue = (ce.Variables.Keys.Contains(item.Key)) ? ce.Variables[item.Key] : ce.DataContext.GetValue(item.Key);
                            var newValue = item.Value.Evaluate();
                            if (!ce.IsEqual(currentValue, newValue))
                            {
                                if (ce.Variables.Keys.Contains(item.Key))
                                {
                                    ce.Variables[item.Key] = newValue;
                                }
                                else
                                {
                                    ce.DataContext.SetValue(item.Key, newValue);
                                }
                            }
                        }

                        // Update changed list
                        (ce.Variables[CalcEngine.CHANGED] as List<object>).Add(ce.Variables[CalcEngine.THIS]);
                    }
                }
            }

            // Restore _this, get return value, remove any added variables and return result
            ce.Variables[CalcEngine.THIS] = previousThis;
            object returnValue = (string.IsNullOrEmpty(ce.ReturnPropertyName))? null : ((ce.Variables.Keys.Contains(ce.ReturnPropertyName)) ?
                ce.Variables[ce.ReturnPropertyName] : ce.DataContext.GetValue(ce.ReturnPropertyName));
            ce.Variables[CalcEngine.CHANGED] = previousChanged;
            ce.Variables.Keys.Except(existingVariables).ToList().ForEach(key => ce.Variables.Remove(key));

            return returnValue;
        }

        /// <summary>
        /// Executes a Literal Switch Statement
        /// <example>LSWITCH(YEARS('[RETRO_DATE]','[EFFECTIVE_DATE]'), 0, 'None', 1, 'One', 2, 'Two', 3, 'Three', 'Four or more')</example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters (3 or more parameters)
        ///     p[0]=Case Expression
        ///     p[n]=Condition
        ///     p[n+1]=Expression
        ///     
        /// The Default case is added at the end of the parameter list and does not have a corresponding Condition
        /// </param>
        /// <returns></returns>
        private static object LSwitch(List<Expression> p)
        {
            // Process expression as a literal 
            List<KeyValuePair<string, Expression>> conditions = new List<KeyValuePair<string, Expression>>();
            for (int i = 1; i < p.Count; i += 2)
            {
                String criteria = "";
                Expression calculation = null;

                if (p.Count > i + 1)
                {
                    // Case
                    criteria = p[i];
                    calculation = p[i + 1];
                }
                else
                {
                    // Default
                    calculation = p[i];
                }

                conditions.Add(new KeyValuePair<string, Expression>(criteria, calculation));
            }

            var caseSwitch = p[0].Evaluate();
            string switchValue = caseSwitch != null ? caseSwitch.ToString() : p[0];

            Expression caseExpression;
            if (conditions.Count(c => c.Key == switchValue) > 0)
            {
                caseExpression = conditions.FirstOrDefault(c => c.Key == switchValue).Value;
            }
            else
            {
                // Default Condition
                caseExpression = conditions.FirstOrDefault(c => c.Key == "").Value;
            }

            return caseExpression.Evaluate();
        }

        /// <summary>
        /// Executes a Switch Statement
        /// <example>SWITCH(context, ID = 0, 'None', ID = 1, 'One', ID = 2, 'Two', ID = 3, 'Three', 'Four or more')</example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters (3 or more parameters)
        ///     p[0]=Context
        ///     p[n]=Condition
        ///     p[n+1]=Expression
        ///     
        /// The Default case is added at the end of the parameter list and does not have a corresponding Condition
        /// </param>
        /// <returns></returns>
        private static object Switch(List<Expression> p)
        {
            for (int i = 1; i < p.Count; i += 2)
            {
                String criteria = "";
                Expression calculation = null;

                if (p.Count > i + 1)
                {
                    // Case
                    criteria = p[i];
                    calculation = p[i + 1];
                }
                else
                {
                    // Default
                    criteria = "TRUE";
                    calculation = p[i];
                }

                // Process expression as a conditional expression
                bool criteriaMatched = false;
                var result = EvaluateConditionalExpression(p[0], criteria, calculation, out criteriaMatched);
                if (criteriaMatched) return result;
            }

            return null;
        }

        private static object GetProperty(List<Expression> p)
        {
            var target = p[0].Evaluate();
            var propertyName = p[1].Evaluate() as string;

            if (target != null && !string.IsNullOrEmpty(propertyName))
                return target.GetValue(propertyName);
            return null;
        }

        private static object SetProperty(List<Expression> p)
        {
            var target = p[0].Evaluate();
            var propertyName = p[1].Evaluate() as string;
            var value = p[2].Evaluate();

            if (target != null && !string.IsNullOrEmpty(propertyName))
                target.SetValue(propertyName, value);
            return value;
        }

        private static object DeepCopy(List<Expression> p)
        {
            object returnValue = null;
            var source = p[0].Evaluate();
            if (source != null)
                returnValue = source.DeepCopy();
            return returnValue;
        }

        /// <summary>
        /// Returns element at specified index
        /// <example>ELEMENTAT(collection, 1)</example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters
        ///     p[0]=IEnumerable<object>,
        ///     p[1]=Property name or item index (string or int),
        ///     p[2]=Property value - only used if property name is specified
        ///     p[3]=Property2 name,
        ///     p[4]=Property2 value - only used if property 2 name is specified
        /// </param>
        /// <returns>Specified element or null if out of range</returns>
        private static object ElementAt(List<Expression> p)
        {
            object returnValue = null;

            IEnumerable<object> collection = p[0].Evaluate() as IEnumerable<object>;
            if (collection != null && collection.Count() != 0)
            {
                string propertyName = p[1].Evaluate() as string;
                if (string.IsNullOrWhiteSpace(propertyName))
                {
                    int index = Convert.ToInt32(p[1].Evaluate());
                    returnValue = collection.ElementAtOrDefault(index);
                }
                else
                {
                    var compareValue = (p.Count > 2) ? p[2].Evaluate() : null;
                    string property2Name = ((p.Count > 3) ? p[3].Evaluate() : null) as string;
                    var compare2Value = (p.Count > 4) ? p[4].Evaluate() : null;
                    bool compareBothParams = (string.IsNullOrWhiteSpace(property2Name)) ? false : true;

                    foreach (var obj in collection)
                    {
                        var propValue = obj.GetValue(propertyName);

                        #region // make sure types are the same for 1st compare property
                        if (propValue.GetType() != compareValue.GetType())
                        {
                            if (propValue is Enum && !(compareValue is Enum))
                            {
                                compareValue = Enum.Parse(propValue.GetType(), compareValue.ToString()) as IComparable;
                            }
                            else if (compareValue is Enum && !(propValue is Enum))
                            {
                                propValue = Enum.Parse(compareValue.GetType(), propValue.ToString()) as IComparable;
                            }
                            else
                            {
                                compareValue = Convert.ChangeType(compareValue, propValue.GetType()) as IComparable;
                            }
                        }
                        #endregion

                        if (propValue.Equals(compareValue))
                        {
                            if (compareBothParams == true)
                            {
                                var prop2Value = obj.GetValue(property2Name);

                                #region // make sure types are the same for 2nd compare property
                                if (prop2Value.GetType() != compare2Value.GetType())
                                {
                                    if (prop2Value is Enum && !(compare2Value is Enum))
                                    {
                                        compare2Value = Enum.Parse(prop2Value.GetType(), compare2Value.ToString()) as IComparable;
                                    }
                                    else if (compare2Value is Enum && !(prop2Value is Enum))
                                    {
                                        prop2Value = Enum.Parse(compare2Value.GetType(), prop2Value.ToString()) as IComparable;
                                    }
                                    else
                                    {
                                        compare2Value = Convert.ChangeType(compare2Value, prop2Value.GetType()) as IComparable;
                                    }
                                }
                                #endregion

                                if (prop2Value.Equals(compare2Value))
                                {
                                    returnValue = obj;
                                    break;
                                }
                            }
                            else
                            {
                                returnValue = obj;
                                break;
                            }
                        }
                    }
                }
            }
            return returnValue;
        }

        /// <summary>
        /// Creates a new instance of Type p[0] and optionally initializes with 1 or more values p[n].
        /// </summary>
        /// <example>
        /// List:
        ///     New('System.Collections.Generic.List`1[[System.String]]',
        ///         'Hello World',
        ///         'QWERTY'
        ///     )
        ///     
        /// Dictionary:
        ///     New('System.Collections.Generic.Dictionary`2[[System.String],[System.String]]',
        ///         'Key 1', 'Value 1',
        ///         'Key 2', 'Value 2',
        ///         'Key 3', 'Value 3',
        ///         'Key 4', 'Value 4',
        ///     )
        /// </example>
        /// <param name="p">
        ///     p[0]=Object Type
        ///     p[n]=Parameter 1
        ///     p[n+1]=Parameter N</param>
        /// <returns></returns>
        private static object New(List<Expression> p)
        {
            Type type = Type.GetType(p[0].Evaluate() as string);
            object newObject = Activator.CreateInstance(type);

            // Initialize List
            IList list = newObject as IList;
            if (list != null)
            {
                var typeArguments = type.GetGenericArguments();
                var listType = typeArguments[0];

                for (int i = 1; i < p.Count; i += 2)
                {
                    object value = p[i].Evaluate();

                    list.Add(value);
                }
            }

            // Initialize Dictionary
            IDictionary dictionary = newObject as IDictionary;
            if (dictionary != null)
            {
                var typeArguments = type.GetGenericArguments();
                var keyType = typeArguments[0];
                var valueType = typeArguments[1];

                for (int i = 1; i < p.Count; i += 2)
                {
                    object key = p[i].Evaluate();
                    object value = p[i + 1].Evaluate();

                    dictionary.Add(key, value);
                }
            }

            return newObject;
        }
        /// <summary>
        /// Throws an exception
        /// <example>THROWEX('Invalid layer price collection - PAL {0} (ID={1})', PricingAnalysisLayerName, PricingAnalysisLayerID)</example>
        /// </summary>
        /// <param name="p">The list of expressions/parameters (1 or more parameters)
        ///     p[0]=Format String
        ///     p[n]=Parameter 1
        ///     p[n+1]=Parameter N
        /// </param>
        /// <returns></returns>
        private static object ThrowEx(List<Expression> p)
        {
            var formatString = p[0].Evaluate() as string;
            object[] parameters = new object[p.Count - 1];
            for (int i = 1; i < p.Count; i++)
            {
                parameters[i - 1] = p[i].Evaluate();
            }
            var message = string.Format(formatString, parameters);
            throw new Exception(message);
        }

        /// <summary>
        /// Splits the calculations starting from the outer most level.
        /// </summary>
        /// <param name="calculations">The calculation.</param>
        /// <returns></returns>
        private static string[] SplitCalculations(string calculation)
        {
            List<string> calculations = new List<string>();

            int nestedLevel = 0;
            StringBuilder unitOfWork = new StringBuilder();

            // Split into unit of work
            foreach (char c in calculation.ToCharArray())
            {
                if (c == CalcEngine.START_CONTEXT_CHAR) nestedLevel++;
                if (c == CalcEngine.END_CONTEXT_CHAR) nestedLevel--;

                if ((c == CalcEngine.CALCULATION_SEPERATOR) && nestedLevel == 0)
                {
                    calculations.Add(unitOfWork.ToString());
                    unitOfWork.Clear();
                }
                else
                {
                    unitOfWork.Append(c);
                }
            }

            if (unitOfWork.Length > 0)
            {
                calculations.Add(unitOfWork.ToString());
            }

            return calculations.ToArray();
        }
    }
}