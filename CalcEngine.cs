using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using CalcEngine.Functions;

namespace CalcEngine
{
    /// <summary>
    /// Delegate that represents CalcEngine functions.
    /// </summary>
    /// <param name="parms">List of <see cref="Expression"/> objects that represent the
    /// parameters to be used in the function call.</param>
    /// <returns>The function result.</returns>
    public delegate object CalcEngineFunction(List<Expression> parms);

    /// <summary>
    /// CalcEngine parses strings and returns Expression objects that can be evaluated. Calculations, which are 
    /// expressions with an assignment, can also be loaded and processed. A dependency map is generated as the 
    /// calculations are loaded so that, when <see cref="DataContextPropertyChanged"/> is call, all calculations 
    /// dependending on the specified property will be evaluated.
    /// 
    /// <para>Variables are defined by prefixing the lValue with the <b><see cref="VARIABLE_INDICATOR"/></b>.</para>
    /// <para>The return value of the sub-context is the result of the last expression processed, in the sub-context, 
    /// unless otherwise specified by prefixing the lValue with the <b><see cref="RETURN_PROPERTY_INDICATOR"/></b>.</para>
    /// <para>This class has three extensibility points:</para>
    /// <para>Use the <b><see cref="DataContext"/></b> property to add an object's properties to the engine scope.</para>
    /// <para>Use the <b><see cref="RegisterFunction"/></b> method to define custom functions.</para>
    /// <para>Override the <b><see cref="GetExternalObject"/></b> method to add arbitrary variables to the engine scope.</para>
    /// <para>Set the <b><see cref="ReInitSubContextVariables"/></b> property to false to disable resetting 
    /// variables when they are used in sub-context calculation methods.</para>
    /// <para>Set the <b><see cref="ReturnChangedCollection"/></b> property to false to disable returning changed collection 
    /// in sub-context calculation methods. When fasle, the last calculation method loaded will be returned, if none exist, 
    /// then last variable loaded will be returned.</para>
    /// </summary>
    public partial class CalcEngine
    {
        #region // Private Properties

        private string _expr;				                    // expression being parsed

        private int _len;				                        // length of the expression being parsed
        private int _ptr;				                        // current pointer into expression
        private string _idChars;                                // valid characters in identifiers (besides alpha and digits)
        private Token _token;				                    // current token being parsed
        private Dictionary<object, Token> _tkTbl;               // table with tokens (+, -, etc)
        private Dictionary<string, FunctionDefinition> _fnTbl;  // table with constants and functions (pi, sin, etc)
        private IDictionary<string, object> _vars;              // table with variables
        private object _dataContext;                            // object with properties
        private bool _optimize;                                 // optimize expressions when parsing
        private ExpressionCache _cache;                         // cache with parsed expressions
        private CultureInfo _ci;                                // culture info used to parse numbers/dates
        private char _decimal, _listSep, _percent;              // localized decimal separator, list separator, percent sign
        private IDictionary<string, Expression> _calcMappings;  // Holds the mapping of property to parsed expression
        private string _returnPropertyName;                     // The property name to use to retrieive the result after processing all calculations

        /// <summary>
        /// Variable used when calculation does not contain an assignment (LValue)
        /// </summary>
        private static readonly string LVALUE_DEFAULT_NAME = "__av__";
        private static readonly string DEFAULT_VARIABLE_NAME_FORMAT = @"^((?<name>({0}))+(?<num>([\d]\d*))?)?$";
        private static readonly List<string> KEY_WORDS = new List<string>() { VARIABLE_INDICATOR, RETURN_PROPERTY_INDICATOR, THIS, ROOT, CHANGED };
        private static readonly List<char> GROUP_SEPERATORS = new List<char>() { '(', ')', START_CONTEXT_CHAR, END_CONTEXT_CHAR };


        /// <summary>
        /// The _dependency manager
        /// </summary>
        protected DependencyManager<string> DependencyManager;

        #endregion // Private Properties

        #region // Public Properties

        /// <summary>
        /// Holds the mapping of property to parsed expression
        /// </summary>
        public IDictionary<string, Expression> CalculationMethodMappings
        {
            get { return _calcMappings; }
            internal set { _calcMappings = value; }
        }

        /// <summary>
        /// The property name to use to retrieive the result after processing all calculations
        /// </summary>
        public string ReturnPropertyName
        {
            get { return _returnPropertyName; }
            internal set { _returnPropertyName = value; }
        }

        /// <summary>
        /// Character used in calculations to seperator <see cref="lValue"/> from the <see cref="Expression"/>
        /// </summary>
        public const string LVALUE_SEPERATOR = "=";

        /// <summary>
        /// Character used to indicate a variable in the <see cref="lValue"/> of the calculation. The VARIABLE_INDICATOR should be
        /// omitted when using the variable in an <see cref="Expression"/>
        /// </summary>
        public const string VARIABLE_INDICATOR = "var ";

        /// <summary>
        /// Character used to indicate the return value in the <see cref="lValue"/> of the calculation. The RETURN_PROPERTY_INDICATOR 
        /// should be omitted when using the variable in an <see cref="Expression"/>
        /// </summary>
        public const string RETURN_PROPERTY_INDICATOR = "return ";

        /// <summary>
        /// Character used to seperate a sequence of calculations when using the <see cref="EvalExprIF"/> and 
        /// <see cref="ExecuteExpr"/> functions. The CALCULATION_SEPERATOR is used to seperate the nested calculations.
        /// </summary>
        public const char CALCULATION_SEPERATOR  = ';';

        /// <summary>
        /// Character used at the end of a line to indicate the line continues on the next line. This character will be removed 
        /// from the input and the next line will be appended to this line.
        /// </summary>
        public const char LINE_CONTINUATION_CHAR = '&';

        /// <summary>
        /// Character sequences that indicate the start of a comment line. Comment lines will be removed from the calculations. 
        /// </summary>
        static public readonly string[] LINE_COMMENT_INDICATORS = { "//", "/*", "#region", "#endregion" };

        /// <summary>
        /// Character used to start context grouping within calculations when using the <see cref="EvalExprIF"/>, 
        /// <see cref="ExecuteExpr"/> and other functions.
        /// </summary>
        public const char START_CONTEXT_CHAR = '{';

        /// <summary>
        /// Character used to end context grouping within calculations when using the <see cref="EvalExprIF"/>, 
        /// <see cref="ExecuteExpr"/> and other functions.
        /// </summary>
        public const char END_CONTEXT_CHAR = '}';

        /// <summary>
        /// Variable containing reference to the enumerator data context
        /// </summary>
        public const string THIS = "this";

        /// <summary>
        /// Variable containing reference to the root data context
        /// </summary>
        public const string ROOT = "root";

        /// <summary>
        /// Variable containing reference to the changed enumerator items
        /// </summary>
        public const string CHANGED = "_changed";

        /// <summary>
        /// Set to false to disable resetting variables when they are used in sub-context calculation methods. 
        /// </summary>
        public bool ReInitSubContextVariables = true;

        /// <summary>
        /// Set to false to disable returning changed collection when no <see cref="RETURN_PROPERTY_INDICATOR"/> is specified.
        /// When fasle, the last calculation method loaded is returned, if only literals, then last variable loaded is returned.
        /// </summary>
        public bool ReturnChangedCollection = true;

        /// <summary>
        /// Gets the dependency graph.
        /// </summary>
        public string DependencyGraph
        {
            get { return DependencyManager.DependencyGraph; }
        }

        /// <summary>
        /// Gets the dependency count.
        /// </summary>
        public int DependencyCount
        {
            get { return DependencyManager.DependencyCount; }
        }

        /// <summary>
        /// Gets the precedents and related counts.
        /// </summary>
        public string Precedents
        {
            get { return DependencyManager.Precedents; }
        }

        #endregion // Public Properties

        #region // Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcEngine"/> class.
        /// </summary>
        public CalcEngine()
        {
            CultureInfo = CultureInfo.InvariantCulture;
            _tkTbl = GetSymbolTable();
            _fnTbl = GetFunctionTable();
            _vars = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            _cache = new ExpressionCache(this);
            _optimize = true;
            CalculationMethodMappings = new Dictionary<string, Expression>();
            DependencyManager = new DependencyManager<string>(new DependencyComparer());
            _returnPropertyName = string.Empty;
#if DEBUG
            this.Test();
#endif
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcEngine"/> class.
        /// </summary>
        /// <param name="dataContext">The data context.</param>
        public CalcEngine(object dataContext) : this()
        {
            DataContext = dataContext;
            _vars[CalcEngine.THIS] = dataContext;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcEngine"/> class.
        /// </summary>
        /// <param name="dataContext">The data context.</param>
        /// <param name="calculations">The list of calculations to load. A calculation is defines as "Property = Expression"</param>
        public CalcEngine(object dataContext, IList<string> calculations) : this(dataContext)
        {
            LoadCalculations(calculations);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcEngine"/> class.
        /// </summary>
        /// <param name="dataContext">The data context.</param>
        /// <param name="calculations">The list of calculations to load. A calculation is defines as "Property = Expression"</param>
        /// <param name="registerFnList">List of methods to register new calculation functions in the CalcEngine</param>
        public CalcEngine(object dataContext, IList<string> calculations, IEnumerable<Action<CalcEngine>> registerFnList)
            : this(dataContext)
        {
            if (registerFnList != null)
            {
                foreach (Action<CalcEngine> registerFn in registerFnList)
                    registerFn(this);
            }
            LoadCalculations(calculations);
        }

        #endregion // Constructors

        #region // Public Methods

        #region // Dependency Calculation Methods

        /// <summary>
        /// Updates the DataContext by re-evaluating all of the expressions dependent on the changed property.
        /// </summary>
        /// <param name="changedPropertyName">Name of the changed property.</param>
        /// <param name="recurseDependents">Set to true, the defaultvalue, to enable evaluation of all dependent 
        /// expressions; otherwise, only expressions containing the changed property name will be evaluated.</param>
        public void DataContextPropertyChanged(string changedPropertyName, bool recurseDependents = true)
        {
            if (recurseDependents == true)
            {
                DataContextPropertyChangedRecursive(changedPropertyName, changedPropertyName);
            }
            else
            {
                var dependents = DependencyManager.GetDependents(changedPropertyName);
                foreach (string propertyName in dependents)
                {
                    if (propertyName != changedPropertyName)
                    {
                        DataContext.SetValue(propertyName, CalculationMethodMappings[propertyName].Evaluate());
                    }
                }
            }
        }

        private void DataContextPropertyChangedRecursive(string changedPropertyName, string originalProperty)
        {
            var dependents = DependencyManager.GetDirectDependents(changedPropertyName);
            if (dependents != null)
            {
                foreach (string propertyName in dependents)
                {
                    if (propertyName != originalProperty)
                    {
                        var currentValue = (Variables.Keys.Contains(propertyName)) ? Variables[propertyName] : DataContext.GetValue(propertyName);
                        var newValue = CalculationMethodMappings[propertyName].Evaluate();
                        if (!IsEqual(currentValue, newValue))
                        {
                            if (Variables.Keys.Contains(propertyName))
                                Variables[propertyName] = newValue;
                            else
                                DataContext.SetValue(propertyName, newValue);
                            DataContextPropertyChangedRecursive(propertyName, originalProperty);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Loads the list of expressions into the calculation engine. The dependency map will be 
        /// updated based on the loaded calculations. If the calculations depend on any variables, 
        /// indicated by being prefixed with an asterik '*', then the variables must be loaded before 
        /// the calculations are loaded. The variables are initialized after all of the calculaions 
        /// have been loaded. Lines starting with "//" or "/*", comment lines, are ignored.
        /// </summary>
        /// <example>
        /// This sample shows the format of the calculations.
        ///     <code>
        ///         List<string> calculations = new List<string>() {
        ///             "*jdgmtFactor = IF(HASVALUE(JudgmentFactor), Entity.JudgmentFactor, 1)",
        ///             "EstimatedLossAmount = SelectedLayerPriceCalculatedPremium",
        ///             "EstimatedLostCostPercent = IF(AND(HASVALUE(MarketPrice), NOT(MarketPrice = 0)), (IF(HASVALUE(EstimatedLossAmount), EstimatedLossAmount, 0)/MarketPrice)*100, NULLVALUE())",
        ///             "GPMPercent = 100 - IF(HASVALUE(EstimatedLostCostPercent), EstimatedLostCostPercent, 0) - IF(HASVALUE(BrokerCommission), BrokerCommission, 0)",
        ///             "ModelExposurePrice = SelectedLayerPriceCalculatedPremium * IF(HASVALUE(JudgmentFactor), JudgmentFactor, 1)",
        ///             "ShareAmount = (SharePercent / 100) * ModelExposurePrice",
        ///             "MarketPrice = ModelExposurePrice",
        ///             // Adding line below will cause CircularReferenceException
        ///             //"SharePercent = IF(AND(HASVALUE(ExposureLimit),  NOT(ExposureLimit=0)), (ShareAmount / ExposureLimit) * 100, NULLVALUE())",
        ///         };
        ///         CalcEngine calcEngine = new CalcEngine(BuildPricingAnalysisLayer());
        ///         calcEngine.LoadCalculations(calculations);
        ///     </code>
        /// </example>   
        /// <param name="calculations">The calculations. A calculation is defines as "Property = Expression"</param>
        public void LoadCalculations(IList<string> calculations)
        {
            string calculation = string.Empty;
            foreach (string calcLine in calculations)
            {
                var nextCalc = (calcLine == null)? string.Empty : calcLine.Trim();

                // Remove blank and comment lines, they start with either "//", "/*", "#region" or "#endregion".
                if (nextCalc == string.Empty || LINE_COMMENT_INDICATORS.Any(c=> nextCalc.StartsWith(c)))
                    continue;

                calculation += nextCalc;
                switch (calculation.Last())
                {
                    case START_CONTEXT_CHAR:
                    case ',':
                        continue;
                    case LINE_CONTINUATION_CHAR:
                        calculation = calculation.TrimEnd(LINE_CONTINUATION_CHAR);
                        continue;
                    case CALCULATION_SEPERATOR:
                        if (HasOpenNestedCalculation(calculation))
                            continue;
                        calculation = calculation.TrimEnd(calculation.Last());
                        break;
                }
                AddCalculationMethod(calculation);
                calculation = string.Empty;
            }
   
            // Make sure all of the variables have been initialized
            foreach (string var in Variables.Keys.ToList())
            {
                if (CalculationMethodMappings.Keys.Contains(var))
                {
                    Variables[var] = CalculationMethodMappings[var].Evaluate();
                }
            }
        }

        /// <summary>
        /// Parses a string expression into an <see cref="Expression" /> stripping the assignment if it exists.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression" /> object that can be evaluated.</returns>
        public Expression ParseCalculation(string calculation)
        {
            string[] eqParams = SeperateCalculation(calculation);
            return (eqParams.Count() == 1) ? Parse(eqParams[0].Trim()) : Parse(eqParams[1].Trim());
        }

        /// <summary>
        /// Splits string on seperator starting from the outer most level. Seperator 
        /// characters within context groupings are ignored.
        /// </summary>
        /// <param name="calculations">The calculation.</param>
        /// <returns></returns>
        private static string[] SeperateCalculation(string calculation)
        {
            int indexOfSeperator = calculation.IndexOf(LVALUE_SEPERATOR);
            int indexOfGroup = calculation.IndexOfAny(GROUP_SEPERATORS.ToArray());
            if (indexOfSeperator < 0 || (indexOfGroup >=0 && indexOfSeperator > indexOfGroup))
                return new string[] { calculation };
            else
                return calculation.Split(LVALUE_SEPERATOR.ToCharArray(), 2);
        }

        /// <summary>
        /// Determines whether the specified calculation has an open nested calculation. 
        /// </summary>
        /// <param name="calculation">The calculation.</param>
        /// <returns></returns>
        private static bool HasOpenNestedCalculation(string calculation)
        {
            int nestedLevel = 0;

            // Split into unit of work
            foreach (char c in calculation.ToCharArray())
            {
                if (c == START_CONTEXT_CHAR) nestedLevel++;
                if (c == END_CONTEXT_CHAR) nestedLevel--;
            }
            return nestedLevel != 0 ? true : false;
        }

        /// <summary>
        /// Parses the expression, adds the calculation method and updates the dependency mappings for the 
        /// specified property name.
        /// </summary>
        /// <example>
        /// This sample shows the format of a calculation.
        ///     <code>
        ///         CalcEngine calcEngine = new CalcEngine(BuildPricingAnalysisLayer());
        ///         calcEngine.AddCalculationMethod("ModelExposurePrice = SelectedLayerPriceCalculatedPremium * IF(HASVALUE(JudgmentFactor), JudgmentFactor, 1)");
        ///     </code>
        /// </example>   
        /// <param name="calculation">The calculation. A calculation is defines as "Property = Expression"</param>
        /// <exception cref="System.Exception">Invalid expression - did not contain equal '=' sign</exception>
        public void AddCalculationMethod(string calculation)
        {
            string[] eqParams = SeperateCalculation(calculation);
            if (eqParams.Count() == 2)
            {
                AddCalculationMethod(eqParams[0], eqParams[1]);
            }
            else if (eqParams.Count() == 1)
            {
                // Initialize property name to next default variable name
                List<string> existingNames = Variables.Keys.Where(k => k.Contains(LVALUE_DEFAULT_NAME)).ToList();
                string propertyName = string.Format("{0}{1}", VARIABLE_INDICATOR, LVALUE_DEFAULT_NAME.GenerateVarName(DEFAULT_VARIABLE_NAME_FORMAT, existingNames));
                AddCalculationMethod(propertyName, eqParams[0]);
            }
            else
            {
                throw new Exception("Invalid expression - contained multiple LValues ('=' signs)");
            }
        }

        /// <summary>
        /// Parses the expression, adds the calculation method and updates the dependency mappings for the 
        /// specified property name. Use the <see cref="RETURN_VALUE_INDICATOR"/> to indicate that the 
        /// <see cref="lValue"/> will be the return value of the calculations. Use the <see cref="VARIABLE_INDICATOR"/> 
        /// to indicate that the calculation is specifing a variable instead of a <see cref="DataContext"/> property. 
        /// Variable calculations are added to the calculation methods only if they are not literals.
        /// </summary>
        /// <example>
        /// This sample shows the format of a calculation.
        ///     <code>
        ///         CalcEngine calcEngine = new CalcEngine(BuildPricingAnalysisLayer());
        ///         calcEngine.AddCalculationMethod("ModelExposurePrice", "SelectedLayerPriceCalculatedPremium * IF(HASVALUE(JudgmentFactor), JudgmentFactor, 1)");
        ///     </code>
        /// </example>   
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="expression">The expression.</param>
        public void AddCalculationMethod(string propertyName, string expression)
        {
            try
            {
                bool isVariable = false;
                bool loadCalculation = true;

                propertyName = propertyName.Trim();

                // set return property name if specified
                if (propertyName.StartsWith(RETURN_PROPERTY_INDICATOR) == true)
                {
                    propertyName = propertyName.Replace(RETURN_PROPERTY_INDICATOR, string.Empty);
                    ReturnPropertyName = propertyName.Replace(VARIABLE_INDICATOR, string.Empty);
                }

                // if property name is a variable, then add it to the variables collection
                if (propertyName.StartsWith(VARIABLE_INDICATOR) == true)
                {
                    propertyName = propertyName.Replace(VARIABLE_INDICATOR, string.Empty);
                    // variable name cannot be a keyword
                    if (KEY_WORDS.Contains(propertyName) == true)
                        throw new Exception(string.Format("Invalid variable - keywords ({0}) are not allowed as variable names", string.Join(", ", KEY_WORDS)));

                    if (!Variables.Keys.Contains(propertyName))
                        Variables.Add(propertyName, null);
                    else if (ReInitSubContextVariables)
                        Variables[propertyName] = null;
                    isVariable = true;
                }

                Expression parsedExpression = Parse(expression.Trim());

                // if property is a variable and expression is a literal, then update the variable with literal value 
                if (isVariable == true && parsedExpression.IsLiteral == true)
                {
                    Variables[propertyName] = parsedExpression.Evaluate();
                    loadCalculation = false;
                }

                // Only update calculation mappings if property is not a varaible and the expression is not a literal
                if (loadCalculation == true && !CalculationMethodMappings.Keys.Contains(propertyName))
                {
                    // Add expression to calculation method mappings
                    CalculationMethodMappings.Add(propertyName, parsedExpression);

                    // Update dependencies
                    IList<string> bindings = parsedExpression.GetBindings();
                    foreach (string contextPropertyName in bindings)
                    {
                        DependencyManager.AddTarget(contextPropertyName);
                        DependencyManager.AddDepedency(contextPropertyName, propertyName);
                    }
                }
            }
            catch (Exception ex)
            {
                Throw(string.Format("{0}\nProcessing: {1} = {2}", ex.Message, propertyName, expression), ex);
            }
        }

        /// <summary>
        /// Sets the value of the specified property of the data context and evaluates property dependent calculations.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <param name="value">The value.</param>
        /// <param name="recurseDependents">Set to true, the defaultvalue, to enable evaluation of all dependent 
        /// expressions; otherwise, only expressions containing the changed property name will be evaluated.</param>
        public void SetProperty(string propertyName, object value, bool recurseDependents = true)
        {
            DataContext.SetValue(propertyName, value);
            DataContextPropertyChanged(propertyName, recurseDependents);
        }

        /// <summary>
        /// Gets the value of the specified property of the data context.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns></returns>
        public object GetProperty(string propertyName)
        {
            return DataContext.GetValue(propertyName);
        }

        /// <summary>
        /// Clears the dependency map.
        /// </summary>
        public void ClearDependencyMap()
        {
            DependencyManager.Clear();
        }


        #endregion // Dependency Calculation Methods

        #region // CalcEngine Methods

        /// <summary>
        /// Parses a string into an <see cref="Expression"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An <see cref="Expression"/> object that can be evaluated.</returns>
        public Expression Parse(string expression)
        {
            // initialize
            _expr = expression;
            _len = _expr.Length;
            _ptr = 0;

            // skip leading equals sign
            if (_len > 0 && _expr[0] == '=')
            {
                _ptr++;
            }

            // parse the expression
            var expr = ParseExpression();

            // check for errors
            if (_token.ID != TKID.END)
            {
                Throw();
            }

            // optimize expression
            if (_optimize)
            {
                expr = expr.Optimize();
            }

            // done
            return expr;
        }

        /// <summary>
        /// Evaluates a string.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        public object Evaluate(string expression)
        {
            var x = _cache != null ? _cache[expression] : Parse(expression);
            return x.Evaluate();
        }

        public void RecalculateAll()
        {
            foreach (KeyValuePair<string, Expression> item in CalculationMethodMappings)
            {
                var currentValue = (Variables.Keys.Contains(item.Key)) ? Variables[item.Key] : DataContext.GetValue(item.Key);
                var newValue = item.Value.Evaluate();
                if (!IsEqual(currentValue, newValue))
                {
                    if (Variables.Keys.Contains(item.Key))
                    {
                        Variables[item.Key] = newValue;
                    }
                    else
                    {
                        DataContext.SetValue(item.Key, newValue);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the calc engine should keep a cache with parsed
        /// expressions.
        /// </summary>
        public bool CacheExpressions
        {
            get { return _cache != null; }
            set
            {
                if (value != CacheExpressions)
                {
                    _cache = value ? new ExpressionCache(this) : null;
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the calc engine should optimize expressions when
        /// they are parsed.
        /// </summary>
        public bool OptimizeExpressions
        {
            get { return _optimize; }
            set { _optimize = value; }
        }

        /// <summary>
        /// Gets or sets a string that specifies special characters that are valid for identifiers.
        /// </summary>
        /// <remarks>
        /// Identifiers must start with a letter or an underscore, which may be followed by
        /// additional letters, underscores, or digits. This string allows you to specify
        /// additional valid characters such as ':' or '!' (used in Excel range references
        /// for example).
        /// </remarks>
        public string IdentifierChars
        {
            get { return _idChars; }
            set { _idChars = value; }
        }

        /// <summary>
        /// Registers a function that can be evaluated by this <see cref="CalcEngine"/>.
        /// </summary>
        /// <param name="functionName">Function name.</param>
        /// <param name="parmMin">Minimum parameter count.</param>
        /// <param name="parmMax">Maximum parameter count.</param>
        /// <param name="fn">Delegate that evaluates the function.</param>
        public void RegisterFunction(string functionName, int parmMin, int parmMax, CalcEngineFunction fn)
        {
            _fnTbl.Add(functionName, new FunctionDefinition(parmMin, parmMax, fn));
        }

        /// <summary>
        /// Registers a function that can be evaluated by this <see cref="CalcEngine"/>.
        /// </summary>
        /// <param name="functionName">Function name.</param>
        /// <param name="parmCount">Parameter count.</param>
        /// <param name="fn">Delegate that evaluates the function.</param>
        public void RegisterFunction(string functionName, int parmCount, CalcEngineFunction fn)
        {
            RegisterFunction(functionName, parmCount, parmCount, fn);
        }

        /// <summary>
        /// Registers a dynamic function that can be evaluated by this <see cref="CalcEngine"/>.
        /// </summary>
        /// <param name="functionName">Function name.</param>
        /// <param name="functionParamters"><see cref="List<string>"/> Containing parameter names.</param>
        /// <param name="functionCode">Dynamic code to be evaluated.</param>
        public void RegisterDynamicFunction(string functionName, List<string> functionParamters, string functionCode, CalcEngineFunction fn)
        {
            var fnDef = new FunctionDefinition(functionParamters, functionCode, fn);
            if (!_fnTbl.ContainsKey(functionName))
            {
                _fnTbl.Add(functionName, fnDef);
            }
            else
            {
                _fnTbl[functionName] = fnDef;
            }
        }

        /// <summary>
        /// Gets an external object based on an identifier.
        /// </summary>
        /// <remarks>
        /// This method is useful when the engine needs to create objects dynamically.
        /// For example, a spreadsheet calc engine would use this method to dynamically create cell
        /// range objects based on identifiers that cannot be enumerated at design time
        /// (such as "AB12", "A1:AB12", etc.)
        /// </remarks>
        public virtual object GetExternalObject(string identifier)
        {
            return null;
        }

        /// <summary>
        /// Gets or sets the DataContext for this <see cref="CalcEngine"/>.
        /// </summary>
        /// <remarks>
        /// Once a DataContext is set, all public properties of the object become available
        /// to the CalcEngine, including sub-properties such as "Address.Street". These may
        /// be used with expressions just like any other constant.
        /// </remarks>
        public virtual object DataContext
        {
            get { return _dataContext; }
            set { 
                _dataContext = value;

                // Only set the root context if it does not already exist
                if (_vars.ContainsKey(CalcEngine.ROOT) == false)
                    _vars[CalcEngine.ROOT] = value;
            }
        }

        /// <summary>
        /// Gets the dictionary that contains function definitions.
        /// </summary>
        public Dictionary<string, FunctionDefinition> Functions
        {
            get { return _fnTbl; }
            internal set { _fnTbl = value; }
        }

        /// <summary>
        /// Gets the dictionary that contains simple variables (not in the DataContext).
        /// </summary>
        public IDictionary<string, object> Variables
        {
            get { return _vars; }
            set { _vars = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="CultureInfo"/> to use when parsing numbers and dates.
        /// </summary>
        public CultureInfo CultureInfo
        {
            get { return _ci; }
            set
            {
                _ci = value;
                var nf = _ci.NumberFormat;
                _decimal = nf.NumberDecimalSeparator[0];
                _percent = nf.PercentSymbol[0];
                _listSep = _ci.TextInfo.ListSeparator[0];
            }
        }

        #endregion // CalcEngine Methods

        #endregion // Public Methods

        #region // Private Methods

        /// <summary>
        /// Determines whether the specified current value is equal to the new value.
        /// </summary>
        /// <param name="currentValue">The current value.</param>
        /// <param name="newValue">The new value.</param>
        /// <returns></returns>
        public bool IsEqual(object currentValue, object newValue)
        {
            bool isEqual = false;

            if (currentValue == null && newValue == null)
                isEqual = true;
            else if (currentValue != null && newValue != null)
            {
                if (currentValue.GetType().IsValueType)
                    isEqual = currentValue.Equals(newValue);
                else if (newValue.GetType().IsValueType)
                    isEqual = newValue.Equals(currentValue);
                else
                    isEqual = currentValue == newValue;
            }

            return isEqual;
        }

        #region // Token/Keyword Tables

        // build/get static token table
        private Dictionary<object, Token> GetSymbolTable()
        {
            if (_tkTbl == null)
            {
                _tkTbl = new Dictionary<object, Token>();
                AddToken('+', TKID.ADD, TKTYPE.ADDSUB);
                AddToken('-', TKID.SUB, TKTYPE.ADDSUB);
                AddToken('(', TKID.OPEN, TKTYPE.GROUP);
                AddToken(')', TKID.CLOSE, TKTYPE.GROUP);
                AddToken('*', TKID.MUL, TKTYPE.MULDIV);
                AddToken('.', TKID.PERIOD, TKTYPE.GROUP);
                AddToken('/', TKID.DIV, TKTYPE.MULDIV);
                AddToken('\\', TKID.DIVINT, TKTYPE.MULDIV);
                AddToken('=', TKID.EQ, TKTYPE.COMPARE);
                AddToken('>', TKID.GT, TKTYPE.COMPARE);
                AddToken('<', TKID.LT, TKTYPE.COMPARE);
                AddToken('^', TKID.POWER, TKTYPE.POWER);
                AddToken("<>", TKID.NE, TKTYPE.COMPARE);
                AddToken(">=", TKID.GE, TKTYPE.COMPARE);
                AddToken("<=", TKID.LE, TKTYPE.COMPARE);

                // list separator is localized, not necessarily a comma
                // so it can't be on the static table
                //AddToken(',', TKID.COMMA, TKTYPE.GROUP);
            }
            return _tkTbl;
        }

        private void AddToken(object symbol, TKID id, TKTYPE type)
        {
            var token = new Token(symbol, id, type);
            _tkTbl.Add(symbol, token);
        }

        // build/get static keyword table
        private Dictionary<string, FunctionDefinition> GetFunctionTable()
        {
            if (_fnTbl == null)
            {
                // create table
                _fnTbl = new Dictionary<string, FunctionDefinition>(StringComparer.InvariantCultureIgnoreCase);

                // register built-in functions (and constants)
                Logical.Register(this);
                MathTrig.Register(this);
                Text.Register(this);
                Statistical.Register(this);
                Date.Register(this);
            }
            return _fnTbl;
        }

        #endregion // Token/Keyword Tables

        #region // Parser

        private void GetToken()
        {
            // eat white space
            while (_ptr < _len && _expr[_ptr] <= ' ')
            {
                _ptr++;
            }

            // are we done?
            if (_ptr >= _len)
            {
                _token = new Token(null, TKID.END, TKTYPE.GROUP);
                return;
            }

            // prepare to parse
            int i;
            var c = _expr[_ptr];

            // operators
            // this gets called a lot, so it's pretty optimized.
            // note that operators must start with non-letter/digit characters.
            var isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
            var isDigit = c >= '0' && c <= '9';
            if (!isLetter && !isDigit)
            {
                // if this is a number starting with a decimal, don't parse as operator
                var nxt = _ptr + 1 < _len ? _expr[_ptr + 1] : 0;
                bool isNumber = c == _decimal && nxt >= '0' && nxt <= '9';
                if (!isNumber)
                {
                    // look up localized list separator
                    if (c == _listSep)
                    {
                        _token = new Token(c, TKID.COMMA, TKTYPE.GROUP);
                        _ptr++;
                        return;
                    }

                    // look up single-char tokens on table
                    Token tk;
                    if (_tkTbl.TryGetValue(c, out tk))
                    {
                        // save token we found
                        _token = tk;
                        _ptr++;

                        // look for double-char tokens (special case)
                        if (_ptr < _len && (c == '>' || c == '<'))
                        {
                            if (_tkTbl.TryGetValue(_expr.Substring(_ptr - 1, 2), out tk))
                            {
                                _token = tk;
                                _ptr++;
                            }
                        }

                        // found token on the table
                        return;
                    }
                }
            }

            // parse numbers
            if (isDigit || c == _decimal)
            {
                var sci = false;
                var pct = false;
                var div = -1.0; // use double, not int (this may get really big)
                var val = 0.0;
                for (i = 0; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];

                    // digits always OK
                    if (c >= '0' && c <= '9')
                    {
                        val = val * 10 + (c - '0');
                        if (div > -1)
                        {
                            div *= 10;
                        }
                        continue;
                    }

                    // one decimal is OK
                    if (c == _decimal && div < 0)
                    {
                        div = 1;
                        continue;
                    }

                    // scientific notation?
                    if ((c == 'E' || c == 'e') && !sci)
                    {
                        sci = true;
                        c = _expr[_ptr + i + 1];
                        if (c == '+' || c == '-') i++;
                        continue;
                    }

                    // percentage?
                    if (c == _percent)
                    {
                        pct = true;
                        i++;
                        break;
                    }

                    // end of literal
                    break;
                }

                // end of number, get value
                if (!sci)
                {
                    // much faster than ParseDouble
                    if (div > 1)
                    {
                        val /= div;
                    }
                    if (pct)
                    {
                        val /= 100.0;
                    }
                }
                else
                {
                    var lit = _expr.Substring(_ptr, i);
                    val = ParseDouble(lit, _ci);
                }

                // build token
                _token = new Token(val, TKID.ATOM, TKTYPE.LITERAL);

                // advance pointer and return
                _ptr += i;
                return;
            }

            // parse strings - starting with double or single qoute
            if (c == '\"' || c == '\'')
            {
                char strStartChar = c;

                // look for end quote, skip double quotes
                for (i = 1; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];
                    if (c != strStartChar) continue;
                    char cNext = i + _ptr < _len - 1 ? _expr[_ptr + i + 1] : ' ';
                    if (cNext != strStartChar) break;
                    i++;
                }

                // check that we got the end of the string
                if (c != strStartChar)
                {
                    Throw("Can't find final quote.");
                }

                // end of string
                var lit = _expr.Substring(_ptr + 1, i - 1);
                _ptr += i + 1;
                switch (strStartChar)
                {
                    case '\"': _token = new Token(lit.Replace("\"\"", "\""), TKID.ATOM, TKTYPE.LITERAL); break;
                    case '\'': _token = new Token(lit.Replace("\'\'", "\'"), TKID.ATOM, TKTYPE.LITERAL); break;
                }
                return;
            }

            // parse dates (review)
            if (c == '#')
            {
                // look for end #
                for (i = 1; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];
                    if (c == '#') break;
                }

                // check that we got the end of the date
                if (c != '#')
                {
                    Throw("Can't find final date delimiter ('#').");
                }

                // end of date
                var lit = _expr.Substring(_ptr + 1, i - 1);
                _ptr += i + 1;
                _token = new Token(DateTime.Parse(lit, _ci), TKID.ATOM, TKTYPE.LITERAL);
                return;
            }

            // parse sub-context - starting with {
            if (c == START_CONTEXT_CHAR)
            {
                char strStartChar = c;
                char strEndChar = END_CONTEXT_CHAR;
                int subContextCount = 0;

                // look for end #
                for (i = 1; i + _ptr < _len; i++)
                {
                    c = _expr[_ptr + i];
                    if (c == strStartChar)
                    {
                        subContextCount++;
                        continue;
                    }
                    else if (c == strEndChar)
                    {
                        if (subContextCount == 0)
                            break;
                        else
                            subContextCount--;
                    }
                }

                // check that we got the end of the date
                if (c != strEndChar)
                {
                    Throw("Can't find final sub-context delimiter ('}').");
                }

                // end of sub-context
                var lit = _expr.Substring(_ptr + 1, i - 1);
                _ptr += i + 1;
                _token = new Token(lit, TKID.ATOM, TKTYPE.LITERAL);
                return;
            }

            // identifiers (functions, objects) must start with alpha or underscore
            if (!isLetter && c != '_' && (_idChars == null || _idChars.IndexOf(c) < 0))
            {
                Throw("Identifier expected.");
            }

            // and must contain only letters/digits/_idChars
            for (i = 1; i + _ptr < _len; i++)
            {
                c = _expr[_ptr + i];
                isLetter = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
                isDigit = c >= '0' && c <= '9';
                if (!isLetter && !isDigit && c != '_' && (_idChars == null || _idChars.IndexOf(c) < 0))
                {
                    break;
                }
            }

            // got identifier
            var id = _expr.Substring(_ptr, i);
            _ptr += i;
            _token = new Token(id, TKID.ATOM, TKTYPE.IDENTIFIER);
        }

        private static double ParseDouble(string str, CultureInfo ci)
        {
            if (str.Length > 0 && str[str.Length - 1] == ci.NumberFormat.PercentSymbol[0])
            {
                str = str.Substring(0, str.Length - 1);
                return double.Parse(str, NumberStyles.Any, ci) / 100.0;
            }
            return double.Parse(str, NumberStyles.Any, ci);
        }

        private List<Expression> GetParameters() // e.g. myfun(a, b, c+2)
        {
            // check whether next token is a (,
            // restore state and bail if it's not
            var pos = _ptr;
            var tk = _token;
            GetToken();
            if (_token.ID != TKID.OPEN)
            {
                _ptr = pos;
                _token = tk;
                return null;
            }

            // check for empty Parameter list
            pos = _ptr;
            GetToken();
            if (_token.ID == TKID.CLOSE)
            {
                return null;
            }
            _ptr = pos;

            // get Parameters until we reach the end of the list
            var parms = new List<Expression>();
            var expr = ParseExpression();
            parms.Add(expr);
            while (_token.ID == TKID.COMMA)
            {
                expr = ParseExpression();
                parms.Add(expr);
            }

            // make sure the list was closed correctly
            if (_token.ID != TKID.CLOSE)
            {
                Throw();
            }

            // done
            return parms;
        }

        private Token GetMember()
        {
            // check whether next token is a MEMBER token ('.'),
            // restore state and bail if it's not
            var pos = _ptr;
            var tk = _token;
            GetToken();
            if (_token.ID != TKID.PERIOD)
            {
                _ptr = pos;
                _token = tk;
                return null;
            }

            // skip member token
            GetToken();
            if (_token.Type != TKTYPE.IDENTIFIER)
            {
                Throw("Identifier expected");
            }
            return _token;
        }

        #region // Parser Helpers

        private Expression ParseExpression()
        {
            GetToken();
            return ParseCompare();
        }

        private Expression ParseCompare()
        {
            var x = ParseAddSub();
            while (_token.Type == TKTYPE.COMPARE)
            {
                var t = _token;
                GetToken();
                var exprArg = ParseAddSub();
                x = new BinaryExpression(t, x, exprArg);
            }
            return x;
        }

        private Expression ParseAddSub()
        {
            var x = ParseMulDiv();
            while (_token.Type == TKTYPE.ADDSUB)
            {
                var t = _token;
                GetToken();
                var exprArg = ParseMulDiv();
                x = new BinaryExpression(t, x, exprArg);
            }
            return x;
        }

        private Expression ParseMulDiv()
        {
            var x = ParsePower();
            while (_token.Type == TKTYPE.MULDIV)
            {
                var t = _token;
                GetToken();
                var a = ParsePower();
                x = new BinaryExpression(t, x, a);
            }
            return x;
        }

        private Expression ParsePower()
        {
            var x = ParseUnary();
            while (_token.Type == TKTYPE.POWER)
            {
                var t = _token;
                GetToken();
                var a = ParseUnary();
                x = new BinaryExpression(t, x, a);
            }
            return x;
        }

        private Expression ParseUnary()
        {
            // unary plus and minus
            if (_token.ID == TKID.ADD || _token.ID == TKID.SUB)
            {
                var t = _token;
                GetToken();
                var a = ParseAtom();
                return new UnaryExpression(t, a);
            }

            // not unary, return atom
            return ParseAtom();
        }

        private Expression ParseAtom()
        {
            string id;
            Expression x = null;
            FunctionDefinition fnDef = null;

            switch (_token.Type)
            {
                // literals
                case TKTYPE.LITERAL:
                    x = new Expression(_token);
                    break;

                // identifiers
                case TKTYPE.IDENTIFIER:

                    // get identifier
                    id = (string)_token.Value;

                    // look for functions
                    if (_fnTbl.TryGetValue(id, out fnDef))
                    {
                        var p = GetParameters();
                        var pCnt = p == null ? 0 : p.Count;
                        if (fnDef.ParmMin != -1 && pCnt < fnDef.ParmMin)
                        {
                            Throw(string.Format("Too few parameters '{0}'.", _token.Value));
                        }
                        if (fnDef.ParmMax != -1 && pCnt > fnDef.ParmMax)
                        {
                            Throw(string.Format("Too many parameters '{0}'.", _token.Value));
                        }
                        x = new FunctionExpression(this, fnDef, p);
                        break;
                    }

                    // look for simple variables (much faster than binding!)
                    if (_vars.ContainsKey(id))
                    {
                        x = new VariableExpression(this, _vars, id);
                        break;
                    }

                    // look for external objects
                    var xObj = GetExternalObject(id);
                    if (xObj != null)
                    {
                        x = new XObjectExpression(xObj);
                        break;
                    }

                    // look for bindings
                    if (DataContext != null)
                    {
                        var list = new List<BindingInfo>();
                        for (var t = _token; t != null; t = GetMember())
                        {
                            list.Add(new BindingInfo((string)t.Value, GetParameters()));
                        }
                        x = new BindingExpression(this, list, _ci);
                        break;
                    }
                    Throw(string.Format("Unexpected identifier '{0}'", _token.Value));
                    break;

                // sub-expressions
                case TKTYPE.GROUP:

                    // anything other than opening parenthesis is illegal here
                    if (_token.ID != TKID.OPEN)
                    {
                        Throw(string.Format("Expression expected '{0}'", _token.Value));
                    }

                    // get expression
                    GetToken();
                    x = ParseCompare();

                    // check that the parenthesis was closed
                    if (_token.ID != TKID.CLOSE)
                    {
                        Throw(string.Format("Unbalanced parenthesis '{0}'", _token.Value));
                    }

                    break;
            }

            // make sure we got something...
            if (x == null)
            {
                Throw();
            }

            // done
            GetToken();
            return x;
        }

        #endregion // Parser Helpers

        #endregion // Parser

        #region // Static Helpers

        private static void Throw()
        {
            Throw("Syntax error.");
        }

        private static void Throw(string msg, Exception innerException = null)
        {
            throw new Exception(msg, innerException);
        }

        #endregion // Static Helpers

        #endregion // Private Methods
    }
}