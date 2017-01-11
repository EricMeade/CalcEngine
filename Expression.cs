using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace CalcEngine
{
    /// <summary>
    /// Base class that represents parsed expressions.
    /// </summary>
    /// <remarks>
    /// For example:
    /// <code>
    /// Expression expr = scriptEngine.Parse(strExpression);
    /// object val = expr.Evaluate();
    /// </code>
    /// </remarks>
    public class Expression : IComparable<Expression>
    {
        //---------------------------------------------------------------------------

        #region // fields

        internal Token _token;
        private static CultureInfo _ci = CultureInfo.InvariantCulture;

        #endregion // fields

        //---------------------------------------------------------------------------

        #region // ctors

        internal Expression()
        {
            _token = new Token(null, TKID.ATOM, TKTYPE.IDENTIFIER);
        }

        internal Expression(object value)
        {
            _token = new Token(value, TKID.ATOM, TKTYPE.LITERAL);
        }

        internal Expression(Token tk)
        {
            _token = tk;
        }

        #endregion // ctors

        //---------------------------------------------------------------------------

        #region // object model

        public virtual object Evaluate()
        {
            if (_token.Type != TKTYPE.LITERAL)
            {
                throw new ArgumentException("Bad expression.");
            }
            return _token.Value;
        }

        public virtual Expression Optimize()
        {
            return this;
        }

        public virtual List<string> GetBindings()
        {
            List<string> bindings = new List<string>();
            return bindings;
        }

        protected void AddBindingRange(List<string> bindings, Expression expr)
        {
            if (bindings != null && expr != null)
            {
                var newBindings = expr.GetBindings();
                foreach (string name in newBindings)
                {
                    if (!bindings.Contains(name))
                        bindings.Add(name);
                }
            }
        }

        public virtual bool IsLiteral
        {
            get { return _token.Type == TKTYPE.LITERAL; }
        }

        #endregion // object model

        //---------------------------------------------------------------------------

        #region // implicit converters

        public static implicit operator string(Expression x)
        {
            var v = x.Evaluate();
            return v == null ? string.Empty : v.ToString();
        }

        public static implicit operator double(Expression x)
        {
            // evaluate
            var v = x.Evaluate();

            // handle doubles
            if (v is double)
            {
                return (double)v;
            }

            // handle booleans
            if (v is bool)
            {
                return (bool)v ? 1 : 0;
            }

            // handle dates
            if (v is DateTime)
            {
                return ((DateTime)v).ToOADate();
            }

            // handle nulls
            if (v == null)
            {
                return 0;
            }

            // handle everything else
            return (double)Convert.ChangeType(v, typeof(double), _ci);
        }

        public static implicit operator bool(Expression x)
        {
            // evaluate
            var v = x.Evaluate();

            // handle booleans
            if (v is bool)
            {
                return (bool)v;
            }

            // handle nulls
            if (v == null)
            {
                return false;
            }

            // handle doubles
            if (v is double)
            {
                return (double)v == 0 ? false : true;
            }

            // handle everything else
            return (double)x == 0 ? false : true;
        }

        public static implicit operator DateTime(Expression x)
        {
            // evaluate
            var v = x.Evaluate();

            // handle dates
            if (v is DateTime)
            {
                return (DateTime)v;
            }

            // handle doubles
            if (v is double)
            {
                return DateTime.FromOADate((double)x);
            }

            // handle nulls
            if (v == null || string.IsNullOrWhiteSpace(v.ToString()))
            {
                return DateTime.MinValue;
            }

            // handle everything else
            return (DateTime)Convert.ChangeType(v, typeof(DateTime), _ci);
        }

        #endregion // implicit converters

        //---------------------------------------------------------------------------

        #region // IComparable<Expression>

        public int CompareTo(Expression other)
        {
            // get both values
            var c1 = this.Evaluate() as IComparable;
            var c2 = other.Evaluate() as IComparable;

            // handle nulls
            if (c1 == null && c2 == null)
            {
                return 0;
            }
            if (c2 == null)
            {
                return -1;
            }
            if (c1 == null)
            {
                return +1;
            }

            // make sure types are the same
            if (c1.GetType() != c2.GetType())
            {
                //if (c1 is Enum && !(c2 is Enum))
                //{
                //    if (Enum.GetUnderlyingType((Enum)c1) == typeof(string))

                //    object uType = GetAsUnderlyingType((Enum)c1);

                //    c1 =  as IComparable;
                //}
                //else if (c2 is Enum && !(c1 is Enum))
                //{
                //    c1 = GetAsUnderlyingType((Enum)c1) as IComparable;
                //}
                //else
                //{
                //    c2 = Convert.ChangeType(c2, c1.GetType()) as IComparable;
                //}
                if (c1 is Enum && !(c2 is Enum))
                {
                    c2 = Enum.Parse(c1.GetType(), c2.ToString()) as IComparable;
                }
                else if (c2 is Enum && !(c1 is Enum))
                {
                    c1 = Enum.Parse(c2.GetType(), c1.ToString()) as IComparable;
                }
                else
                {
                    c2 = Convert.ChangeType(c2, c1.GetType()) as IComparable;
                }
            }

            // compare
            return c1.CompareTo(c2);
        }

        private object GetAsUnderlyingType(Enum enval)
        {
            Type entype = enval.GetType();

            Type undertype = Enum.GetUnderlyingType(entype);

            return Convert.ChangeType(enval, undertype);
        }

        #endregion // IComparable<Expression>

        //---------------------------------------------------------------------------

        /// <summary>
        /// Creates the sub context calculate engine and initializes its variables, function and 
        /// ReInitSubContextVariables property using the parent calc engine.
        /// </summary>
        /// <param name="exp">Variable or Binding Expression</param>
        /// <returns></returns>
        public CalcEngine CreateSubContextCalcEngine()
        {
            var ce = new CalcEngine();
            var pce = GetParentCalcEngine(this);
            if (pce != null)
            {
                ce.Variables = pce.Variables;
                ce.Functions = pce.Functions;
                ce.ReInitSubContextVariables = pce.ReInitSubContextVariables;
                ce.ReturnChangedCollection = pce.ReturnChangedCollection;
            }
            return ce;
        }

        private CalcEngine GetParentCalcEngine(Expression exp)
        {
            if ((exp as BindingExpression) != null)
                return (exp as BindingExpression)._ce;
            else if ((exp as VariableExpression) != null)
                return (exp as VariableExpression)._ce;
            else if ((exp as FunctionExpression) != null)
                return (exp as FunctionExpression)._ce;
            return null;
        }
    }

    /// <summary>
    /// Unary expression, e.g. +123
    /// </summary>
    internal class UnaryExpression : Expression
    {
        // fields
        private Expression _expr;

        // ctor
        public UnaryExpression(Token tk, Expression expr)
            : base(tk)
        {
            _expr = expr;
        }

        // object model
        override public object Evaluate()
        {
            switch (_token.ID)
            {
                case TKID.ADD:
                    return +(double)_expr;
                case TKID.SUB:
                    return -(double)_expr;
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            _expr = _expr.Optimize();
            return _expr._token.Type == TKTYPE.LITERAL ? new Expression(this.Evaluate()) : this;
        }

        public override List<string> GetBindings()
        {
            List<string> bindings = new List<string>();
            AddBindingRange(bindings, _expr);
            return bindings;
        }

        public override bool IsLiteral
        {
            get { return _expr._token.Type == TKTYPE.LITERAL; }
        }
    }

    /// <summary>
    /// Binary expression, e.g. 1+2
    /// </summary>
    internal class BinaryExpression : Expression
    {
        // fields
        private Expression _lft;

        private Expression _rgt;

        // ctor
        public BinaryExpression(Token tk, Expression exprLeft, Expression exprRight)
            : base(tk)
        {
            _lft = exprLeft;
            _rgt = exprRight;
        }

        // object model
        override public object Evaluate()
        {
            // handle comparisons
            if (_token.Type == TKTYPE.COMPARE)
            {
                var cmp = _lft.CompareTo(_rgt);
                switch (_token.ID)
                {
                    case TKID.GT: return cmp > 0;
                    case TKID.LT: return cmp < 0;
                    case TKID.GE: return cmp >= 0;
                    case TKID.LE: return cmp <= 0;
                    case TKID.EQ: return cmp == 0;
                    case TKID.NE: return cmp != 0;
                }
            }

            // handle everything else
            switch (_token.ID)
            {
                case TKID.ADD:
                    return (double)_lft + (double)_rgt;
                case TKID.SUB:
                    return (double)_lft - (double)_rgt;
                case TKID.MUL:
                    return (double)_lft * (double)_rgt;
                case TKID.DIV:
                    return (double)_lft / (double)_rgt;
                case TKID.DIVINT:
                    return (double)(int)((double)_lft / (double)_rgt);
                case TKID.MOD:
                    return (double)(int)((double)_lft % (double)_rgt);
                case TKID.POWER:
                    var a = (double)_lft;
                    var b = (double)_rgt;
                    if (b == 0.0) return 1.0;
                    if (b == 0.5) return System.Math.Sqrt(a);
                    if (b == 1.0) return a;
                    if (b == 2.0) return a * a;
                    if (b == 3.0) return a * a * a;
                    if (b == 4.0) return a * a * a * a;
                    return System.Math.Pow((double)_lft, (double)_rgt);
            }
            throw new ArgumentException("Bad expression.");
        }

        public override Expression Optimize()
        {
            _lft = _lft.Optimize();
            _rgt = _rgt.Optimize();
            return _lft._token.Type == TKTYPE.LITERAL && _rgt._token.Type == TKTYPE.LITERAL ? new Expression(this.Evaluate()) : this;
        }

        public override List<string> GetBindings()
        {
            List<string> bindings = new List<string>();
            AddBindingRange(bindings, _lft);
            AddBindingRange(bindings, _rgt);
            return bindings;
        }
    }

    /// <summary>
    /// Function call expression, e.g. sin(0.5)
    /// </summary>
    internal class FunctionExpression : Expression
    {
        // fields
        public CalcEngine _ce; // Changed to public so that it can be used inside functions
        private FunctionDefinition _fn;

        private List<Expression> _parms;

        // ctor
        internal FunctionExpression()
        {
        }

        public FunctionExpression(CalcEngine engine, FunctionDefinition function, List<Expression> parms)
        {
            _ce = engine;
            _fn = function;
            _parms = parms;
        }

        // object model
        override public object Evaluate()
        {
            object result = null;
            string DYNAMIC_CODE_DEFAULT_NAME = "__dynCode__";
            string DEFAULT_DYNAMIC_CODE_NAME_FORMAT = @"^((?<name>({0}))+(?<num>([\d]\d*))?)?$";

            if (string.IsNullOrEmpty(_fn.DynamicCode))
            {
                result = _fn.Function(_parms);
            }
            else
            {
                // Initialize property name to next default variable name
                List<string> existingNames = _ce.Variables.Keys.Where(k => k.Contains(DYNAMIC_CODE_DEFAULT_NAME)).ToList();
                string dynCodeName = DYNAMIC_CODE_DEFAULT_NAME.GenerateVarName(DEFAULT_DYNAMIC_CODE_NAME_FORMAT, existingNames);
                // There could be no parameters; therefore, make sure it's valid anyway
                if (_parms == null)
                    _parms = new List<Expression>();
                // Add required variables
                _ce.Variables.Add(dynCodeName, _fn.DynamicCode);
                for (int i = 0; _fn.DynamicParameters != null && i < _fn.DynamicParameters.Count; i++)
                    _ce.Variables.Add(_fn.DynamicParameters[i], _parms[i].Evaluate());

                // Pass code into function as varible so that we have the full CalcEngine context
                _parms.Clear();
                _parms.Add(new VariableExpression(_ce, _ce.Variables, dynCodeName));
                result = _fn.Function(_parms);
                // Remove added variables
                for (int i = 0; _fn.DynamicParameters != null && i < _fn.DynamicParameters.Count; i++)
                    _ce.Variables.Remove(_fn.DynamicParameters[i]);
                _ce.Variables.Remove(dynCodeName);
            }
            return result;
        }

        public override Expression Optimize()
        {
            bool allLits = true;
            if (_parms != null)
            {
                for (int i = 0; i < _parms.Count; i++)
                {
                    var p = _parms[i].Optimize();
                    _parms[i] = p;
                    if (p._token.Type != TKTYPE.LITERAL)
                    {
                        allLits = false;
                    }
                }
                // If "RegisterFunction", then add CalcEngine to parms so that the 
                // named function can be registered with-in this instance of the CalcEngine.
                var mInfo = _fn.Function.GetMethodInfo();
                if (mInfo.Name  == "RegisterFunction")
                {
                    _parms.Add(this);
                }
            }
            return allLits ? new Expression(this.Evaluate()) : this;
        }

        public override List<string> GetBindings()
        {
            List<string> bindings = new List<string>();
            if (_parms != null)
            {
                foreach (Expression e in _parms)
                {
                    AddBindingRange(bindings, e);
                }
            }
            return bindings;
        }
    }

    /// <summary>
    /// Simple variable reference.
    /// </summary>
    internal class VariableExpression : Expression
    {
        public CalcEngine _ce; // Changed to public so that it can be used inside functions
        private IDictionary<string, object> _dct;
        private string _name;

        public VariableExpression(CalcEngine engine, IDictionary<string, object> dct, string name)
        {
            _ce = engine;
            _dct = dct;
            _name = name;
        }

        public override object Evaluate()
        {
            return _dct[_name];
        }

        public override List<string> GetBindings()
        {
            List<string> bindings = new List<string>();
            bindings.Add(_name);
            return bindings;
        }
    }

    /// <summary>
    /// Expression based on an object's properties.
    /// </summary>
    internal class BindingExpression : Expression
    {
        public CalcEngine _ce; // Changed to public so that it can be used inside functions
        private CultureInfo _ci;
        private List<BindingInfo> _bindingPath;

        // ctor
        internal BindingExpression(CalcEngine engine, List<BindingInfo> bindingPath, CultureInfo ci)
        {
            _ce = engine;
            _bindingPath = bindingPath;
            _ci = ci;
        }

        // object model
        override public object Evaluate()
        {
            return GetValue(_ce.DataContext);
        }

        // implementation
        private object GetValue(object obj)
        {
            const BindingFlags bf = BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.Public | BindingFlags.Static;

            if (obj != null)
            {
                foreach (var bi in _bindingPath)
                {
                    // get property
                    if (bi.PropertyInfo == null)
                    {
                        bi.PropertyInfo = obj.GetType().GetProperty(bi.Name, bf);
                    }

                    // get object
                    try
                    {
                        obj = bi.PropertyInfo.GetValue(obj, null);
                    }
                    catch
                    {
                        // This could happed when processing the following functions: ExecuteExpr, EvalExprIf, Switch
                        string errMsg = string.Format("Error: '{0}' property does not exist on object '{1}'", bi.Name, obj.GetType().Name);
                        bi.PropertyInfo = obj.GetType().GetProperty(bi.Name, bf);
                        bi.PropertyInfoItem = null;
                        if (bi.PropertyInfo != null)
                        obj = bi.PropertyInfo.GetValue(obj, null);
                        else
                            throw new Exception(errMsg);
                    }

                    // handle indexers (lists and dictionaries)
                    if (bi.Parms != null && bi.Parms.Count > 0)
                    {
                        // get indexer property (always called "Item")
                        if (bi.PropertyInfoItem == null)
                        {
                            bi.PropertyInfoItem = obj.GetType().GetProperty("Item", bf);
                        }

                        // get indexer parameters
                        var pip = bi.PropertyInfoItem.GetIndexParameters();
                        var list = new List<object>();
                        for (int i = 0; i < pip.Length; i++)
                        {
                            var pv = bi.Parms[i].Evaluate();
                            pv = Convert.ChangeType(pv, pip[i].ParameterType, _ci);
                            list.Add(pv);
                        }

                        // get value
                        obj = bi.PropertyInfoItem.GetValue(obj, list.ToArray());
                    }
                }
            }

            // all done
            return obj;
        }

        //ETM_RESOLVE: Need to resolve issues with full path bindings: allow N circular references and signaling property changes using full path
        private static bool UseFullBindingPath = true;

        public override List<string> GetBindings()
        {
            List<string> bindings = new List<string>();

            if (UseFullBindingPath)
                bindings.Add(string.Join(".", _bindingPath.Select(b => b.Name).ToArray()));

            foreach (BindingInfo bi in _bindingPath)
            {
                if (!UseFullBindingPath)
                    bindings.Add(bi.Name);

                if (bi.Parms != null)
                {
                    foreach (Expression e in bi.Parms)
                    {
                        AddBindingRange(bindings, e);
                    }
                }
            }
            return bindings;
        }
    }

    /// <summary>
    /// Helper used for building BindingExpression objects.
    /// </summary>
    internal class BindingInfo
    {
        public BindingInfo(string member, List<Expression> parms)
        {
            Name = member;
            Parms = parms;
        }

        public string Name { get; set; }

        public PropertyInfo PropertyInfo { get; set; }

        public PropertyInfo PropertyInfoItem { get; set; }

        public List<Expression> Parms { get; set; }
    }

    /// <summary>
    /// Expression that represents an external object.
    /// </summary>
    internal class XObjectExpression : Expression, IEnumerable
    {
        private object _value;

        // ctor
        internal XObjectExpression(object value)
        {
            _value = value;
        }

        // object model
        public override object Evaluate()
        {
            // use IValueObject if available
            var iv = _value as IValueObject;
            if (iv != null)
            {
                return iv.GetValue();
            }

            // return raw object
            return _value;
        }

        public IEnumerator GetEnumerator()
        {
            var ie = _value as IEnumerable;
            return ie != null ? ie.GetEnumerator() : null;
        }
    }

    /// <summary>
    /// Interface supported by external objects that have to return a value
    /// other than themselves (e.g. a cell range object should return the
    /// cell content instead of the range itself).
    /// </summary>
    public interface IValueObject
    {
        object GetValue();
    }
}