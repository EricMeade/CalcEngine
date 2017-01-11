using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace CalcEngine
{
    public static class Extensions
    {
        #region Public Methods

        /// <summary>
        /// Returns an object from the specified property from the specified object.
        /// </summary>
        /// <param name="source">Source object to fetch the property from.</param>
        /// <param name="propertyName">The name of the property to fetch from the source.</param>
        /// <returns>Object containing the property value of the specified property name from the specified source.</returns>
        public static object GetPropertyValue(this object source, string propertyName)
        {
            if (source == null || string.IsNullOrEmpty(propertyName)) return null;

            object value = source;
            foreach (string property in propertyName.Split(".".ToArray()))
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(value);
                value = properties[propertyName].GetValue(value);
            }
            return value;
        }
        public static T GetValue<T>(this object target, System.Linq.Expressions.Expression<Func<T>> fieldAccessExpression)
        {
            System.Linq.Expressions.Expression bodyExpression = fieldAccessExpression.Body;
            MemberExpression memberExpression = bodyExpression as MemberExpression;
            FieldInfo fieldInfo = memberExpression.Member as FieldInfo;
            return (T)fieldInfo.GetValue(target);
        }

        public static T GetValue<T>(this object target, string propertyName)
        {
            BindingFlags bindings = BindingFlags.Public | BindingFlags.Static | BindingFlags.GetField | BindingFlags.GetProperty | BindingFlags.Instance;
            FieldInfo fieldInfo = target.GetType().GetField(propertyName, bindings);
            return (T)fieldInfo.GetValue(target);
        }

        public static object GetValue(this object target, string propertyName)
        {
            //PropertyInfo pInfo = target.GetType().GetProperty(propertyName);
            //return pInfo.GetValue(target, null);

            if (target == null || string.IsNullOrEmpty(propertyName)) return null;

            //ETM_UPDATE: Other GetValue extensions to iterate property path
            string[] properties = propertyName.Split(".".ToArray());
            object value = target;
            foreach (string property in properties)
            {
                PropertyInfo pInfo = value.GetType().GetProperty(property);
                value = pInfo.GetValue(value, null);
            }
            return value;

            //BindingFlags bindings = BindingFlags.Public | BindingFlags.Static | BindingFlags.GetField | BindingFlags.GetProperty | BindingFlags.Instance;
            //FieldInfo fieldInfo = target.GetType().GetField(propertyName, bindings);
            //return fieldInfo.GetValue(target);
        }

        public static void SetValue<T>(this object target, System.Linq.Expressions.Expression<Func<T>> fieldAccessExpression, T value)
        {
            System.Linq.Expressions.Expression bodyExpression = fieldAccessExpression.Body;
            MemberExpression memberExpression = bodyExpression as MemberExpression;
            FieldInfo fieldInfo = memberExpression.Member as FieldInfo;
            fieldInfo.SetValue(target, value);
        }

        public static void SetValue<T>(this object target, string propertyName, T value)
        {
            BindingFlags bindings = BindingFlags.Public | BindingFlags.Static | BindingFlags.SetField | BindingFlags.SetProperty | BindingFlags.Instance;
            FieldInfo fieldInfo = target.GetType().GetField(propertyName, bindings);
            fieldInfo.SetValue(target, value);
        }

        public static T To<T>(this IConvertible obj)
        {
            Type t = typeof(T);
            Type u = Nullable.GetUnderlyingType(t);

            if (u != null)
            {
                if (obj == null)
                    return default(T);

                return (T)Convert.ChangeType(obj, u);
            }
            else
            {
                return (T)Convert.ChangeType(obj, t);
            }
        }

        public static object To(this IConvertible value, Type type)
        {
            Type underlyingType = Nullable.GetUnderlyingType(type);

            if (underlyingType != null)
            {
                if (value == null)
                    return default(object);

                return Convert.ChangeType(value, underlyingType);
            }
            else
            {
                return Convert.ChangeType(value, type);
            }
        }
        public static void SetValue(this object target, string propertyName, object value)
        {
            //PropertyInfo pInfo = target.GetType().GetProperty(propertyName);
            //Type uType = Nullable.GetUnderlyingType(pInfo.PropertyType);
            //var convertedValue = (uType != null) ? ((value == null) ? GetDefault(pInfo.PropertyType) : Convert.ChangeType(value, uType)) : Convert.ChangeType(value, pInfo.PropertyType);
            //pInfo.SetValue(target, convertedValue, null);

            //ETM_UPDATE: Other SetValue extensions to iterate property path
            PropertyInfo pInfo = null;
            string[] properties = propertyName.Split(".".ToArray());
            object currentTarget = target;
            for (int i = 0; i < properties.Count(); i++)
            {
                pInfo = currentTarget.GetType().GetProperty(properties[i]);
                if (i < (properties.Count() - 1))
                    currentTarget = pInfo.GetValue(currentTarget, null);
            }
            Type uType = Nullable.GetUnderlyingType(pInfo.PropertyType);
            var convertedValue = (uType != null) ? ((value == null) ? GetDefault(pInfo.PropertyType) : Convert.ChangeType(value, uType)) : Convert.ChangeType(value, pInfo.PropertyType);
            pInfo.SetValue(currentTarget, convertedValue, null);

            //BindingFlags bindings = BindingFlags.Public | BindingFlags.Static | BindingFlags.SetField | BindingFlags.SetProperty | BindingFlags.Instance;
            //FieldInfo fieldInfo = target.GetType().GetField(propertyName, bindings);
            //fieldInfo.SetValue(target, value);
        }

        public static string GetMemberName<T>(this T instance, System.Linq.Expressions.Expression<Func<T, object>> expression)
        {
            return GetMemberName(expression);
        }

        public static string GetMemberName<T>(System.Linq.Expressions.Expression<Func<T, object>> expression)
        {
            if (expression == null)
            {
                throw new ArgumentException("The expression cannot be null.");
            }

            return GetMemberName(expression.Body);
        }

        public static string GetMemberName<T>(this T instance, System.Linq.Expressions.Expression<Action<T>> expression)
        {
            return GetMemberName(expression);
        }

        public static string GetMemberName<T>(System.Linq.Expressions.Expression<Action<T>> expression)
        {
            if (expression == null)
            {
                throw new ArgumentException("The expression cannot be null.");
            }

            return GetMemberName(expression.Body);
        }

        public static IEnumerable<MethodInfo> GetExtensionMethods(Assembly assembly, Type extendedType)
        {
            var query = from type in assembly.GetTypes()
                        where type.IsSealed && !type.IsGenericType && !type.IsNested
                        from method in type.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                        //Uncomment line below when refference to assembly containing ExtensionAttribute can be added
                        //where method.IsDefined(typeof(ExtensionAttribute), false)
                        where method.GetParameters()[0].ParameterType == extendedType
                        select method;
            return query;
        }

        public static object GetDefault(Type type)
        {
           if(type.IsValueType)
           {
              return Activator.CreateInstance(type);
           }
           return null;
        }

        /// <summary>
        /// Extension method to create a deep copy of any type.
        /// </summary>
        /// <typeparam name="T">The type</typeparam>
        /// <param name="objectToCopy">The object to copy.</param>
        /// <returns>The copied object of the type T.</returns>
        public static T DeepCopy<T>(this T objectToCopy)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(memoryStream, objectToCopy);
                memoryStream.Position = 0;
                return (T)binaryFormatter.Deserialize(memoryStream);
            }
        }
        #endregion Public Methods

        #region Private Methods

        private static string GetMemberName(System.Linq.Expressions.Expression expression)
        {
            if (expression == null)
            {
                throw new ArgumentException("The expression cannot be null.");
            }

            if (expression is MemberExpression)
            {
                // Reference type property or field
                var memberExpression = (MemberExpression)expression;
                return memberExpression.Member.Name;
            }

            if (expression is MethodCallExpression)
            {
                // Reference type method
                var methodCallExpression = (MethodCallExpression)expression;
                return methodCallExpression.Method.Name;
            }

            if (expression is System.Linq.Expressions.UnaryExpression)
            {
                // Property, field of method returning value type
                var unaryExpression = (System.Linq.Expressions.UnaryExpression)expression;
                return GetMemberName(unaryExpression);
            }

            throw new ArgumentException("Invalid expression");
        }

        private static string GetMemberName(System.Linq.Expressions.UnaryExpression unaryExpression)
        {
            if (unaryExpression.Operand is MethodCallExpression)
            {
                var methodExpression = (MethodCallExpression)unaryExpression.Operand;
                return methodExpression.Method.Name;
            }

            return ((MemberExpression)unaryExpression.Operand).Member.Name;
        }

        #endregion Private Methods

        #region // Private Calculation Methods To Auto Generate Variable Name

        /// <summary>
        /// Generates the next name based on the source matching the regex. The source is returned if no match is found.
        /// The output string is built in the order of the group names in the regex. The group name are processed as follows:
        ///     num    = If found the number will be incremented; otherwise, if the num group is not found and AppendNumberIfNotExist 
        ///              is true, then a number will be appened to the generated name.
        ///     suffix = If the suffix exist (st,nd,rd,th,), then a new suffix will be generated based on the incremented number 
        ///              from the num group.
        ///     prefix = Seperator for parts of the generated name.
        ///     name   = Constant string portion of the generated name.
        ///     option = If found, will be dropped from the newly generated name.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="regexExpression">The regex expression.</param>
        /// <param name="alwaysIncrementNumber">if set to <c>true</c> the "num" group will always be incremented even if it is not alresy used in the existing names.</param>
        /// <param name="appendNumberIfNotExist">if set to <c>true</c> [append number if not exist].</param>
        /// <param name="existingNames">The list of existing names which the newly generated name cannot duplicate.</param>
        /// <param name="parameters">The parameters.</param>
        /// <returns></returns>
        public static string GenerateVarName(this string source, string regexExpression, IEnumerable<string> existingNames = null)
        {
            string outputString = source;
            string[] groupNames = { "num", "name" };
            NameValueCollection regExGroupValues = new NameValueCollection();
            bool foundNumber = false;
            int numberGroupNameIndex = -1;
            string nextNumber = string.Empty;
            regexExpression = string.Format(regexExpression, source);

            Regex re = new Regex(regexExpression, RegexOptions.IgnoreCase); //ETM_REMOVE: | RegexOptions.IgnorePatternWhitespace);
            Match match = re.Match(source);
            if (match.Captures.Count == 1)
            {
                // Get regex group vaules
                for (int gIdx = 0; gIdx < match.Groups.Count; gIdx++)
                {
                    if (groupNames.Contains(re.GetGroupNames()[gIdx]))
                        regExGroupValues.Add(re.GetGroupNames()[gIdx], match.Groups[gIdx].Value);
                }

                // Get existing numbera from existing names
                var existingNumbers = GetRegExGroupByName(existingNames, regexExpression, "num").Select(g => { int n = 0; return (!string.IsNullOrEmpty(g.Trim()) && int.TryParse(g, out n)) ? n : 0; });

                // Increment counter
                if (regExGroupValues.AllKeys.Contains("num"))
                {
                    foundNumber = true;
                    if (!string.IsNullOrEmpty(regExGroupValues["num"].Trim()))
                    {
                        var currentNumber = Convert.ToInt32(regExGroupValues["num"].Trim());
                        currentNumber++; // Always increment number
                        while (existingNumbers.Contains(currentNumber))
                            currentNumber++;
                        regExGroupValues["num"] = currentNumber.ToString();
                    }
                    else
                    {
                        // Increment max number from existing names
                        regExGroupValues["num"] = ((existingNumbers.Count() != 0) ? existingNumbers.Max() + 1 : 1).ToString();
                    }
                }
                // If num group not found and using num in the expression, then get next number from existing names
                if (!foundNumber)
                {
                    List<string> groups = GetRegExGroupNameOrder(regexExpression, groupNames);
                    if (groups.Contains("num"))
                    {
                        // Increment max number from existing names
                        nextNumber = ((existingNumbers.Count() != 0) ? existingNumbers.Max() + 1 : 1).ToString();
                        numberGroupNameIndex = groups.IndexOf("num");
                    }
                }

                // Build new name so that number can be inserted
                outputString = string.Empty;
                for (int index = 0; index < regExGroupValues.Keys.Count; index++)
                {
                    if (index == numberGroupNameIndex)
                        outputString += nextNumber;
                    outputString += regExGroupValues[index];
                }
            }

            if (existingNames != null && existingNames.Count() != 0 && existingNames.Select(n => n.ToUpperInvariant()).Contains(outputString.ToUpperInvariant()))
            {
                int layerSuffixId = 0;
                // Get next number suffix
                foreach (string name in existingNames)
                {
                    int number = 0;
                    string[] split2 = name.Split('_');
                    if (split2.Length > 1)
                        int.TryParse(split2[split2.Length - 1], out number);
                    if (number > layerSuffixId)
                        layerSuffixId = number;
                }
                outputString = string.Format("{0}_{1}", outputString, layerSuffixId + 1);
            }

            return outputString;
        }

        private static List<string> GetRegExGroupNameOrder(string regexExpression, string[] groupNames)
        {
            string[] regexSeperators = { @"(?<", @">(" };
            List<string> list = new List<string>();
            if (!string.IsNullOrEmpty(regexExpression) && groupNames != null)
            {
                foreach (string s in regexExpression.Split(regexSeperators, StringSplitOptions.None))
                {
                    if (groupNames.Contains(s))
                        list.Add(s);
                }
            }
            return list;
        }

        private static string GetRegExGroupValue(string regexExpression, string groupName)
        {
            string[] regexSeperators = { @"(?<", @">(" };
            string value = string.Empty;
            if (!string.IsNullOrEmpty(regexExpression) && groupName != null)
            {
                bool processNext = false;
                foreach (string s in regexExpression.Split(regexSeperators, StringSplitOptions.None))
                {
                    if (processNext == true)
                    {
                        string[] groupSeperators = { @"(", @"+", @"?", @"\", @")" };
                        value = string.Join(string.Empty, s.Split(groupSeperators, StringSplitOptions.None));
                        break;
                    }
                    if (s == groupName)
                        processNext = true;
                }
            }
            return value;
        }

        private static List<string> GetRegExGroupByName(IEnumerable<string> collection, string regexExpression, string groupName)
        {
            List<string> groupValues = new List<string>();

            if (collection != null && collection.Count() != 0)
            {
                Regex re = new Regex(regexExpression, RegexOptions.IgnoreCase);
                foreach (string str in collection)
                {
                    Match match = re.Match(str);
                    if (match.Captures.Count == 1)
                    {
                        // Get regex group vaules
                        for (int gIdx = 0; gIdx < match.Groups.Count; gIdx++)
                        {
                            if (groupName == re.GetGroupNames()[gIdx])
                                groupValues.Add(match.Groups[gIdx].Value);
                        }
                    }
                }
            }
            return groupValues;
        }

        #endregion // Private Calculation Methods To Auto Generate Variable Name
    }
}
