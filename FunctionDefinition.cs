using System.Collections.Generic;

namespace CalcEngine
{
    /// <summary>
    /// Function definition class (keeps function name, parameter counts, and delegate).
    /// </summary>
    public class FunctionDefinition
    {
        // fields
        public int ParmMin, ParmMax;
        public string DynamicCode;
        public List<string> DynamicParameters;

        public CalcEngineFunction Function;

        // ctor
        public FunctionDefinition(int parmMin, int parmMax, CalcEngineFunction function)
        {
            ParmMin = parmMin;
            ParmMax = parmMax;
            Function = function;
        }

        public FunctionDefinition(List<string> paramters, string code, CalcEngineFunction function)
        {
            DynamicParameters = paramters == null ? new List<string>() : paramters;
            ParmMin = ParmMax = paramters.Count;
            Function = function;
            DynamicCode = code;
        }
    }
}