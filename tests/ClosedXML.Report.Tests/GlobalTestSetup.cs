namespace ClosedXML.Report.Tests
{
    class GlobalTestSetup
    {
        public GlobalTestSetup()
        {
            //Warm-up DynamicExpressionParser
            var evaluator = new FormulaEvaluator();
            evaluator.AddVariable("item", "warm-up");
            evaluator.Evaluate("&={{item.Length}}");
        }
    }
}
