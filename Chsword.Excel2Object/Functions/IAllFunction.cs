namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Every supported Excel function in one place, for formulas that mix categories.
/// </summary>
public interface IAllFunction : IMathFunction, IStatisticsFunction, IConditionFunction, IReferenceFunction,
    IDateTimeFunction, ITextFunction, IInformationFunction, IFinancialFunction, IEngineeringFunction,
    IDatabaseFunction;
