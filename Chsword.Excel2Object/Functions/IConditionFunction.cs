namespace Chsword.Excel2Object.Functions; 

public interface IConditionFunction
{
    ColumnValue If(ColumnValue condition, ColumnValue value1, ColumnValue value2);
}