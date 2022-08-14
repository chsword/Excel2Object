namespace Chsword.Excel2Object.Functions;

public interface IStatisticsFunction
{
    ColumnValue Sum(params ColumnMatrix[] matrix);
}