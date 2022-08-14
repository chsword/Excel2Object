namespace Chsword.Excel2Object.Functions;

public interface IMathFunction
{
    int Abs(object val);

    /// <summary>
    ///     正数向上取整至偶数，负数向下取整至偶数
    /// </summary>
    int Even(object val);

    //阶fact
    int Fact(object val);

    /// <summary>
    ///     To integer
    /// </summary>
    int Int(object val);

    double PI();

    /// <summary>
    ///     random 0-1
    /// </summary>
    double Rand();

    int Round(object val, object digits);
    int RoundDown(object val, object digits);
    int RoundUp(object val, object digits);
    int Sqrt(object val);
}