namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel statistical functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
/// <remarks>
///     Ranges are passed as <see cref="ColumnMatrix" /> (for example <c>c.Matrix("One", 1, "Two", 9)</c>);
///     an implicit conversion lets a range be used anywhere a <see cref="ColumnValue" /> is expected, so
///     values and ranges can be mixed freely in the argument lists below.
/// </remarks>
public interface IStatisticsFunction : IExcelFunction
{
    // ---- sums, counts and averages ----

    /// <summary>SUM - adds its arguments.</summary>
    ColumnValue Sum(params ColumnValue[] values);

    /// <summary>AVERAGE - arithmetic mean of its arguments.</summary>
    ColumnValue Average(params ColumnValue[] values);

    /// <summary>AVERAGEA - arithmetic mean, counting text and logical values.</summary>
    [ExcelFunctionName("AVERAGEA")]
    ColumnValue AverageA(params ColumnValue[] values);

    /// <summary>AVERAGEIF - mean of the cells of a range that meet a criteria.</summary>
    ColumnValue AverageIf(ColumnMatrix range, ColumnValue criteria);

    /// <summary>AVERAGEIF - mean of <paramref name="averageRange" /> where <paramref name="range" /> meets a criteria.</summary>
    ColumnValue AverageIf(ColumnMatrix range, ColumnValue criteria, ColumnMatrix averageRange);

    /// <summary>AVERAGEIFS - mean of the cells that meet several criteria; pass criteria as range, criteria pairs.</summary>
    ColumnValue AverageIfs(ColumnMatrix averageRange, params ColumnValue[] criteria);

    /// <summary>COUNT - counts the numeric cells.</summary>
    ColumnValue Count(params ColumnValue[] values);

    /// <summary>COUNTA - counts the cells that are not empty.</summary>
    [ExcelFunctionName("COUNTA")]
    ColumnValue CountA(params ColumnValue[] values);

    /// <summary>COUNTBLANK - counts the empty cells of a range.</summary>
    ColumnValue CountBlank(ColumnMatrix range);

    /// <summary>COUNTIF - counts the cells of a range that meet a criteria.</summary>
    ColumnValue CountIf(ColumnMatrix range, ColumnValue criteria);

    /// <summary>COUNTIFS - counts the cells that meet several criteria; pass criteria as range, criteria pairs.</summary>
    ColumnValue CountIfs(params ColumnValue[] criteria);

    // ---- extremes and order ----

    /// <summary>MAX - largest value.</summary>
    ColumnValue Max(params ColumnValue[] values);

    /// <summary>MAXA - largest value, counting text and logical values.</summary>
    [ExcelFunctionName("MAXA")]
    ColumnValue MaxA(params ColumnValue[] values);

    /// <summary>MAXIFS - largest value among the cells that meet several criteria.</summary>
    [ExcelFunctionName("MAXIFS", Future = true)]
    ColumnValue MaxIfs(ColumnMatrix maxRange, params ColumnValue[] criteria);

    /// <summary>MIN - smallest value.</summary>
    ColumnValue Min(params ColumnValue[] values);

    /// <summary>MINA - smallest value, counting text and logical values.</summary>
    [ExcelFunctionName("MINA")]
    ColumnValue MinA(params ColumnValue[] values);

    /// <summary>MINIFS - smallest value among the cells that meet several criteria.</summary>
    [ExcelFunctionName("MINIFS", Future = true)]
    ColumnValue MinIfs(ColumnMatrix minRange, params ColumnValue[] criteria);

    /// <summary>LARGE - the k-th largest value of a range.</summary>
    ColumnValue Large(ColumnMatrix array, ColumnValue k);

    /// <summary>SMALL - the k-th smallest value of a range.</summary>
    ColumnValue Small(ColumnMatrix array, ColumnValue k);

    /// <summary>RANK - rank of a number within a range.</summary>
    ColumnValue Rank(ColumnValue number, ColumnMatrix reference);

    /// <summary>RANK - rank of a number within a range, ascending when <paramref name="order" /> is non zero.</summary>
    ColumnValue Rank(ColumnValue number, ColumnMatrix reference, ColumnValue order);

    /// <summary>RANK.EQ - rank of a number, ties sharing the highest rank.</summary>
    [ExcelFunctionName("RANK.EQ", Future = true)]
    ColumnValue RankEq(ColumnValue number, ColumnMatrix reference, ColumnValue order);

    /// <summary>RANK.AVG - rank of a number, ties sharing the average rank.</summary>
    [ExcelFunctionName("RANK.AVG", Future = true)]
    ColumnValue RankAvg(ColumnValue number, ColumnMatrix reference, ColumnValue order);

    /// <summary>PERCENTRANK - percentage rank of a value within a range.</summary>
    [ExcelFunctionName("PERCENTRANK")]
    ColumnValue PercentRank(ColumnMatrix array, ColumnValue x);

    /// <summary>PERCENTILE - the k-th percentile of a range.</summary>
    ColumnValue Percentile(ColumnMatrix array, ColumnValue k);

    /// <summary>PERCENTILE.INC - percentile, inclusive of 0 and 1.</summary>
    [ExcelFunctionName("PERCENTILE.INC", Future = true)]
    ColumnValue PercentileInc(ColumnMatrix array, ColumnValue k);

    /// <summary>PERCENTILE.EXC - percentile, exclusive of 0 and 1.</summary>
    [ExcelFunctionName("PERCENTILE.EXC", Future = true)]
    ColumnValue PercentileExc(ColumnMatrix array, ColumnValue k);

    /// <summary>QUARTILE - the quartile of a range.</summary>
    ColumnValue Quartile(ColumnMatrix array, ColumnValue quart);

    /// <summary>QUARTILE.INC - quartile, inclusive of 0 and 1.</summary>
    [ExcelFunctionName("QUARTILE.INC", Future = true)]
    ColumnValue QuartileInc(ColumnMatrix array, ColumnValue quart);

    /// <summary>QUARTILE.EXC - quartile, exclusive of 0 and 1.</summary>
    [ExcelFunctionName("QUARTILE.EXC", Future = true)]
    ColumnValue QuartileExc(ColumnMatrix array, ColumnValue quart);

    // ---- central tendency and spread ----

    /// <summary>MEDIAN - median of its arguments.</summary>
    ColumnValue Median(params ColumnValue[] values);

    /// <summary>MODE - most frequent value.</summary>
    ColumnValue Mode(params ColumnValue[] values);

    /// <summary>MODE.SNGL - most frequent value.</summary>
    [ExcelFunctionName("MODE.SNGL", Future = true)]
    ColumnValue ModeSngl(params ColumnValue[] values);

    /// <summary>MODE.MULT - the most frequent values, as an array.</summary>
    [ExcelFunctionName("MODE.MULT", Future = true)]
    ColumnValue ModeMult(params ColumnValue[] values);

    /// <summary>GEOMEAN - geometric mean.</summary>
    [ExcelFunctionName("GEOMEAN")]
    ColumnValue GeoMean(params ColumnValue[] values);

    /// <summary>HARMEAN - harmonic mean.</summary>
    [ExcelFunctionName("HARMEAN")]
    ColumnValue HarMean(params ColumnValue[] values);

    /// <summary>TRIMMEAN - mean of the interior of a data set.</summary>
    [ExcelFunctionName("TRIMMEAN")]
    ColumnValue TrimMean(ColumnMatrix array, ColumnValue percent);

    /// <summary>AVEDEV - average of the absolute deviations from the mean.</summary>
    [ExcelFunctionName("AVEDEV")]
    ColumnValue AveDev(params ColumnValue[] values);

    /// <summary>DEVSQ - sum of the squares of the deviations from the mean.</summary>
    [ExcelFunctionName("DEVSQ")]
    ColumnValue DevSq(params ColumnValue[] values);

    /// <summary>STDEV - sample standard deviation.</summary>
    [ExcelFunctionName("STDEV")]
    ColumnValue StDev(params ColumnValue[] values);

    /// <summary>STDEV.S - sample standard deviation.</summary>
    [ExcelFunctionName("STDEV.S", Future = true)]
    ColumnValue StDevS(params ColumnValue[] values);

    /// <summary>STDEV.P - population standard deviation.</summary>
    [ExcelFunctionName("STDEV.P", Future = true)]
    ColumnValue StDevP(params ColumnValue[] values);

    /// <summary>STDEVA - sample standard deviation, counting text and logical values.</summary>
    [ExcelFunctionName("STDEVA")]
    ColumnValue StDevA(params ColumnValue[] values);

    /// <summary>STDEVPA - population standard deviation, counting text and logical values.</summary>
    [ExcelFunctionName("STDEVPA")]
    ColumnValue StDevPA(params ColumnValue[] values);

    /// <summary>VAR - sample variance.</summary>
    [ExcelFunctionName("VAR")]
    ColumnValue Var(params ColumnValue[] values);

    /// <summary>VAR.S - sample variance.</summary>
    [ExcelFunctionName("VAR.S", Future = true)]
    ColumnValue VarS(params ColumnValue[] values);

    /// <summary>VAR.P - population variance.</summary>
    [ExcelFunctionName("VAR.P", Future = true)]
    ColumnValue VarP(params ColumnValue[] values);

    /// <summary>VARA - sample variance, counting text and logical values.</summary>
    [ExcelFunctionName("VARA")]
    ColumnValue VarA(params ColumnValue[] values);

    /// <summary>VARPA - population variance, counting text and logical values.</summary>
    [ExcelFunctionName("VARPA")]
    ColumnValue VarPA(params ColumnValue[] values);

    /// <summary>SKEW - skewness of a distribution.</summary>
    ColumnValue Skew(params ColumnValue[] values);

    /// <summary>KURT - kurtosis of a data set.</summary>
    [ExcelFunctionName("KURT")]
    ColumnValue Kurt(params ColumnValue[] values);

    /// <summary>STANDARDIZE - normalized value of a distribution.</summary>
    ColumnValue Standardize(ColumnValue x, ColumnValue mean, ColumnValue standardDev);

    /// <summary>FREQUENCY - frequency distribution of a data set.</summary>
    ColumnValue Frequency(ColumnMatrix dataArray, ColumnMatrix binsArray);

    // ---- relationships and forecasting ----

    /// <summary>CORREL - correlation coefficient of two data sets.</summary>
    [ExcelFunctionName("CORREL")]
    ColumnValue Correl(ColumnMatrix array1, ColumnMatrix array2);

    /// <summary>COVARIANCE.P - population covariance of two data sets.</summary>
    [ExcelFunctionName("COVARIANCE.P", Future = true)]
    ColumnValue CovarianceP(ColumnMatrix array1, ColumnMatrix array2);

    /// <summary>COVARIANCE.S - sample covariance of two data sets.</summary>
    [ExcelFunctionName("COVARIANCE.S", Future = true)]
    ColumnValue CovarianceS(ColumnMatrix array1, ColumnMatrix array2);

    /// <summary>SLOPE - slope of the linear regression line.</summary>
    ColumnValue Slope(ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>INTERCEPT - intercept of the linear regression line.</summary>
    ColumnValue Intercept(ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>RSQ - square of the Pearson correlation coefficient.</summary>
    [ExcelFunctionName("RSQ")]
    ColumnValue Rsq(ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>STEYX - standard error of the predicted y values.</summary>
    [ExcelFunctionName("STEYX")]
    ColumnValue SteyX(ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>FORECAST.LINEAR - predicts a value along a linear trend.</summary>
    [ExcelFunctionName("FORECAST.LINEAR", Future = true)]
    ColumnValue ForecastLinear(ColumnValue x, ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>FORECAST - predicts a value along a linear trend.</summary>
    ColumnValue Forecast(ColumnValue x, ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>TREND - values along a linear trend.</summary>
    ColumnValue Trend(ColumnMatrix knownYs, ColumnMatrix knownXs, ColumnMatrix newXs);

    /// <summary>GROWTH - values along an exponential trend.</summary>
    ColumnValue Growth(ColumnMatrix knownYs, ColumnMatrix knownXs, ColumnMatrix newXs);

    /// <summary>LINEST - parameters of a linear trend.</summary>
    [ExcelFunctionName("LINEST")]
    ColumnValue LinEst(ColumnMatrix knownYs, ColumnMatrix knownXs);

    /// <summary>LOGEST - parameters of an exponential trend.</summary>
    [ExcelFunctionName("LOGEST")]
    ColumnValue LogEst(ColumnMatrix knownYs, ColumnMatrix knownXs);

    // ---- distributions ----

    /// <summary>PERMUT - number of permutations without repetition.</summary>
    ColumnValue Permut(ColumnValue number, ColumnValue numberChosen);

    /// <summary>PERMUTATIONA - number of permutations with repetition.</summary>
    [ExcelFunctionName("PERMUTATIONA", Future = true)]
    ColumnValue PermutationA(ColumnValue number, ColumnValue numberChosen);

    /// <summary>NORM.DIST - normal cumulative distribution.</summary>
    [ExcelFunctionName("NORM.DIST", Future = true)]
    ColumnValue NormDist(ColumnValue x, ColumnValue mean, ColumnValue standardDev, ColumnValue cumulative);

    /// <summary>NORM.INV - inverse of the normal cumulative distribution.</summary>
    [ExcelFunctionName("NORM.INV", Future = true)]
    ColumnValue NormInv(ColumnValue probability, ColumnValue mean, ColumnValue standardDev);

    /// <summary>NORM.S.DIST - standard normal cumulative distribution.</summary>
    [ExcelFunctionName("NORM.S.DIST", Future = true)]
    ColumnValue NormSDist(ColumnValue z, ColumnValue cumulative);

    /// <summary>NORM.S.INV - inverse of the standard normal cumulative distribution.</summary>
    [ExcelFunctionName("NORM.S.INV", Future = true)]
    ColumnValue NormSInv(ColumnValue probability);

    /// <summary>BINOM.DIST - binomial distribution probability.</summary>
    [ExcelFunctionName("BINOM.DIST", Future = true)]
    ColumnValue BinomDist(ColumnValue numberS, ColumnValue trials, ColumnValue probabilityS, ColumnValue cumulative);

    /// <summary>POISSON.DIST - Poisson distribution probability.</summary>
    [ExcelFunctionName("POISSON.DIST", Future = true)]
    ColumnValue PoissonDist(ColumnValue x, ColumnValue mean, ColumnValue cumulative);

    /// <summary>EXPON.DIST - exponential distribution probability.</summary>
    [ExcelFunctionName("EXPON.DIST", Future = true)]
    ColumnValue ExponDist(ColumnValue x, ColumnValue lambda, ColumnValue cumulative);

    /// <summary>T.DIST - Student's t-distribution.</summary>
    [ExcelFunctionName("T.DIST", Future = true)]
    ColumnValue TDist(ColumnValue x, ColumnValue degFreedom, ColumnValue cumulative);

    /// <summary>T.INV - inverse of Student's t-distribution.</summary>
    [ExcelFunctionName("T.INV", Future = true)]
    ColumnValue TInv(ColumnValue probability, ColumnValue degFreedom);

    /// <summary>T.TEST - probability associated with a Student's t-test.</summary>
    [ExcelFunctionName("T.TEST", Future = true)]
    ColumnValue TTest(ColumnMatrix array1, ColumnMatrix array2, ColumnValue tails, ColumnValue type);

    /// <summary>CHISQ.TEST - test for independence.</summary>
    [ExcelFunctionName("CHISQ.TEST", Future = true)]
    ColumnValue ChiSqTest(ColumnMatrix actualRange, ColumnMatrix expectedRange);

    /// <summary>CONFIDENCE.NORM - confidence interval for a population mean.</summary>
    [ExcelFunctionName("CONFIDENCE.NORM", Future = true)]
    ColumnValue ConfidenceNorm(ColumnValue alpha, ColumnValue standardDev, ColumnValue size);

    /// <summary>CONFIDENCE.T - confidence interval using a Student's t-distribution.</summary>
    [ExcelFunctionName("CONFIDENCE.T", Future = true)]
    ColumnValue ConfidenceT(ColumnValue alpha, ColumnValue standardDev, ColumnValue size);

    /// <summary>PROB - probability that values are between two limits.</summary>
    ColumnValue Prob(ColumnMatrix xRange, ColumnMatrix probRange, ColumnValue lowerLimit, ColumnValue upperLimit);
}
