namespace Chsword.Excel2Object.Functions;

/// <summary>
///     Excel financial functions. See <see cref="IExcelFunction" /> for how these are translated.
/// </summary>
public interface IFinancialFunction : IExcelFunction
{
    // ---- annuities ----

    /// <summary>PMT - the payment of a loan with constant payments and a constant rate.</summary>
    [ExcelFunctionName("PMT")]
    ColumnValue Pmt(ColumnValue rate, ColumnValue nper, ColumnValue pv);

    /// <summary>PMT - as above, with a future value and a payment type.</summary>
    [ExcelFunctionName("PMT")]
    ColumnValue Pmt(ColumnValue rate, ColumnValue nper, ColumnValue pv, ColumnValue fv, ColumnValue type);

    /// <summary>IPMT - the interest part of a payment.</summary>
    [ExcelFunctionName("IPMT")]
    ColumnValue IPmt(ColumnValue rate, ColumnValue per, ColumnValue nper, ColumnValue pv);

    /// <summary>PPMT - the principal part of a payment.</summary>
    [ExcelFunctionName("PPMT")]
    ColumnValue PPmt(ColumnValue rate, ColumnValue per, ColumnValue nper, ColumnValue pv);

    /// <summary>PV - the present value of an investment.</summary>
    [ExcelFunctionName("PV")]
    ColumnValue Pv(ColumnValue rate, ColumnValue nper, ColumnValue pmt);

    /// <summary>PV - as above, with a future value and a payment type.</summary>
    [ExcelFunctionName("PV")]
    ColumnValue Pv(ColumnValue rate, ColumnValue nper, ColumnValue pmt, ColumnValue fv, ColumnValue type);

    /// <summary>FV - the future value of an investment.</summary>
    [ExcelFunctionName("FV")]
    ColumnValue Fv(ColumnValue rate, ColumnValue nper, ColumnValue pmt);

    /// <summary>FV - as above, with a present value and a payment type.</summary>
    [ExcelFunctionName("FV")]
    ColumnValue Fv(ColumnValue rate, ColumnValue nper, ColumnValue pmt, ColumnValue pv, ColumnValue type);

    /// <summary>NPER - the number of periods of an investment.</summary>
    [ExcelFunctionName("NPER")]
    ColumnValue NPer(ColumnValue rate, ColumnValue pmt, ColumnValue pv);

    /// <summary>RATE - the interest rate per period of an annuity.</summary>
    [ExcelFunctionName("RATE")]
    ColumnValue Rate(ColumnValue nper, ColumnValue pmt, ColumnValue pv);

    /// <summary>CUMIPMT - the cumulative interest paid between two periods.</summary>
    [ExcelFunctionName("CUMIPMT")]
    ColumnValue CumIPmt(ColumnValue rate, ColumnValue nper, ColumnValue pv, ColumnValue startPeriod,
        ColumnValue endPeriod, ColumnValue type);

    /// <summary>CUMPRINC - the cumulative principal paid between two periods.</summary>
    [ExcelFunctionName("CUMPRINC")]
    ColumnValue CumPrinc(ColumnValue rate, ColumnValue nper, ColumnValue pv, ColumnValue startPeriod,
        ColumnValue endPeriod, ColumnValue type);

    // ---- cash flows ----

    /// <summary>NPV - the net present value of a series of cash flows.</summary>
    [ExcelFunctionName("NPV")]
    ColumnValue Npv(ColumnValue rate, params ColumnValue[] values);

    /// <summary>XNPV - the net present value of cash flows on given dates.</summary>
    [ExcelFunctionName("XNPV")]
    ColumnValue XNpv(ColumnValue rate, ColumnMatrix values, ColumnMatrix dates);

    /// <summary>IRR - the internal rate of return of a series of cash flows.</summary>
    [ExcelFunctionName("IRR")]
    ColumnValue Irr(ColumnMatrix values);

    /// <summary>IRR - the internal rate of return, starting from a guess.</summary>
    [ExcelFunctionName("IRR")]
    ColumnValue Irr(ColumnMatrix values, ColumnValue guess);

    /// <summary>XIRR - the internal rate of return of cash flows on given dates.</summary>
    [ExcelFunctionName("XIRR")]
    ColumnValue XIrr(ColumnMatrix values, ColumnMatrix dates);

    /// <summary>MIRR - the internal rate of return with different borrowing and reinvestment rates.</summary>
    [ExcelFunctionName("MIRR")]
    ColumnValue MIrr(ColumnMatrix values, ColumnValue financeRate, ColumnValue reinvestRate);

    // ---- depreciation and rates ----

    /// <summary>SLN - the straight line depreciation of an asset for one period.</summary>
    [ExcelFunctionName("SLN")]
    ColumnValue Sln(ColumnValue cost, ColumnValue salvage, ColumnValue life);

    /// <summary>SYD - the sum of years digits depreciation for a period.</summary>
    [ExcelFunctionName("SYD")]
    ColumnValue Syd(ColumnValue cost, ColumnValue salvage, ColumnValue life, ColumnValue per);

    /// <summary>DB - the fixed declining balance depreciation for a period.</summary>
    [ExcelFunctionName("DB")]
    ColumnValue Db(ColumnValue cost, ColumnValue salvage, ColumnValue life, ColumnValue period);

    /// <summary>DDB - the double declining balance depreciation for a period.</summary>
    [ExcelFunctionName("DDB")]
    ColumnValue Ddb(ColumnValue cost, ColumnValue salvage, ColumnValue life, ColumnValue period);

    /// <summary>VDB - the declining balance depreciation between two periods.</summary>
    [ExcelFunctionName("VDB")]
    ColumnValue Vdb(ColumnValue cost, ColumnValue salvage, ColumnValue life, ColumnValue startPeriod,
        ColumnValue endPeriod);

    /// <summary>EFFECT - the effective annual interest rate.</summary>
    [ExcelFunctionName("EFFECT")]
    ColumnValue Effect(ColumnValue nominalRate, ColumnValue nPeryear);

    /// <summary>NOMINAL - the nominal annual interest rate.</summary>
    [ExcelFunctionName("NOMINAL")]
    ColumnValue Nominal(ColumnValue effectRate, ColumnValue nPeryear);
}
