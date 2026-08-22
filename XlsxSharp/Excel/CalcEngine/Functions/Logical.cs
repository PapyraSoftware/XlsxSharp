using System;
using static XlsxSharp.Excel.CalcEngine.Functions.SignatureAdapter;

namespace XlsxSharp.Excel.CalcEngine.Functions;

internal static class Logical
{
    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("AND", 1, 255, And, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("FALSE", 0, 0, Adapt(False), FunctionFlags.Scalar);
        ce.RegisterFunction(
            "IF",
            2,
            3,
            AdaptLastOptional(If, false),
            FunctionFlags.Range,
            AllowRange.Only,
            1,
            2
        );
        ce.RegisterFunction("IFERROR", 2, 2, Adapt(IfError), FunctionFlags.Scalar);
        ce.RegisterFunction("NOT", 1, 1, AdaptCoerced(Not), FunctionFlags.Scalar);
        ce.RegisterFunction("OR", 1, 255, Or, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("TRUE", 0, 0, Adapt(True), FunctionFlags.Scalar);
    }

    private static AnyValue And(CalcContext ctx, Span<AnyValue> args)
    {
        OneOf<bool, XLError> aggResult = args.Aggregate(
            ctx,
            true,
            XLError.IncompatibleValue,
            static (acc, val) => acc && val,
            static (v, _) =>
            {
                // Skip values that can't be converted, but aren't errors, like "text"
                if (v.IsError)
                {
                    return v.GetError();
                }

                if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out bool logical, out XLError _))
                {
                    return true;
                }

                return logical;
            },
            static v => v.IsLogical || v.IsNumber
        ); // No text conversion for element of collection, blanks are ignored in references

        if (!aggResult.TryPickT0(out bool value, out XLError error))
        {
            return error;
        }

        return value;
    }

    private static ScalarValue False() => false;

    private static AnyValue If(ScalarValue condition, AnyValue valueIfTrue, AnyValue valueIfFalse)
    {
        if (!condition.TryCoerceLogicalOrBlankOrNumberOrText(out bool value, out XLError error))
        {
            return error;
        }

        return value ? valueIfTrue : valueIfFalse;
    }

    private static AnyValue IfError(ScalarValue potentialError, ScalarValue alternative)
    {
        if (!potentialError.IsError)
        {
            return potentialError.ToAnyValue();
        }

        return alternative.ToAnyValue();
    }

    private static AnyValue Not(bool value) => !value;

    private static AnyValue Or(CalcContext ctx, Span<AnyValue> args)
    {
        OneOf<bool, XLError> aggResult = args.Aggregate(
            ctx,
            false,
            XLError.IncompatibleValue,
            static (acc, val) => acc || val,
            static (v, _) =>
            {
                // Skip values that can't be converted, but aren't errors, like "text"
                if (v.IsError)
                {
                    return v.GetError();
                }

                if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out bool logical, out XLError _))
                {
                    return false;
                }

                return logical;
            },
            static v => v.IsLogical || v.IsNumber
        ); // No text conversion for element of collection, blanks are ignored in references

        if (!aggResult.TryPickT0(out bool value, out XLError error))
        {
            return error;
        }

        return value;
    }

    private static ScalarValue True() => true;
}
