using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.Cells;

public class XlCellValueTests
{
    [Test]
    public void CreationBlank()
    {
        XLCellValue blank = Blank.Value;
        ClassicAssert.AreEqual(XLDataType.Blank, blank.Type);
        ClassicAssert.True(blank.IsBlank);
    }

    [Test]
    public void CreationBoolean()
    {
        XLCellValue logical = true;
        ClassicAssert.AreEqual(XLDataType.Boolean, logical.Type);
        ClassicAssert.True(logical.GetBoolean());
        ClassicAssert.True(logical.IsBoolean);
    }

    [Test]
    public void CreationNumber()
    {
        XLCellValue number = 14.0;
        ClassicAssert.AreEqual(XLDataType.Number, number.Type);
        ClassicAssert.True(number.IsNumber);
        ClassicAssert.AreEqual(14.0, number.GetNumber());
    }

    [Test]
    [Arguments(double.NaN)]
    [Arguments(double.PositiveInfinity)]
    [Arguments(double.NegativeInfinity)]
    public void CreationNumberCantBeNonNumber(double nonNumber) =>
        ClassicAssert.Throws<ArgumentException>(() => _ = (XLCellValue)nonNumber);

    // Decimal is not allowed as a member of an attribute, so TestCase can't be used.
    private static readonly object[] DecimalTestCases =
    [
        new object[] { 5.875m, 5.875d },
        new object[] { decimal.MaxValue, 7.922816251426434E+28 },
        new object[] { 1.0E-28m, 1.0000000000000001E-28d },
    ];

    [Test]
    [MethodDataSource(nameof(DecimalTestCases))]
    public void CreationDecimal(decimal decimalNumber, double expectedNumber)
    {
        XLCellValue cellValue = decimalNumber;
        ClassicAssert.True(cellValue.IsNumber);
        ClassicAssert.AreEqual(expectedNumber, cellValue.GetNumber());
    }

    [Test]
    public void CreationText()
    {
        XLCellValue text = "Hello World";
        ClassicAssert.AreEqual(XLDataType.Text, text.Type);
        ClassicAssert.AreEqual("Hello World", text.GetText());
    }

    [Test]
    public void NullStringIsConvertedToBlank()
    {
        XLCellValue value = (string)null;
        ClassicAssert.IsTrue(value.IsBlank);
        ClassicAssert.IsFalse(value.IsText);
    }

    [Test]
    public void CreationTextHasLimitedLength()
    {
        string longText = new('A', 32768);
        ClassicAssert.Throws<ArgumentOutOfRangeException>(() => _ = (XLCellValue)longText);
    }

    [Test]
    public void CreationError()
    {
        XLCellValue error = XLError.NumberInvalid;
        ClassicAssert.AreEqual(XLDataType.Error, error.Type);
        ClassicAssert.True(error.IsError);
        ClassicAssert.AreEqual(XLError.NumberInvalid, error.GetError());
    }

    [Test]
    public void CreationDateTime()
    {
        XLCellValue dateTime = new DateTime(2021, 1, 1);
        ClassicAssert.AreEqual(XLDataType.DateTime, dateTime.Type);
        ClassicAssert.True(dateTime.IsDateTime);
        ClassicAssert.AreEqual(new DateTime(2021, 1, 1), dateTime.GetDateTime());
    }

    [Test]
    public void CreationTimeSpan()
    {
        XLCellValue dateTime = new TimeSpan(10, 1, 2, 3, 456);
        ClassicAssert.AreEqual(XLDataType.TimeSpan, dateTime.Type);
        ClassicAssert.True(dateTime.IsTimeSpan);
        ClassicAssert.AreEqual(new TimeSpan(10, 1, 2, 3, 456), dateTime.GetTimeSpan());
    }

    [Test]
    public void CreationFromObject()
    {
        ClassicAssert.AreEqual(XLDataType.Blank, XLCellValue.FromObject(null).Type);
        ClassicAssert.AreEqual(XLDataType.Blank, XLCellValue.FromObject(Blank.Value).Type);
        ClassicAssert.AreEqual(XLDataType.Boolean, XLCellValue.FromObject(true).Type);
        ClassicAssert.AreEqual(XLDataType.Text, XLCellValue.FromObject("Hello World").Type);
        ClassicAssert.AreEqual(
            XLDataType.Error,
            XLCellValue.FromObject(XLError.NumberInvalid).Type
        );
        ClassicAssert.AreEqual(
            XLDataType.DateTime,
            XLCellValue.FromObject(new DateTime(2021, 1, 1)).Type
        );
        ClassicAssert.AreEqual(
            XLDataType.TimeSpan,
            XLCellValue.FromObject(new TimeSpan(10, 1, 2, 3, 456)).Type
        );
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((sbyte)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((byte)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((short)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((ushort)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((int)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((uint)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((long)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((ulong)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((float)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((double)42).Type);
        ClassicAssert.AreEqual(XLDataType.Number, XLCellValue.FromObject((decimal)42).Type);
        ClassicAssert.AreEqual(XLDataType.Text, XLCellValue.FromObject(DayOfWeek.Sunday).Type);
    }

    [Test]
    public void NumberTypesHaveUnambiguousConversion()
    {
        {
            sbyte sbyteNumber = 5;
            XLCellValue sbyteCellValue = sbyteNumber;
            ClassicAssert.IsTrue(sbyteCellValue.IsNumber);
            ClassicAssert.AreEqual(5d, sbyteCellValue.GetNumber());
        }
        {
            byte byteNumber = 6;
            XLCellValue byteCellValue = byteNumber;
            ClassicAssert.IsTrue(byteCellValue.IsNumber);
            ClassicAssert.AreEqual(6d, byteCellValue.GetNumber());
        }
        {
            short shortNumber = 7;
            XLCellValue shortCellValue = shortNumber;
            ClassicAssert.IsTrue(shortCellValue.IsNumber);
            ClassicAssert.AreEqual(7d, shortCellValue.GetNumber());
        }
        {
            ushort ushortNumber = 8;
            XLCellValue ushortCellValue = ushortNumber;
            ClassicAssert.IsTrue(ushortCellValue.IsNumber);
            ClassicAssert.AreEqual(8d, ushortCellValue.GetNumber());
        }
        {
            int intNumber = 9;
            XLCellValue intCellValue = intNumber;
            ClassicAssert.IsTrue(intCellValue.IsNumber);
            ClassicAssert.AreEqual(9d, intCellValue.GetNumber());
        }
        {
            uint uintNumber = 10;
            XLCellValue uintCellValue = uintNumber;
            ClassicAssert.IsTrue(uintCellValue.IsNumber);
            ClassicAssert.AreEqual(10d, uintCellValue.GetNumber());
        }
        {
            long longNumber = 11;
            XLCellValue longCellValue = longNumber;
            ClassicAssert.IsTrue(longCellValue.IsNumber);
            ClassicAssert.AreEqual(11d, longCellValue.GetNumber());
        }
        {
            ulong ulongNumber = 12;
            XLCellValue ulongCellValue = ulongNumber;
            ClassicAssert.IsTrue(ulongCellValue.IsNumber);
            ClassicAssert.AreEqual(12d, ulongCellValue.GetNumber());
        }
        {
            float floatNumber = 13.5f;
            XLCellValue floatCellValue = floatNumber;
            ClassicAssert.IsTrue(floatCellValue.IsNumber);
            ClassicAssert.AreEqual(13.5d, floatCellValue.GetNumber());
        }
        {
            double doubleNumber = 14.5;
            XLCellValue doubleCellValue = doubleNumber;
            ClassicAssert.IsTrue(doubleCellValue.IsNumber);
            ClassicAssert.AreEqual(14.5d, doubleCellValue.GetNumber());
        }
        {
            decimal decimalNumber = 15.75m;
            XLCellValue decimalCellValue = decimalNumber;
            ClassicAssert.IsTrue(decimalCellValue.IsNumber);
            ClassicAssert.AreEqual(15.75d, decimalCellValue.GetNumber());
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableNumberWithNullValueAreConvertedToBlank()
    {
        {
            sbyte? sbyteNull = null;
            XLCellValue sbyteCellValue = sbyteNull;
            ClassicAssert.IsFalse(sbyteCellValue.IsNumber);
            ClassicAssert.IsTrue(sbyteCellValue.IsBlank);
        }
        {
            byte? byteNull = null;
            XLCellValue byteCellValue = byteNull;
            ClassicAssert.IsFalse(byteCellValue.IsNumber);
            ClassicAssert.IsTrue(byteCellValue.IsBlank);
        }
        {
            short? shortNull = null;
            XLCellValue shortCellValue = shortNull;
            ClassicAssert.IsFalse(shortCellValue.IsNumber);
            ClassicAssert.IsTrue(shortCellValue.IsBlank);
        }
        {
            ushort? ushortNull = null;
            XLCellValue ushortCellValue = ushortNull;
            ClassicAssert.IsFalse(ushortCellValue.IsNumber);
            ClassicAssert.IsTrue(ushortCellValue.IsBlank);
        }
        {
            int? intNull = null;
            XLCellValue intCellValue = intNull;
            ClassicAssert.IsFalse(intCellValue.IsNumber);
            ClassicAssert.IsTrue(intCellValue.IsBlank);
        }
        {
            uint? uintNull = null;
            XLCellValue uintCellValue = uintNull;
            ClassicAssert.IsFalse(uintCellValue.IsNumber);
            ClassicAssert.IsTrue(uintCellValue.IsBlank);
        }
        {
            long? longNull = null;
            XLCellValue longCellValue = longNull;
            ClassicAssert.IsFalse(longCellValue.IsNumber);
            ClassicAssert.IsTrue(longCellValue.IsBlank);
        }
        {
            ulong? ulongNull = null;
            XLCellValue ulongCellValue = ulongNull;
            ClassicAssert.IsFalse(ulongCellValue.IsNumber);
            ClassicAssert.IsTrue(ulongCellValue.IsBlank);
        }
        {
            float? floatValue = null;
            XLCellValue floatCellValue = floatValue;
            ClassicAssert.IsFalse(floatCellValue.IsNumber);
            ClassicAssert.IsTrue(floatCellValue.IsBlank);
        }
        {
            double? doubleValue = null;
            XLCellValue doubleCellValue = doubleValue;
            ClassicAssert.IsFalse(doubleCellValue.IsNumber);
            ClassicAssert.IsTrue(doubleCellValue.IsBlank);
        }
        {
            decimal? decimalValue = null;
            XLCellValue decimalCellValue = decimalValue;
            ClassicAssert.IsFalse(decimalCellValue.IsNumber);
            ClassicAssert.IsTrue(decimalCellValue.IsBlank);
        }
    }

    [Test]
    public void NullableNumberWithNumberValueAreConvertedToNumber()
    {
        {
            sbyte? sbyteNumber = 5;
            XLCellValue sbyteCellValue = sbyteNumber;
            ClassicAssert.IsTrue(sbyteCellValue.IsNumber);
            ClassicAssert.AreEqual(5d, sbyteCellValue.GetNumber());
        }
        {
            byte? byteNumber = 6;
            XLCellValue byteCellValue = byteNumber;
            ClassicAssert.IsTrue(byteCellValue.IsNumber);
            ClassicAssert.AreEqual(6d, byteCellValue.GetNumber());
        }
        {
            short? shortNumber = 7;
            XLCellValue shortCellValue = shortNumber;
            ClassicAssert.IsTrue(shortCellValue.IsNumber);
            ClassicAssert.AreEqual(7d, shortCellValue.GetNumber());
        }
        {
            ushort? ushortNumber = 8;
            XLCellValue ushortCellValue = ushortNumber;
            ClassicAssert.IsTrue(ushortCellValue.IsNumber);
            ClassicAssert.AreEqual(8d, ushortCellValue.GetNumber());
        }
        {
            int? intNumber = 9;
            XLCellValue intCellValue = intNumber;
            ClassicAssert.IsTrue(intCellValue.IsNumber);
            ClassicAssert.AreEqual(9d, intCellValue.GetNumber());
        }
        {
            uint? uintNumber = 9;
            XLCellValue uintCellValue = uintNumber;
            ClassicAssert.IsTrue(uintCellValue.IsNumber);
            ClassicAssert.AreEqual(9d, uintCellValue.GetNumber());
        }
        {
            long? longNumber = 10;
            XLCellValue longCellValue = longNumber;
            ClassicAssert.IsTrue(longCellValue.IsNumber);
            ClassicAssert.AreEqual(10d, longCellValue.GetNumber());
        }
        {
            ulong? ulongNumber = 11;
            XLCellValue ulongCellValue = ulongNumber;
            ClassicAssert.IsTrue(ulongCellValue.IsNumber);
            ClassicAssert.AreEqual(11d, ulongCellValue.GetNumber());
        }
        {
            float? floatNumber = 12.875f;
            XLCellValue floatCellValue = floatNumber;
            ClassicAssert.IsTrue(floatCellValue.IsNumber);
            ClassicAssert.AreEqual(12.875d, floatCellValue.GetNumber());
        }
        {
            double? doubleNumber = 13.875d;
            XLCellValue doubleCellValue = doubleNumber;
            ClassicAssert.IsTrue(doubleCellValue.IsNumber);
            ClassicAssert.AreEqual(13.875d, doubleCellValue.GetNumber());
        }
        {
            decimal? decimalNumber = 14.875m;
            XLCellValue decimalCellValue = decimalNumber;
            ClassicAssert.IsTrue(decimalCellValue.IsNumber);
            ClassicAssert.AreEqual(14.875d, decimalCellValue.GetNumber());
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableDateTimeWithNullValueIsConvertedToBlank()
    {
        DateTime? dateTimeNull = null;
        XLCellValue dateTimeCellValue = dateTimeNull;
        ClassicAssert.IsFalse(dateTimeCellValue.IsDateTime);
        ClassicAssert.IsTrue(dateTimeCellValue.IsBlank);
    }

    [Test]
    public void NullableDateTimeWithDateValueIsConvertedToDateTime()
    {
        DateTime? dateTime = new DateTime(2020, 5, 14, 8, 14, 30);
        XLCellValue dateTimeCellValue = dateTime;
        ClassicAssert.IsTrue(dateTimeCellValue.IsDateTime);
        ClassicAssert.AreEqual(dateTime.Value, dateTimeCellValue.GetDateTime());
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableTimeSpanWithNullValueIsConvertedToBlank()
    {
        TimeSpan? timeSpanNull = null;
        XLCellValue timeSpanCellValue = timeSpanNull;
        ClassicAssert.IsFalse(timeSpanCellValue.IsTimeSpan);
        ClassicAssert.IsTrue(timeSpanCellValue.IsBlank);
    }

    [Test]
    public void NullableTimeSpanWithTimeSpanValueIsConvertedToTimeSpan()
    {
        TimeSpan? timeSpan = new TimeSpan(48, 12, 45, 30);
        XLCellValue timeSpanCellValue = timeSpan;
        ClassicAssert.IsTrue(timeSpanCellValue.IsTimeSpan);
        ClassicAssert.AreEqual(timeSpan.Value, timeSpanCellValue.GetTimeSpan());
    }

    [Test]
    public void UnifiedNumberIsFormOfNumberDateTimeAndTimeSpan()
    {
        XLCellValue value = Blank.Value;
        ClassicAssert.False(value.IsUnifiedNumber);

        value = true;
        ClassicAssert.False(value.IsUnifiedNumber);

        value = 14;
        ClassicAssert.True(value.IsUnifiedNumber);
        ClassicAssert.AreEqual(14.0, value.GetUnifiedNumber());

        value = new DateTime(1900, 1, 1);
        ClassicAssert.True(value.IsUnifiedNumber);
        ClassicAssert.AreEqual(1.0, value.GetUnifiedNumber());

        value = new TimeSpan(2, 12, 0, 0);
        ClassicAssert.True(value.IsUnifiedNumber);
        ClassicAssert.AreEqual(2.5, value.GetUnifiedNumber());

        value = "Text";
        ClassicAssert.False(value.IsUnifiedNumber);

        value = XLError.CellReference;
        ClassicAssert.False(value.IsUnifiedNumber);
    }

    [Test]
    [Arguments("1900-01-01", 1)]
    [Arguments("1900-01-02", 2)]
    [Arguments("1900-02-01", 32)]
    [Arguments("1900-02-28", 59)] // Excel assumes 1900 was a leap year and 29.1.1900 existed
    [Arguments("1900-03-01", 61)]
    [Arguments("2017-01-01", 42736)]
    public void SerialDateTime(string dateString, double expectedSerial)
    {
        XLCellValue date = DateTime.Parse(dateString);
        ClassicAssert.AreEqual(expectedSerial, date.GetUnifiedNumber());
    }

    [Test]
    [Culture("cs-CZ")]
    public void ToStringRespectsCulture()
    {
        XLCellValue v = Blank.Value;
        ClassicAssert.AreEqual(string.Empty, v.ToString());

        v = true;
        ClassicAssert.AreEqual("TRUE", v.ToString());

        v = 25.4;
        ClassicAssert.AreEqual("25,4", v.ToString());

        v = "Hello";
        ClassicAssert.AreEqual("Hello", v.ToString());

        v = XLError.IncompatibleValue;
        ClassicAssert.AreEqual("#VALUE!", v.ToString());

        v = new DateTime(1900, 1, 2);
        ClassicAssert.AreEqual("02.01.1900 0:00:00", v.ToString());

        v = new DateTime(1900, 3, 1, 4, 10, 5);
        ClassicAssert.AreEqual("01.03.1900 4:10:05", v.ToString());

        v = new TimeSpan(4, 5, 6, 7, 82);
        ClassicAssert.AreEqual("101:06:07,082", v.ToString());
    }

    [Test]
    public void TryConvertBlank()
    {
        XLCellValue value = Blank.Value;
        ClassicAssert.True(value.TryConvert(out Blank blank));
        ClassicAssert.AreEqual(Blank.Value, blank);

        value = string.Empty;
        ClassicAssert.True(value.TryConvert(out blank));
        ClassicAssert.AreEqual(Blank.Value, blank);
    }

    [Test]
    public void TryConvertBoolean()
    {
        XLCellValue value = true;
        ClassicAssert.True(value.TryConvert(out bool boolean));
        ClassicAssert.True(boolean);

        value = "True";
        ClassicAssert.True(value.TryConvert(out boolean));
        ClassicAssert.True(boolean);

        value = "False";
        ClassicAssert.True(value.TryConvert(out boolean));
        ClassicAssert.False(boolean);

        value = 0;
        ClassicAssert.True(value.TryConvert(out boolean));
        ClassicAssert.False(boolean);

        value = 0.001;
        ClassicAssert.True(value.TryConvert(out boolean));
        ClassicAssert.True(boolean);
    }

    [Test]
    public void TryConvertNumber()
    {
        CultureInfo c = CultureInfo.GetCultureInfo("cs-CZ");
        XLCellValue value = 5;
        ClassicAssert.True(value.TryConvert(out double number, c));
        ClassicAssert.AreEqual(5.0, number);

        value = "1,5";
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(1.5, number);

        value = "1 1/4";
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(1.25, number);

        value = "3.1.1900";
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(3, number);

        value = true;
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(1.0, number);

        value = false;
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(0.0, number);

        value = new DateTime(2020, 4, 5, 10, 14, 5);
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(43926.42644675926, number);

        value = new TimeSpan(18, 0, 0);
        ClassicAssert.True(value.TryConvert(out number, c));
        ClassicAssert.AreEqual(0.75, number);
    }

    [Test]
    public void TryConvertDateTime()
    {
        XLCellValue v = new DateTime(2020, 1, 1);
        ClassicAssert.True(v.TryConvert(out DateTime dt));
        ClassicAssert.AreEqual(new DateTime(2020, 1, 1), dt);

        int lastSerialDate = 2958465;
        v = lastSerialDate;
        ClassicAssert.True(v.TryConvert(out dt));
        ClassicAssert.AreEqual(new DateTime(9999, 12, 31), dt);

        v = lastSerialDate + 1;
        ClassicAssert.False(v.TryConvert(out dt));

        v = new TimeSpan(14, 0, 0, 0);
        ClassicAssert.True(v.TryConvert(out dt));
        ClassicAssert.AreEqual(new DateTime(1900, 1, 14), dt);
    }

    [Test]
    public void TryConvertTimeSpan()
    {
        CultureInfo c = CultureInfo.GetCultureInfo("cs-CZ");
        XLCellValue v = new TimeSpan(10, 15, 30);
        ClassicAssert.True(v.TryConvert(out TimeSpan ts, c));
        ClassicAssert.AreEqual(new TimeSpan(10, 15, 30), ts);

        v = "26:15:30,5";
        ClassicAssert.True(v.TryConvert(out ts, c));
        ClassicAssert.AreEqual(new TimeSpan(1, 2, 15, 30, 500), ts);

        v = 0.75;
        ClassicAssert.True(v.TryConvert(out ts, c));
        ClassicAssert.AreEqual(new TimeSpan(18, 0, 0), ts);
    }

    [Test]
    [Arguments(1)]
    [Arguments(10)] // microsecond
    [Arguments(3000000001)] // 5 min 1 tick
    public void TimeSpanCanHaveSubMillisecondPrecision(long ticks)
    {
        TimeSpan subMsTimeSpan = TimeSpan.FromTicks(ticks);
        XLCellValue value = subMsTimeSpan;
        ClassicAssert.AreEqual(subMsTimeSpan, value.GetTimeSpan());
    }

    [Test]
    [Arguments(1)]
    [Arguments(10)] // microsecond
    [Arguments(3000000001)] // 5 min 1 tick
    public void TimeSpanWithSubMillisecondPrecisionIsWrittenAndLoadedCorrectly(long ticks)
    {
        // NetFx converts double to string using G15. Core changed it to G17, but XlsxSharp still use G15.
        TimeSpan subMsTimeSpan = TimeSpan.FromTicks(ticks);
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = subMsTimeSpan;
            },
            (_, ws) =>
            {
                XLCellValue cellValue = ws.Cell("A1").CachedValue;
                ClassicAssert.AreEqual(subMsTimeSpan, cellValue.GetTimeSpan());
            }
        );
    }

    [Test]
    [Arguments(long.MaxValue / (double)TimeSpan.TicksPerDay + 0.01)]
    [Arguments(long.MinValue / (double)TimeSpan.TicksPerDay - 0.01)]
    public void TimeSpanThrowsWhenNotRepresentable(double serialDateTime)
    {
        XLCellValue value = XLCellValue.FromSerialTimeSpan(serialDateTime);
        OverflowException ex = ClassicAssert.Throws<OverflowException>(() => value.GetTimeSpan())!;
        ClassicAssert.AreEqual(
            "The serial date time value is too large to be represented in a TimeSpan.",
            ex.Message
        );
    }
}
