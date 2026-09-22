using System;
using System.Globalization;
using ExcelDna.Integration;
using NepDate;

namespace ExNepDate;

/// <summary>
/// Helper utilities for normalizing Excel-DNA inputs, cell references, and formatting outputs for NepDate.
/// </summary>
public static class NepDateHelper
{
    /// <summary>
    /// Resolves an Excel-DNA input object, extracting the raw value if it is an ExcelReference,
    /// 2D array, or handling missing/empty cells.
    /// </summary>
    public static object? ResolveValue(object? input)
    {
        if (input == null || input is ExcelMissing || input is ExcelEmpty)
        {
            return null;
        }

        if (input is ExcelReference reference)
        {
            var val = reference.GetValue();
            if (val is object[,] arr && arr.GetLength(0) > 0 && arr.GetLength(1) > 0)
            {
                return ResolveValue(arr[0, 0]);
            }
            return ResolveValue(val);
        }

        if (input is object[,] array && array.GetLength(0) > 0 && array.GetLength(1) > 0)
        {
            return ResolveValue(array[0, 0]);
        }

        return input;
    }

    /// <summary>
    /// Resolves an input into a valid NepaliDate instance.
    /// Supports Nepali date strings, Unicode Devanagari numerals, DateTime, and Excel date numbers.
    /// </summary>
    public static NepaliDate ResolveNepaliDate(object? input)
    {
        var val = ResolveValue(input);
        if (val == null)
        {
            throw new ArgumentException("Nepali date input cannot be empty.");
        }

        if (val is NepaliDate nd)
        {
            return nd;
        }

        if (val is DateTime dt)
        {
            return dt.ToNepaliDate();
        }

        if (val is string str)
        {
            str = str.Trim();
            if (string.IsNullOrEmpty(str))
            {
                throw new ArgumentException("Date string cannot be empty.");
            }

            if (SmartDateParser.TryParse(str, out var parsed))
            {
                return parsed;
            }

            if (NepaliDate.TryParse(str, out var parsed2))
            {
                return parsed2;
            }

            if (DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.None, out var engDate) ||
                DateTime.TryParse(str, out engDate))
            {
                return engDate.ToNepaliDate();
            }

            throw new ArgumentException($"Unable to parse '{str}' as a valid Nepali date.");
        }

        if (val is double d)
        {
            if (d > 10000 && d < 150000)
            {
                return DateTime.FromOADate(d).ToNepaliDate();
            }

            if (d >= 1900 && d <= 2200)
            {
                return new NepaliDate((int)d, 1, 1);
            }
        }

        if (val is int year && year >= 1900 && year <= 2200)
        {
            return new NepaliDate(year, 1, 1);
        }

        throw new ArgumentException($"Invalid Nepali date value: '{val}'.");
    }

    /// <summary>
    /// Resolves an input into a Gregorian DateTime.
    /// Supports DateTime, Excel serial dates (double/int), and string dates.
    /// </summary>
    public static DateTime ResolveEnglishDate(object? input)
    {
        var val = ResolveValue(input);
        if (val == null)
        {
            throw new ArgumentException("English date input cannot be empty.");
        }

        if (val is DateTime dt)
        {
            return dt.Date;
        }

        if (val is double d)
        {
            return DateTime.FromOADate(d).Date;
        }

        if (val is int i)
        {
            return DateTime.FromOADate(i).Date;
        }

        if (val is string str)
        {
            str = str.Trim();
            if (DateTime.TryParse(str, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDt) ||
                DateTime.TryParse(str, out parsedDt))
            {
                return parsedDt.Date;
            }

            if (SmartDateParser.TryParse(str, out var nepDate))
            {
                return nepDate.EnglishDate.Date;
            }

            throw new ArgumentException($"Unable to parse '{str}' as a valid date.");
        }

        throw new ArgumentException($"Invalid English date value: '{val}'.");
    }

    /// <summary>
    /// Formats a NepaliDate using the provided format string or defaults to yyyy-MM-dd.
    /// </summary>
    public static string FormatResult(NepaliDate date, object? format)
    {
        var fmtVal = ResolveValue(format);
        string fmt = fmtVal is string s && !string.IsNullOrWhiteSpace(s) ? s.Trim() : "yyyy-MM-dd";
        return date.ToString(fmt, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Resolves a boolean value from various Excel input types.
    /// </summary>
    public static bool ResolveBool(object? input, bool defaultValue = false)
    {
        var val = ResolveValue(input);
        if (val == null) return defaultValue;
        if (val is bool b) return b;
        if (val is double d) return d != 0;
        if (val is int i) return i != 0;
        if (val is string s)
        {
            if (bool.TryParse(s.Trim(), out var result)) return result;
            if (s.Trim() == "1") return true;
            if (s.Trim() == "0") return false;
        }
        return defaultValue;
    }

    /// <summary>
    /// Resolves an integer value from various Excel input types.
    /// </summary>
    public static int ResolveInt(object? input, int defaultValue = 0)
    {
        var val = ResolveValue(input);
        if (val == null) return defaultValue;
        if (val is int i) return i;
        if (val is double d) return (int)Math.Round(d);
        if (val is string s && int.TryParse(s.Trim(), out var parsed)) return parsed;
        return defaultValue;
    }

    /// <summary>
    /// Resolves a double value from various Excel input types.
    /// </summary>
    public static double ResolveDouble(object? input, double defaultValue = 0.0)
    {
        var val = ResolveValue(input);
        if (val == null) return defaultValue;
        if (val is double d) return d;
        if (val is int i) return i;
        if (val is string s && double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)) return parsed;
        return defaultValue;
    }

    /// <summary>
    /// Resolves a FiscalYearQuarters enum value.
    /// </summary>
    public static FiscalYearQuarters ResolveFiscalQuarter(object? quarter)
    {
        var val = ResolveValue(quarter);
        if (val == null) return FiscalYearQuarters.Current;

        if (val is double d) val = (int)d;
        if (val is int q)
        {
            return q switch
            {
                1 => FiscalYearQuarters.First,
                2 => FiscalYearQuarters.Second,
                3 => FiscalYearQuarters.Third,
                4 => FiscalYearQuarters.Fourth,
                _ => FiscalYearQuarters.Current
            };
        }
        if (val is string s)
        {
            s = s.Trim().ToLowerInvariant();
            return s switch
            {
                "1" or "q1" or "first" => FiscalYearQuarters.First,
                "2" or "q2" or "second" => FiscalYearQuarters.Second,
                "3" or "q3" or "third" => FiscalYearQuarters.Third,
                "4" or "q4" or "fourth" => FiscalYearQuarters.Fourth,
                _ => FiscalYearQuarters.Current
            };
        }
        return FiscalYearQuarters.Current;
    }

    /// <summary>
    /// Resolves a Separators enum value.
    /// </summary>
    public static Separators ResolveSeparator(object? separator)
    {
        var val = ResolveValue(separator);
        if (val == null) return Separators.Dash;
        if (val is double d) return (Separators)(int)d;
        if (val is int i && Enum.IsDefined(typeof(Separators), i)) return (Separators)i;
        if (val is string s)
        {
            s = s.Trim().ToLowerInvariant();
            return s switch
            {
                "/" or "forwardslash" or "slash" => Separators.ForwardSlash,
                "\\" or "backslash" => Separators.BackwardSlash,
                "." or "dot" => Separators.Dot,
                "_" or "underscore" => Separators.Underscore,
                "-" or "dash" or "hyphen" => Separators.Dash,
                " " or "space" => Separators.Space,
                _ => Separators.Dash
            };
        }
        return Separators.Dash;
    }

    /// <summary>
    /// Resolves a DateFormats enum value.
    /// </summary>
    public static DateFormats ResolveDateFormat(object? format)
    {
        var val = ResolveValue(format);
        if (val == null) return DateFormats.YearMonthDay;
        if (val is double d) return (DateFormats)(int)d;
        if (val is int i && Enum.IsDefined(typeof(DateFormats), i)) return (DateFormats)i;
        if (val is string s)
        {
            s = s.Trim().ToLowerInvariant();
            return s switch
            {
                "ymd" or "yearmonthday" => DateFormats.YearMonthDay,
                "ydm" or "yeardaymonth" => DateFormats.YearDayMonth,
                "myd" or "monthyearday" => DateFormats.MonthYearDay,
                "mdy" or "monthdayyear" => DateFormats.MonthDayYear,
                "dym" or "dayyearmonth" => DateFormats.DayYearMonth,
                "dmy" or "daymonthyear" => DateFormats.DayMonthYear,
                _ => DateFormats.YearMonthDay
            };
        }
        return DateFormats.YearMonthDay;
    }
}
