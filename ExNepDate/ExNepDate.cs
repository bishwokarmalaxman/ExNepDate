using System;
using System.Linq;
using ExcelDna.Integration;
using NepDate;

namespace ExNepDate;

/// <summary>
/// Excel Custom Functions (UDFs) for Nepali Date (Bikram Sambat) operations powered by NepDate.
/// </summary>
public class ExNepDate
{
    private const string CategoryName = "NepDate";

    #region Sample & Connectivity

    [ExcelFunction(Name = "SayHello", Description = "My first .NET function", Category = CategoryName)]
    public static string SayHello(string name)
    {
        return "Hello " + name;
    }

    #endregion

    #region Core & Conversions

    [ExcelFunction(Name = "EX_Today", Description = "Returns the current Nepali date.", Category = CategoryName)]
    public static string EX_Today(
        [ExcelArgument(Description = "Optional output format (e.g. yyyy-MM-dd, yyyy/MM/dd, dd-MM-yyyy). Default is yyyy-MM-dd.")] object? format = null)
    {
        return NepDateHelper.FormatResult(NepaliDate.Today, format);
    }

    [ExcelFunction(Name = "EX_Now", Description = "Returns the current Nepali date.", Category = CategoryName)]
    public static string EX_Now(
        [ExcelArgument(Description = "Optional output format (e.g. yyyy-MM-dd, yyyy/MM/dd). Default is yyyy-MM-dd.")] object? format = null)
    {
        return NepDateHelper.FormatResult(NepaliDate.Now, format);
    }

    [ExcelFunction(Name = "EX_NepaliDate", Description = "Converts an English (Gregorian) date to Nepali date (Bikram Sambat).", Category = CategoryName)]
    public static string EX_NepaliDate(
        [ExcelArgument(AllowReference = true, Description = "English date reference, cell, date serial number, or date string.")] object date,
        [ExcelArgument(Description = "Optional output format (e.g. yyyy-MM-dd). Default is yyyy-MM-dd.")] object? format = null)
    {
        var engDate = NepDateHelper.ResolveEnglishDate(date);
        return NepDateHelper.FormatResult(engDate.ToNepaliDate(), format);
    }

    [ExcelFunction(Name = "EX_ToNepaliDate", Description = "Converts an English (Gregorian) date to Nepali date (Bikram Sambat).", Category = CategoryName)]
    public static string EX_ToNepaliDate(
        [ExcelArgument(AllowReference = true, Description = "English date reference, cell, date serial number, or date string.")] object date,
        [ExcelArgument(Description = "Optional output format (e.g. yyyy-MM-dd). Default is yyyy-MM-dd.")] object? format = null)
    {
        return EX_NepaliDate(date, format);
    }

    [ExcelFunction(Name = "EX_ToEnglishDate", Description = "Converts a Nepali date (B.S.) to an English (Gregorian) date. Returns native Excel date.", Category = CategoryName)]
    public static object EX_ToEnglishDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string (e.g. 2080-01-15, 2080/01/15, or Unicode).")] object nepaliDate,
        [ExcelArgument(Description = "Optional string format (e.g. yyyy-MM-dd). If omitted, returns native Excel date serial.")] object? format = null)
    {
        var nepDate = NepDateHelper.ResolveNepaliDate(nepaliDate);
        var engDate = nepDate.EnglishDate;

        var fmtVal = NepDateHelper.ResolveValue(format);
        if (fmtVal is string fmt && !string.IsNullOrWhiteSpace(fmt))
        {
            return engDate.ToString(fmt.Trim());
        }

        return engDate;
    }

    [ExcelFunction(Name = "EX_EnglishDate", Description = "Converts a Nepali date (B.S.) to an English (Gregorian) date. Returns native Excel date.", Category = CategoryName)]
    public static object EX_EnglishDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string (e.g. 2080-01-15, 2080/01/15, or Unicode).")] object nepaliDate,
        [ExcelArgument(Description = "Optional string format (e.g. yyyy-MM-dd). If omitted, returns native Excel date serial.")] object? format = null)
    {
        return EX_ToEnglishDate(nepaliDate, format);
    }

    [ExcelFunction(Name = "EX_Date", Description = "Creates a Nepali date from year, month, and day integers.", Category = CategoryName)]
    public static string EX_Date(
        [ExcelArgument(Description = "Nepali year (1901–2199 B.S., e.g. 2080).")] object year,
        [ExcelArgument(Description = "Nepali month (1–12).")] object month,
        [ExcelArgument(Description = "Nepali day (1–32).")] object day,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        int y = NepDateHelper.ResolveInt(year);
        int m = NepDateHelper.ResolveInt(month);
        int d = NepDateHelper.ResolveInt(day);

        var date = new NepaliDate(y, m, d);
        return NepDateHelper.FormatResult(date, format);
    }

    [ExcelFunction(Name = "EX_Parse", Description = "Smart-parses any Nepali date text (including 100+ month spellings and Nepali Unicode digits).", Category = CategoryName)]
    public static string EX_Parse(
        [ExcelArgument(AllowReference = true, Description = "Nepali date text (e.g. '15 Baishakh 2080', '२०८०/०१/१५').")] object dateText,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var nepDate = NepDateHelper.ResolveNepaliDate(dateText);
        return NepDateHelper.FormatResult(nepDate, format);
    }

    #endregion

    #region Date Components & Getters

    [ExcelFunction(Name = "EX_Year", Description = "Returns the year component of a Nepali date.", Category = CategoryName)]
    public static int EX_Year(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).Year;
    }

    [ExcelFunction(Name = "EX_Month", Description = "Returns the month number (1–12) of a Nepali date.", Category = CategoryName)]
    public static int EX_Month(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).Month;
    }

    [ExcelFunction(Name = "EX_Day", Description = "Returns the day of the month (1–32) of a Nepali date.", Category = CategoryName)]
    public static int EX_Day(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).Day;
    }

    [ExcelFunction(Name = "EX_MonthName", Description = "Returns the English name of the Nepali month (e.g. Baishakh, Jestha).", Category = CategoryName)]
    public static string EX_MonthName(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).MonthName.ToString();
    }

    [ExcelFunction(Name = "EX_DayOfWeek", Description = "Returns the day of the week name (e.g. Sunday, Monday) for a Nepali date.", Category = CategoryName)]
    public static string EX_DayOfWeek(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).DayOfWeek.ToString();
    }

    [ExcelFunction(Name = "EX_DayOfYear", Description = "Returns the day number within the Nepali year (1–366).", Category = CategoryName)]
    public static int EX_DayOfYear(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).DayOfYear;
    }

    [ExcelFunction(Name = "EX_MonthEndDay", Description = "Returns the total number of days in the month of the specified Nepali date (e.g. 30, 31, 32).", Category = CategoryName)]
    public static int EX_MonthEndDay(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).MonthEndDay;
    }

    [ExcelFunction(Name = "EX_MonthEndDate", Description = "Returns the last date of the month for the given Nepali date.", Category = CategoryName)]
    public static string EX_MonthEndDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var nepDate = NepDateHelper.ResolveNepaliDate(nepaliDate);
        return NepDateHelper.FormatResult(nepDate.MonthEndDate(), format);
    }

    #endregion

    #region Date Verification & Status

    [ExcelFunction(Name = "EX_IsLeapYear", Description = "Determines whether the specified Nepali date or year is a leap year.", Category = CategoryName)]
    public static bool EX_IsLeapYear(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or year number (e.g. 2080).")] object nepaliDateOrYear)
    {
        var val = NepDateHelper.ResolveValue(nepaliDateOrYear);
        if (val is int y || (val is double d && d >= 1900 && d <= 2200 && Math.Abs(d - Math.Round(d)) < 0.001) || (val is string s && int.TryParse(s, out y) && y >= 1900 && y <= 2200))
        {
            int yearVal = val is int i ? i : (val is double db ? (int)db : int.Parse((string)val!));
            return new NepaliDate(yearVal, 1, 1).IsLeapYear();
        }

        return NepDateHelper.ResolveNepaliDate(nepaliDateOrYear).IsLeapYear();
    }

    [ExcelFunction(Name = "EX_IsToday", Description = "Checks if the specified Nepali date is today.", Category = CategoryName)]
    public static bool EX_IsToday(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).IsToday();
    }

    [ExcelFunction(Name = "EX_IsTomorrow", Description = "Checks if the specified Nepali date is tomorrow.", Category = CategoryName)]
    public static bool EX_IsTomorrow(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).IsTomorrow();
    }

    [ExcelFunction(Name = "EX_IsYesterday", Description = "Checks if the specified Nepali date is yesterday.", Category = CategoryName)]
    public static bool EX_IsYesterday(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).IsYesterday();
    }

    [ExcelFunction(Name = "EX_IsValidDate", Description = "Checks whether year, month, and day form a valid Nepali date (accounting for the exact calendar day limits).", Category = CategoryName)]
    public static bool EX_IsValidDate(
        [ExcelArgument(Description = "Nepali year (1901–2199 B.S.).")] object year,
        [ExcelArgument(Description = "Nepali month (1–12).")] object month,
        [ExcelArgument(Description = "Nepali day (1–32).")] object day)
    {
        try
        {
            int y = NepDateHelper.ResolveInt(year);
            int m = NepDateHelper.ResolveInt(month);
            int d = NepDateHelper.ResolveInt(day);
            var _ = new NepaliDate(y, m, d);
            return true;
        }
        catch
        {
            return false;
        }
    }

    [ExcelFunction(Name = "EX_IsValid", Description = "Checks whether a text string represents a valid parseable Nepali date.", Category = CategoryName)]
    public static bool EX_IsValid(
        [ExcelArgument(AllowReference = true, Description = "Date text to validate.")] object dateText)
    {
        try
        {
            var val = NepDateHelper.ResolveValue(dateText);
            if (val is string s && !string.IsNullOrWhiteSpace(s))
            {
                return SmartDateParser.TryParse(s.Trim(), out _) || NepaliDate.TryParse(s.Trim(), out _);
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    #endregion

    #region Date Arithmetic

    [ExcelFunction(Name = "EX_AddDays", Description = "Adds or subtracts days from a Nepali date.", Category = CategoryName)]
    public static string EX_AddDays(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Number of days to add (can be negative or fractional).")] object days,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        double d = NepDateHelper.ResolveDouble(days);
        return NepDateHelper.FormatResult(date.AddDays(d), format);
    }

    [ExcelFunction(Name = "EX_AddMonths", Description = "Adds or subtracts whole or fractional months from a Nepali date.", Category = CategoryName)]
    public static string EX_AddMonths(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Number of months to add (can be negative or fractional).")] object months,
        [ExcelArgument(Description = "Optional awayFromMonthEnd adjustment (TRUE/FALSE). Default is FALSE.")] object? awayFromMonthEnd = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        double m = NepDateHelper.ResolveDouble(months);
        bool away = NepDateHelper.ResolveBool(awayFromMonthEnd, false);
        return NepDateHelper.FormatResult(date.AddMonths(m, away), format);
    }

    [ExcelFunction(Name = "EX_AddYears", Description = "Adds or subtracts years from a Nepali date.", Category = CategoryName)]
    public static string EX_AddYears(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Number of years to add (can be negative).")] object years,
        [ExcelArgument(Description = "Optional awayFromMonthEnd adjustment (TRUE/FALSE). Default is FALSE.")] object? awayFromMonthEnd = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        int y = NepDateHelper.ResolveInt(years);
        bool away = NepDateHelper.ResolveBool(awayFromMonthEnd, false);
        return NepDateHelper.FormatResult(date.AddYears(y, away), format);
    }

    [ExcelFunction(Name = "EX_DateDiff", Description = "Calculates the difference in days between two Nepali dates (endDate - startDate).", Category = CategoryName)]
    public static int EX_DateDiff(
        [ExcelArgument(AllowReference = true, Description = "Start Nepali date string or reference.")] object startDate,
        [ExcelArgument(AllowReference = true, Description = "End Nepali date string or reference.")] object endDate)
    {
        var start = NepDateHelper.ResolveNepaliDate(startDate);
        var end = NepDateHelper.ResolveNepaliDate(endDate);
        return (int)end.Subtract(start).TotalDays;
    }

    #endregion

    #region Formatting & Localization

    [ExcelFunction(Name = "EX_Format", Description = "Formats a Nepali date using custom format specifiers (e.g. yyyy-MM-dd, dd MMMM yyyy, MMMM dd, etc.).", Category = CategoryName)]
    public static string EX_Format(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Custom format string (e.g. 'yyyy-MM-dd', 'dd MMMM yyyy', 'MMMM dd, yyyy').")] object format)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        return NepDateHelper.FormatResult(date, format);
    }

    [ExcelFunction(Name = "EX_ToUnicodeDate", Description = "Formats a Nepali date using Nepali Devanagari numerals (e.g. २०८०/०१/१५).", Category = CategoryName)]
    public static string EX_ToUnicodeDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Date format order: 'YearMonthDay', 'DayMonthYear', 'MonthDayYear'. Default is YearMonthDay.")] object? format = null,
        [ExcelArgument(Description = "Separator: '/', '-', '.', '_', or space. Default is '-'.")] object? separator = null,
        [ExcelArgument(Description = "Include leading zeros (TRUE/FALSE). Default is TRUE.")] object? leadingZeros = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        var fmt = NepDateHelper.ResolveDateFormat(format);
        var sep = NepDateHelper.ResolveSeparator(separator);
        bool lz = NepDateHelper.ResolveBool(leadingZeros, true);
        return date.ToUnicodeString(fmt, sep, lz);
    }

    [ExcelFunction(Name = "EX_ToLongDate", Description = "Formats a Nepali date in long English text (e.g. 'Friday, Baishakh 01, 2080').", Category = CategoryName)]
    public static string EX_ToLongDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Include leading zeros (TRUE/FALSE). Default is TRUE.")] object? leadingZeros = null,
        [ExcelArgument(Description = "Display day of week name (TRUE/FALSE). Default is TRUE.")] object? displayDayName = null,
        [ExcelArgument(Description = "Display year (TRUE/FALSE). Default is TRUE.")] object? displayYear = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        bool lz = NepDateHelper.ResolveBool(leadingZeros, true);
        bool dn = NepDateHelper.ResolveBool(displayDayName, true);
        bool dy = NepDateHelper.ResolveBool(displayYear, true);
        return date.ToLongDateString(lz, dn, dy);
    }

    [ExcelFunction(Name = "EX_ToLongDateUnicode", Description = "Formats a Nepali date in long Nepali Devanagari text (e.g. 'शुक्रवार, बैशाख ०१, २०८०').", Category = CategoryName)]
    public static string EX_ToLongDateUnicode(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Include leading zeros (TRUE/FALSE). Default is TRUE.")] object? leadingZeros = null,
        [ExcelArgument(Description = "Display day of week name (TRUE/FALSE). Default is TRUE.")] object? displayDayName = null,
        [ExcelArgument(Description = "Display year (TRUE/FALSE). Default is TRUE.")] object? displayYear = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        bool lz = NepDateHelper.ResolveBool(leadingZeros, true);
        bool dn = NepDateHelper.ResolveBool(displayDayName, true);
        bool dy = NepDateHelper.ResolveBool(displayYear, true);
        return date.ToLongDateUnicodeString(lz, dn, dy);
    }

    #endregion

    #region Calendar Metadata (Tithi & Events)

    [ExcelFunction(Name = "EX_Tithi", Description = "Returns the lunar Tithi for a Nepali date (available for 2001–2089 B.S.).", Category = CategoryName)]
    public static string EX_Tithi(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Return Tithi in Nepali Devanagari (TRUE) or English (FALSE). Default is FALSE (English).")] object? inNepali = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        bool np = NepDateHelper.ResolveBool(inNepali, false);
        return np ? date.TithiNp : date.TithiEn;
    }

    [ExcelFunction(Name = "EX_IsPublicHoliday", Description = "Checks if the Nepali date is an official public holiday in Nepal (available for 2001–2089 B.S.).", Category = CategoryName)]
    public static bool EX_IsPublicHoliday(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        return NepDateHelper.ResolveNepaliDate(nepaliDate).IsPublicHoliday;
    }

    [ExcelFunction(Name = "EX_Events", Description = "Returns festivals, events, or observances on the given Nepali date (available for 2001–2089 B.S.).", Category = CategoryName)]
    public static string EX_Events(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Return events in Nepali Devanagari (TRUE) or English (FALSE). Default is FALSE (English).")] object? inNepali = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        bool np = NepDateHelper.ResolveBool(inNepali, false);
        var events = np ? date.EventsNp : date.EventsEn;
        return events != null && events.Length > 0 ? string.Join(", ", events) : string.Empty;
    }

    #endregion

    #region Fiscal Year & Quarters

    [ExcelFunction(Name = "EX_FiscalYear", Description = "Returns the Nepal Fiscal Year string for a given Nepali date (e.g. '2080/81').", Category = CategoryName)]
    public static string EX_FiscalYear(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        int startYear = date.Month >= 4 ? date.Year : date.Year - 1;
        int endYearShort = (startYear + 1) % 100;
        return $"{startYear}/{endYearShort:D2}";
    }

    [ExcelFunction(Name = "EX_FiscalQuarter", Description = "Returns the Nepal Fiscal Quarter number (1: Shrawan–Ashoj, 2: Kartik–Poush, 3: Magh–Chaitra, 4: Baishakh–Ashadh).", Category = CategoryName)]
    public static int EX_FiscalQuarter(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        return date.Month switch
        {
            >= 4 and <= 6 => 1,
            >= 7 and <= 9 => 2,
            >= 10 and <= 12 => 3,
            _ => 4
        };
    }

    [ExcelFunction(Name = "EX_FiscalYearStartDate", Description = "Returns the start date of the Nepal Fiscal Year (Shrawan 1).", Category = CategoryName)]
    public static string EX_FiscalYearStartDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or fiscal year number (e.g. 2080).")] object nepaliDateOrYear,
        [ExcelArgument(Description = "Optional fiscal year offset (-1 for previous, +1 for next). Default is 0.")] object? yearOffset = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var val = NepDateHelper.ResolveValue(nepaliDateOrYear);
        int offset = NepDateHelper.ResolveInt(yearOffset, 0);
        NepaliDate result;

        if (val is int y || (val is double d && d >= 1900 && d <= 2200 && Math.Abs(d - Math.Round(d)) < 0.001) || (val is string s && int.TryParse(s, out y) && y >= 1900 && y <= 2200))
        {
            int yearVal = val is int i ? i : (val is double db ? (int)db : int.Parse((string)val!));
            result = NepaliDate.GetFiscalYearStartDate(yearVal + offset);
        }
        else
        {
            var date = NepDateHelper.ResolveNepaliDate(nepaliDateOrYear);
            result = date.FiscalYearStartDate(offset);
        }

        return NepDateHelper.FormatResult(result, format);
    }

    [ExcelFunction(Name = "EX_FiscalYearEndDate", Description = "Returns the end date of the Nepal Fiscal Year (Ashadh end).", Category = CategoryName)]
    public static string EX_FiscalYearEndDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or fiscal year number (e.g. 2080).")] object nepaliDateOrYear,
        [ExcelArgument(Description = "Optional fiscal year offset (-1 for previous, +1 for next). Default is 0.")] object? yearOffset = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var val = NepDateHelper.ResolveValue(nepaliDateOrYear);
        int offset = NepDateHelper.ResolveInt(yearOffset, 0);
        NepaliDate result;

        if (val is int y || (val is double d && d >= 1900 && d <= 2200 && Math.Abs(d - Math.Round(d)) < 0.001) || (val is string s && int.TryParse(s, out y) && y >= 1900 && y <= 2200))
        {
            int yearVal = val is int i ? i : (val is double db ? (int)db : int.Parse((string)val!));
            result = NepaliDate.GetFiscalYearEndDate(yearVal + offset);
        }
        else
        {
            var date = NepDateHelper.ResolveNepaliDate(nepaliDateOrYear);
            result = date.FiscalYearEndDate(offset);
        }

        return NepDateHelper.FormatResult(result, format);
    }

    [ExcelFunction(Name = "EX_FiscalQuarterStartDate", Description = "Returns the start date of the specified fiscal quarter.", Category = CategoryName)]
    public static string EX_FiscalQuarterStartDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Fiscal quarter (1–4, 'Q1'–'Q4', or 0 for current). Default is current quarter.")] object? quarter = null,
        [ExcelArgument(Description = "Optional fiscal year offset. Default is 0.")] object? yearOffset = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        var q = NepDateHelper.ResolveFiscalQuarter(quarter);
        int offset = NepDateHelper.ResolveInt(yearOffset, 0);
        var result = date.FiscalYearQuarterStartDate(q, offset);
        return NepDateHelper.FormatResult(result, format);
    }

    [ExcelFunction(Name = "EX_FiscalQuarterEndDate", Description = "Returns the end date of the specified fiscal quarter.", Category = CategoryName)]
    public static string EX_FiscalQuarterEndDate(
        [ExcelArgument(AllowReference = true, Description = "Nepali date string or reference.")] object nepaliDate,
        [ExcelArgument(Description = "Fiscal quarter (1–4, 'Q1'–'Q4', or 0 for current). Default is current quarter.")] object? quarter = null,
        [ExcelArgument(Description = "Optional fiscal year offset. Default is 0.")] object? yearOffset = null,
        [ExcelArgument(Description = "Optional output format. Default is yyyy-MM-dd.")] object? format = null)
    {
        var date = NepDateHelper.ResolveNepaliDate(nepaliDate);
        var q = NepDateHelper.ResolveFiscalQuarter(quarter);
        int offset = NepDateHelper.ResolveInt(yearOffset, 0);
        var result = date.FiscalYearQuarterEndDate(q, offset);
        return NepDateHelper.FormatResult(result, format);
    }

    #endregion

    #region Date Ranges & Working Days

    [ExcelFunction(Name = "EX_DaysBetween", Description = "Calculates the total number of calendar days in a Nepali date range (inclusive).", Category = CategoryName)]
    public static int EX_DaysBetween(
        [ExcelArgument(AllowReference = true, Description = "Start Nepali date string.")] object startDate,
        [ExcelArgument(AllowReference = true, Description = "End Nepali date string.")] object endDate)
    {
        var start = NepDateHelper.ResolveNepaliDate(startDate);
        var end = NepDateHelper.ResolveNepaliDate(endDate);
        var range = new NepaliDateRange(start, end);
        return range.Length;
    }

    [ExcelFunction(Name = "EX_WorkingDays", Description = "Counts working days between two Nepali dates (excludes Saturdays by default in Nepal, optionally excludes Sundays).", Category = CategoryName)]
    public static int EX_WorkingDays(
        [ExcelArgument(AllowReference = true, Description = "Start Nepali date string.")] object startDate,
        [ExcelArgument(AllowReference = true, Description = "End Nepali date string.")] object endDate,
        [ExcelArgument(Description = "Also exclude Sundays from working days (TRUE/FALSE). Default is FALSE.")] object? excludeSunday = null)
    {
        var start = NepDateHelper.ResolveNepaliDate(startDate);
        var end = NepDateHelper.ResolveNepaliDate(endDate);
        var range = new NepaliDateRange(start, end);
        bool exSun = NepDateHelper.ResolveBool(excludeSunday, false);
        return range.WorkingDays(exSun).Count();
    }

    [ExcelFunction(Name = "EX_WeekendDays", Description = "Counts weekend days between two Nepali dates (Saturdays, and optionally Sundays).", Category = CategoryName)]
    public static int EX_WeekendDays(
        [ExcelArgument(AllowReference = true, Description = "Start Nepali date string.")] object startDate,
        [ExcelArgument(AllowReference = true, Description = "End Nepali date string.")] object endDate,
        [ExcelArgument(Description = "Also count Sundays as weekend days (TRUE/FALSE). Default is FALSE.")] object? includeSunday = null)
    {
        var start = NepDateHelper.ResolveNepaliDate(startDate);
        var end = NepDateHelper.ResolveNepaliDate(endDate);
        var range = new NepaliDateRange(start, end);
        bool incSun = NepDateHelper.ResolveBool(includeSunday, false);
        return range.WeekendDays(incSun).Count();
    }

    [ExcelFunction(Name = "EX_RangeContains", Description = "Checks if a check date falls within the Nepali date range [startDate, endDate].", Category = CategoryName)]
    public static bool EX_RangeContains(
        [ExcelArgument(AllowReference = true, Description = "Start Nepali date string.")] object startDate,
        [ExcelArgument(AllowReference = true, Description = "End Nepali date string.")] object endDate,
        [ExcelArgument(AllowReference = true, Description = "Nepali date to check.")] object checkDate)
    {
        var start = NepDateHelper.ResolveNepaliDate(startDate);
        var end = NepDateHelper.ResolveNepaliDate(endDate);
        var check = NepDateHelper.ResolveNepaliDate(checkDate);
        var range = new NepaliDateRange(start, end);
        return range.Contains(check);
    }

    [ExcelFunction(Name = "EX_RangeOverlaps", Description = "Checks if two Nepali date ranges overlap.", Category = CategoryName)]
    public static bool EX_RangeOverlaps(
        [ExcelArgument(AllowReference = true, Description = "Start of first range.")] object start1,
        [ExcelArgument(AllowReference = true, Description = "End of first range.")] object end1,
        [ExcelArgument(AllowReference = true, Description = "Start of second range.")] object start2,
        [ExcelArgument(AllowReference = true, Description = "End of second range.")] object end2)
    {
        var s1 = NepDateHelper.ResolveNepaliDate(start1);
        var e1 = NepDateHelper.ResolveNepaliDate(end1);
        var s2 = NepDateHelper.ResolveNepaliDate(start2);
        var e2 = NepDateHelper.ResolveNepaliDate(end2);
        var r1 = new NepaliDateRange(s1, e1);
        var r2 = new NepaliDateRange(s2, e2);
        return r1.Overlaps(r2);
    }

    #endregion
}
