# ExNepDate: Nepali Date (Bikram Sambat) Add-in for Microsoft Excel

**ExNepDate** is an Excel Custom Functions (UDF) library built with [Excel-DNA](https://excel-dna.net/) and powered by the [.NET NepDate](https://github.com/nepdate) library. It brings seamless, high-performance Bikram Sambat (B.S.) date conversions, calendar arithmetic, Nepal fiscal year calculations, working day logic, and lunar Tithi metadata directly into Microsoft Excel spreadsheets.

---

## Table of Contents

- [Features](#features)
- [Installation & Setup](#installation--setup)
- [Function Reference](#function-reference)
  - [Core & Conversions](#1-core--conversions)
  - [Date Components & Getters](#2-date-components--getters)
  - [Validation & Verification](#3-validation--verification)
  - [Date Arithmetic](#4-date-arithmetic)
  - [Formatting & Localization](#5-formatting--localization)
  - [Calendar Metadata (Tithi & Events)](#6-calendar-metadata-tithi--events)
  - [Nepal Fiscal Year & Quarters](#7-nepal-fiscal-year--quarters)
  - [Date Ranges & Working Days](#8-date-ranges--working-days)
- [Formatting Options](#formatting-options)
- [License & Dependencies](#license--dependencies)

---

## Features

- **Bidirectional Conversions**: Easily convert between Gregorian (A.D.) dates and Bikram Sambat (B.S.) dates.
- **Smart Parsing**: Recognizes over 100+ Nepali month spellings, Unicode Devanagari numerals (e.g., `२०८०/०१/१५`), and standard formatted strings.
- **Nepal Fiscal Year Calculations**: Built-in support for Nepal fiscal years ($2080/81$), quarter calculations (Q1 Shrawan–Ashoj through Q4 Baishakh–Ashadh), and automatic period start/end dates.
- **Local Business Logic**: Calculate working days and weekends accounting for Nepal's Saturday holiday convention (with optional Sunday inclusion).
- **Lunar Calendar Metadata**: Fetch Tithi, official public holiday status, and cultural festival events for B.S. dates ($2001–2089\text{ B.S.}$).

---

## Installation & Setup

### Prerequisites

- Microsoft Excel (2013 or newer, 32-bit or 64-bit)
- .NET Framework / .NET Runtime (matching your Excel-DNA target build)

### Loading the Add-in

1. Build the project to output the compiled `.xll` add-in file (`ExNepDate-AddIn.xll` or `ExNepDate-AddIn64.xll`).
2. Open Microsoft Excel.
3. Go to **File** $\rightarrow$ **Options** $\rightarrow$ **Add-Ins**.
4. At the bottom, select **Excel Add-ins** from the *Manage* dropdown and click **Go...**.
5. Click **Browse...**, navigate to your output folder, and select the `.xll` file.
6. Click **OK**. The functions will now be available in Excel under the **NepDate** category.

---

## Function Reference

### 1. Core & Conversions

| Function | Description | Example Syntax |
| :--- | :--- | :--- |
| `EX_Today([format])` | Returns today's Nepali date. | `=EX_Today("yyyy-MM-dd")` |
| `EX_Now([format])` | Returns the current Nepali date. | `=EX_Now()` |
| `EX_NepaliDate(date, [format])` | Converts an English (Gregorian) date to Nepali B.S. date. | `=EX_NepaliDate(A2, "yyyy/MM/dd")` |
| `EX_ToNepaliDate(date, [format])` | Alias for `EX_NepaliDate`. | `=EX_ToNepaliDate("2023-04-14")` |
| `EX_ToEnglishDate(nepaliDate, [format])` | Converts a Nepali B.S. date to an English Gregorian date. Returns native Excel date serial or formatted string. | `=EX_ToEnglishDate("2080-01-01")` |
| `EX_EnglishDate(nepaliDate, [format])` | Alias for `EX_ToEnglishDate`. | `=EX_EnglishDate("2080/01/15")` |
| `EX_Date(year, month, day, [format])` | Creates a Nepali date from year, month, and day integers. | `=EX_Date(2080, 1, 15)` |
| `EX_Parse(dateText, [format])` | Smart-parses text (English, Devanagari, or month spellings) into a normalized date. | `=EX_Parse("15 Baishakh 2080")` |

---

### 2. Date Components & Getters

| Function | Description | Example |
| :--- | :--- | :--- |
| `EX_Year(nepaliDate)` | Returns the year integer (e.g., `2080`). | `=EX_Year("2080-01-15")` $\rightarrow$ `2080` |
| `EX_Month(nepaliDate)` | Returns month number ($1 \le m \le 12$). | `=EX_Month("2080-01-15")` $\rightarrow$ `1` |
| `EX_Day(nepaliDate)` | Returns day of the month ($1 \le d \le 32$). | `=EX_Day("2080-01-15")` $\rightarrow$ `15` |
| `EX_MonthName(nepaliDate)` | Returns English name of the Nepali month. | `=EX_MonthName("2080-01-15")` $\rightarrow$ `"Baishakh"` |
| `EX_DayOfWeek(nepaliDate)` | Returns weekday name (e.g., `Friday`). | `=EX_DayOfWeek("2080-01-01")` $\rightarrow$ `"Friday"` |
| `EX_DayOfYear(nepaliDate)` | Returns ordinal day within the year ($1 \le d \le 366$). | `=EX_DayOfYear("2080-01-15")` $\rightarrow$ `15` |
| `EX_MonthEndDay(nepaliDate)` | Returns total days in the month (e.g., `30`, `31`, `32`). | `=EX_MonthEndDay("2080-01-01")` $\rightarrow$ `31` |
| `EX_MonthEndDate(nepaliDate, [format])` | Returns the last calendar date of the month. | `=EX_MonthEndDate("2080-01-15")` $\rightarrow$ `"2080-01-31"` |

---

### 3. Validation & Verification

| Function | Description | Output |
| :--- | :--- | :--- |
| `EX_IsLeapYear(nepaliDateOrYear)` | Checks if a Nepali year or date is a leap year. | `TRUE` / `FALSE` |
| `EX_IsToday(nepaliDate)` | Checks if the given date is today's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsTomorrow(nepaliDate)` | Checks if the given date is tomorrow's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsYesterday(nepaliDate)` | Checks if the given date was yesterday's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsValidDate(year, month, day)` | Validates whether integer Y, M, D form a valid B.S. date. | `TRUE` / `FALSE` |
| `EX_IsValid(dateText)` | Checks if a text string is a valid parseable B.S. date. | `TRUE` / `FALSE` |

---

### 4. Date Arithmetic

| Function | Description | Parameters |
| :--- | :--- | :--- |
| `EX_AddDays(nepaliDate, days, [format])` | Adds or subtracts $n$ days from a date. | `days` can be positive, negative, or fractional. |
| `EX_AddMonths(nepaliDate, months, [awayFromMonthEnd], [format])` | Adds or subtracts whole/fractional months. | `awayFromMonthEnd`: `TRUE`/`FALSE` adjustment rule. |
| `EX_AddYears(nepaliDate, years, [awayFromMonthEnd], [format])` | Adds or subtracts years. | `years`: Integer count. |
| `EX_DateDiff(startDate, endDate)` | Returns difference in days (`endDate - startDate`). | Returns integer count of days. |

---

### 5. Formatting & Localization

| Function | Description | Output Example |
| :--- | :--- | :--- |
| `EX_Format(nepaliDate, format)` | Custom format output using specifiers (`yyyy-MM-dd`, `dd MMMM yyyy`). | `"15 Baishakh 2080"` |
| `EX_ToUnicodeDate(nepaliDate, [format], [separator], [leadingZeros])` | Formats in Devanagari numerals. | `"२०८०/०१/१५"` |
| `EX_ToLongDate(nepaliDate, [leadingZeros], [displayDayName], [displayYear])` | Formats long English date text. | `"Friday, Baishakh 01, 2080"` |
| `EX_ToLongDateUnicode(...)` | Formats long Devanagari text string. | `"शुक्रवार, बैशाख ०१, २०८०"` |

---

### 6. Calendar Metadata (Tithi & Events)

> **Note**: Metadata coverage is available for **2001 B.S. to 2089 B.S.**

| Function | Description | Parameters / Output |
| :--- | :--- | :--- |
| `EX_Tithi(nepaliDate, [inNepali])` | Returns the lunar Tithi for the date. | `inNepali = TRUE` for Devanagari, `FALSE` for English. |
| `EX_IsPublicHoliday(nepaliDate)` | Returns `TRUE` if the date is an official public holiday in Nepal. | `TRUE` / `FALSE` |
| `EX_Events(nepaliDate, [inNepali])` | Returns comma-separated list of festivals/events on that date. | e.g. `"Dashain, Vijaya Dashami"` |

---

### 7. Nepal Fiscal Year & Quarters

Nepal's fiscal year runs from **Shrawan 1** (Mid-July) to **Ashadh end** (Mid-July of the following year).

- **Quarter 1**: Shrawan – Ashoj
- **Quarter 2**: Kartik – Poush
- **Quarter 3**: Magh – Chaitra
- **Quarter 4**: Baishakh – Ashadh

| Function | Description | Example Output |
| :--- | :--- | :--- |
| `EX_FiscalYear(nepaliDate)` | Returns the Nepal FY string representation. | `=EX_FiscalYear("2080-05-10")` $\rightarrow$ `"2080/81"` |
| `EX_FiscalQuarter(nepaliDate)` | Returns the quarter number ($1, 2, 3,$ or $4$). | `=EX_FiscalQuarter("2080-05-10")` $\rightarrow$ `1` |
| `EX_FiscalYearStartDate(nepaliDateOrYear, [yearOffset], [format])` | Returns the start date of the FY (Shrawan 1). | `"2080-04-01"` |
| `EX_FiscalYearEndDate(nepaliDateOrYear, [yearOffset], [format])` | Returns the end date of the FY (Ashadh end). | `"2081-03-31"` |
| `EX_FiscalQuarterStartDate(nepaliDate, [quarter], [yearOffset], [format])` | Start date of specified quarter ($1–4$ or current). | `"2080-04-01"` |
| `EX_FiscalQuarterEndDate(nepaliDate, [quarter], [yearOffset], [format])` | End date of specified quarter ($1–4$ or current). | `"2080-06-31"` |

---

### 8. Date Ranges & Working Days

| Function | Description | Notes |
| :--- | :--- | :--- |
| `EX_DaysBetween(startDate, endDate)` | Returns total calendar days in range (inclusive). | Inclusive count: $[start, end]$. |
| `EX_WorkingDays(startDate, endDate, [excludeSunday])` | Calculates business days excluding Saturdays. | Set `excludeSunday = TRUE` to exclude 2-day weekend. |
| `EX_WeekendDays(startDate, endDate, [includeSunday])` | Counts weekend days in the range. | Saturdays only by default; include Sundays if set. |
| `EX_RangeContains(startDate, endDate, checkDate)` | Checks if `checkDate` lies within range. | Returns `TRUE` / `FALSE`. |
| `EX_RangeOverlaps(start1, end1, start2, end2)` | Checks if two date ranges overlap. | Returns `TRUE` / `FALSE`. |

---

## Formatting Options

When passing format specifiers to functions such as `EX_Format` or optional `format` arguments:

| Specifier | Description | Example Output |
| :--- | :--- | :--- |
| `yyyy` | 4-digit Nepali Year | `2080` |
| `MM` | 2-digit Nepali Month | `01` |
| `dd` | 2-digit Nepali Day | `15` |
| `MMMM` | Full English month name | `Baishakh` |
| `dddd` | Full English day of week name | `Friday` |

---

## License & Dependencies

- **Excel-DNA**: Licensed under the [Excel-DNA License](https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt).
- **NepDate**: Native .NET library for Nepali Date conversions and calendar features.