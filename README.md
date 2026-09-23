# 🇳🇵 ExNepDate: Nepali Date (Bikram Sambat) Add-in for Microsoft Excel

<p align="center">
  <a href="https://github.com/bishwokarmalaxman/ExNepDate/releases/latest">
    <img src="https://img.shields.io/github/v/release/bishwokarmalaxman/ExNepDate?color=0078D4&label=Latest%20Release&logo=github" alt="Latest Release" />
  </a>
  <a href="LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License" />
  </a>
  <img src="https://img.shields.io/badge/Excel-2013%20--%20365-217346?logo=microsoft-excel&logoColor=white" alt="Microsoft Excel" />
  <img src="https://img.shields.io/badge/Platform-Windows%20(32%2F64--bit)-0078D6?logo=windows&logoColor=white" alt="Platform: Windows" />
  <a href="https://github.com/Excel-DNA/ExcelDna">
    <img src="https://img.shields.io/badge/Built%20with-Excel--DNA-orange" alt="Excel-DNA" />
  </a>
  <a href="https://github.com/nepdate">
    <img src="https://img.shields.io/badge/Powered%20by-NepDate%20v2.0.7-blueviolet" alt="NepDate" />
  </a>
</p>

**ExNepDate** is a fast, lightweight, and full-featured Excel Custom Functions (UDF) library for Microsoft Excel. Built with [Excel-DNA](https://excel-dna.net/) and powered by [.NET NepDate](https://github.com/nepdate), it brings native **Bikram Sambat (B.S. / वि.सं.)** date conversions, calendar arithmetic, Nepal fiscal year calculations, business working day analysis, and lunar Tithi metadata right into your Excel spreadsheets.

---

## 📑 Table of Contents

- [📥 Downloads](#-downloads)
- [✨ Key Features](#-key-features)
- [⚡ Quick Formula Cheat Sheet](#-quick-formula-cheat-sheet)
- [🚀 Installation & Setup](#-installation--setup)
  - [Prerequisites](#prerequisites)
  - [Option 1: Automatic Setup via MSI (Recommended)](#option-1-automatic-setup-via-msi-installer-recommended)
  - [Option 2: Standalone XLL Add-in](#option-2-manual-setup-via-standalone-xll-portable)
- [📖 Function Reference](#-function-reference)
  - [🔄 1. Core & Date Conversion](#1-core--date-conversion)
  - [🧩 2. Date Components & Extraction](#2-date-components--extraction)
  - [✅ 3. Date Validation & Verification](#3-date-validation--verification)
  - [➕ 4. Date Arithmetic & Differences](#4-date-arithmetic--differences)
  - [🎨 5. Formatting & Localization](#5-formatting--localization)
  - [🌕 6. Lunar Calendar, Tithi & Holidays](#6-lunar-calendar-tithi--events)
  - [💼 7. Nepal Fiscal Year (FY) & Quarters](#7-nepal-fiscal-year--quarters)
  - [📊 8. Date Ranges & Working Days](#8-date-ranges--working-days)
- [🏷️ Formatting Options](#️-formatting-options)
- [🛠️ Building from Source](#️-building-from-source)
- [📄 License & Credits](#-license--credits)

---

## 📥 Downloads

Get the latest stable release (**v1.0.0**):

| Package | Type | Description | Download Link |
| :--- | :--- | :--- | :--- |
| **`ExcelNepDateAddIn.msi`** | 🪟 Windows Installer (`.msi`) | **Recommended** • Automatic setup & registration in Microsoft Excel | [📥 **Download MSI**](https://github.com/bishwokarmalaxman/ExNepDate/releases/download/v1.0.0/ExcelNepDateAddIn.msi) |
| **`ExNepDate-AddIn64-packed.xll`** | ⚡ Standalone Add-in (`.xll`) | **Portable** • Pre-packaged 64-bit Excel Add-in (No installation needed) | [📥 **Download XLL (64-bit)**](https://github.com/bishwokarmalaxman/ExNepDate/releases/download/v1.0.0/ExNepDate-AddIn64-packed.xll) |

> 📌 All releases, checksums, and source bundles are available on the [GitHub Releases](https://github.com/bishwokarmalaxman/ExNepDate/releases) page.

---

## ✨ Key Features

- 🔄 **Bidirectional Conversions**: Seamlessly convert English (Gregorian A.D.) dates $\leftrightarrow$ Nepali (Bikram Sambat B.S.) dates.
- 🧠 **Smart Nepali Text & Devanagari Parser**: Recognizes over 100+ Nepali month spellings, English date formats, and Unicode Devanagari numbers (e.g. `२०८०/०१/१५`).
- 💼 **Nepal Fiscal Year Engine**: Built-in support for Nepal fiscal years (e.g. `2080/81`), quarterly periods (Q1 Shrawan–Ashoj to Q4 Baishakh–Ashadh), and automatic period start/end dates.
- 🏢 **Nepal Business Calendar**: Calculates working days and weekends tailored to Nepal's standard 6-day work week (Saturday holidays, with optional Sunday exclusion).
- 🌕 **Lunar Calendar & Tithi**: Lookup lunar Tithi, official public holidays, and religious/national festivals ($2001–2089\text{ B.S.}$).
- 💡 **Excel IntelliSense Integration**: Auto-complete and in-cell documentation popups as you type `=EX_...`.

---

## ⚡ Quick Formula Cheat Sheet

Here are the most popular formulas to copy and paste right away:

| Formula | Description | Sample Output |
| :--- | :--- | :--- |
| `=EX_Today()` | Today's Nepali date (B.S.) | `2080-06-07` |
| `=EX_ToNepaliDate(A2)` | Convert A.D. date in cell `A2` to B.S. | `2080-01-01` |
| `=EX_ToEnglishDate("2080-01-01")` | Convert B.S. date to native Excel A.D. date | `2023-04-14` |
| `=EX_FiscalYear("2080-05-15")` | Get current Nepal Fiscal Year | `2080/81` |
| `=EX_WorkingDays("2080-01-01", "2080-01-31")` | Calculate Nepal business days (excl. Saturdays) | `26` |
| `=EX_ToUnicodeDate("2080-01-15")` | Convert date to Nepali Devanagari numerals | `२०८०/०१/१५` |
| `=EX_Tithi("2080-07-07", TRUE)` | Get lunar Tithi in Nepali | `दशमी` |

---

## 🚀 Installation & Setup

### Prerequisites

- **Operating System**: Windows 10, 11, or Windows Server (32-bit or 64-bit)
- **Microsoft Excel**: Excel 2013, 2016, 2019, 2021, or Microsoft 365 (Desktop version)
- **Runtime**: .NET Runtime / .NET Framework (included with modern Windows updates)

---

### Option 1: Automatic Setup via MSI Installer (Recommended)

1. Download [**`ExcelNepDateAddIn.msi`**](https://github.com/bishwokarmalaxman/ExNepDate/releases/download/v1.0.0/ExcelNepDateAddIn.msi).
2. Double-click the `.msi` file and follow the quick setup wizard.
3. Open or restart Microsoft Excel.
4. In any cell, type `=EX_` — all functions are immediately ready to use!

---

### Option 2: Manual Setup via Standalone XLL (Portable)

1. Download [**`ExNepDate-AddIn64-packed.xll`**](https://github.com/bishwokarmalaxman/ExNepDate/releases/download/v1.0.0/ExNepDate-AddIn64-packed.xll) (or build locally).
2. Save the `.xll` file to a stable folder (e.g. `Documents\ExcelAddins\`).
3. Open Microsoft Excel.
4. Go to **File** $\rightarrow$ **Options** $\rightarrow$ **Add-Ins**.
5. At the bottom, ensure **Excel Add-ins** is selected in the *Manage* dropdown, then click **Go...**.
6. In the dialog, click **Browse...**, navigate to your downloaded `.xll` file, and click **OK**.
7. Ensure **ExNepDate-AddIn64** is checked in the list, then click **OK**.

> 💡 **Tip**: If you are using 32-bit Excel, compile the 32-bit build or use the MSI installer which automatically handles architecture configuration.

---

## 📖 Function Reference

All functions start with the prefix `EX_` and are registered under the **NepDate** category in Excel's Function Wizard (`fx`).

---

### 🔄 1. Core & Date Conversion

| Function | Description | Example Syntax |
| :--- | :--- | :--- |
| `EX_Today([format])` | Returns today's Nepali date. | `=EX_Today("yyyy-MM-dd")` |
| `EX_Now([format])` | Returns the current Nepali date and time. | `=EX_Now()` |
| `EX_NepaliDate(date, [format])` | Converts an English (Gregorian) date to Nepali B.S. date. | `=EX_NepaliDate(A2, "yyyy/MM/dd")` |
| `EX_ToNepaliDate(date, [format])` | Convenient alias for `EX_NepaliDate`. | `=EX_ToNepaliDate("2023-04-14")` |
| `EX_ToEnglishDate(nepaliDate, [format])` | Converts a Nepali B.S. date to an English Gregorian date. Returns native Excel date serial or formatted string. | `=EX_ToEnglishDate("2080-01-01")` |
| `EX_EnglishDate(nepaliDate, [format])` | Convenient alias for `EX_ToEnglishDate`. | `=EX_EnglishDate("2080/01/15")` |
| `EX_Date(year, month, day, [format])` | Creates a Nepali date from year, month, and day integers. | `=EX_Date(2080, 1, 15)` |
| `EX_Parse(dateText, [format])` | Smart-parses text (English, Devanagari, or month spellings) into a normalized date. | `=EX_Parse("15 Baishakh 2080")` |

---

### 🧩 2. Date Components & Extraction

| Function | Description | Example & Output |
| :--- | :--- | :--- |
| `EX_Year(nepaliDate)` | Extracts the Nepali year as an integer. | `=EX_Year("2080-01-15")` $\rightarrow$ `2080` |
| `EX_Month(nepaliDate)` | Extracts the month number ($1 \le m \le 12$). | `=EX_Month("2080-01-15")` $\rightarrow$ `1` |
| `EX_Day(nepaliDate)` | Extracts the day of the month ($1 \le d \le 32$). | `=EX_Day("2080-01-15")` $\rightarrow$ `15` |
| `EX_MonthName(nepaliDate)` | Returns the English name of the Nepali month. | `=EX_MonthName("2080-01-15")` $\rightarrow$ `"Baishakh"` |
| `EX_DayOfWeek(nepaliDate)` | Returns the day of the week (e.g., `Friday`). | `=EX_DayOfWeek("2080-01-01")` $\rightarrow$ `"Friday"` |
| `EX_DayOfYear(nepaliDate)` | Returns the ordinal day within the year ($1 \le d \le 366$). | `=EX_DayOfYear("2080-01-15")` $\rightarrow$ `15` |
| `EX_MonthEndDay(nepaliDate)` | Returns total days in that Nepali month (e.g., `30`, `31`, `32`). | `=EX_MonthEndDay("2080-01-01")` $\rightarrow$ `31` |
| `EX_MonthEndDate(nepaliDate, [format])` | Returns the full date of the last day of the month. | `=EX_MonthEndDate("2080-01-15")` $\rightarrow$ `"2080-01-31"` |

---

### ✅ 3. Date Validation & Verification

| Function | Description | Return Type |
| :--- | :--- | :--- |
| `EX_IsLeapYear(nepaliDateOrYear)` | Checks if a Nepali year or date falls in a leap year. | `TRUE` / `FALSE` |
| `EX_IsToday(nepaliDate)` | Checks if the given date matches today's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsTomorrow(nepaliDate)` | Checks if the given date is tomorrow's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsYesterday(nepaliDate)` | Checks if the given date was yesterday's Nepali date. | `TRUE` / `FALSE` |
| `EX_IsValidDate(year, month, day)` | Validates whether integer Y, M, D form a valid B.S. date. | `TRUE` / `FALSE` |
| `EX_IsValid(dateText)` | Checks whether a string can be parsed into a valid B.S. date. | `TRUE` / `FALSE` |

---

### ➕ 4. Date Arithmetic & Differences

| Function | Description | Parameters & Notes |
| :--- | :--- | :--- |
| `EX_AddDays(nepaliDate, days, [format])` | Adds or subtracts $n$ days from a date. | `days` can be positive or negative. |
| `EX_AddMonths(nepaliDate, months, [awayFromMonthEnd], [format])` | Adds or subtracts whole/fractional months. | `awayFromMonthEnd`: `TRUE`/`FALSE` adjustment rule. |
| `EX_AddYears(nepaliDate, years, [awayFromMonthEnd], [format])` | Adds or subtracts years. | `years`: Integer count. |
| `EX_DateDiff(startDate, endDate)` | Returns the exact difference in calendar days (`endDate - startDate`). | Returns integer count of days. |

---

### 🎨 5. Formatting & Localization

| Function | Description | Sample Output |
| :--- | :--- | :--- |
| `EX_Format(nepaliDate, format)` | Formats a B.S. date using custom pattern specifiers. | `=EX_Format("2080-01-15", "dd MMMM yyyy")` $\rightarrow$ `"15 Baishakh 2080"` |
| `EX_ToUnicodeDate(nepaliDate, [format], [separator], [leadingZeros])` | Formats the date using Nepali Devanagari numerals. | `=EX_ToUnicodeDate("2080-01-15")` $\rightarrow$ `"२०८०/०१/१५"` |
| `EX_ToLongDate(nepaliDate, [leadingZeros], [displayDayName], [displayYear])` | Produces a full English descriptive date string. | `"Friday, Baishakh 01, 2080"` |
| `EX_ToLongDateUnicode(...)` | Produces a full Devanagari descriptive date string. | `"शुक्रवार, बैशाख ०१, २०८०"` |

---

### 🌕 6. Lunar Calendar, Tithi & Events

> ℹ️ **Coverage**: Lunar metadata, Tithis, and festival events are supported from **2001 B.S. through 2089 B.S.**

| Function | Description | Output Example |
| :--- | :--- | :--- |
| `EX_Tithi(nepaliDate, [inNepali])` | Returns the lunar Tithi for the date. | `inNepali = TRUE`: `"एकादशी"`, `FALSE`: `"Ekadashi"` |
| `EX_IsPublicHoliday(nepaliDate)` | Returns `TRUE` if the date is an official public holiday in Nepal. | `TRUE` / `FALSE` |
| `EX_Events(nepaliDate, [inNepali])` | Returns comma-separated festivals/events on that date. | `"Dashain, Vijaya Dashami"` |

---

### 💼 7. Nepal Fiscal Year & Quarters

Nepal's fiscal year starts on **Shrawan 1** (mid-July) and concludes on **Ashadh end** (mid-July of the subsequent Gregorian year).

- 🥇 **Quarter 1**: Shrawan – Ashoj
- 🥈 **Quarter 2**: Kartik – Poush
- 🥉 **Quarter 3**: Magh – Chaitra
- 🏅 **Quarter 4**: Baishakh – Ashadh

| Function | Description | Example Output |
| :--- | :--- | :--- |
| `EX_FiscalYear(nepaliDate)` | Returns the Nepal FY notation string. | `=EX_FiscalYear("2080-05-10")` $\rightarrow$ `"2080/81"` |
| `EX_FiscalQuarter(nepaliDate)` | Returns the fiscal quarter ($1, 2, 3,$ or $4$). | `=EX_FiscalQuarter("2080-05-10")` $\rightarrow$ `1` |
| `EX_FiscalYearStartDate(nepaliDateOrYear, [yearOffset], [format])` | Start date of the fiscal year (Shrawan 1). | `"2080-04-01"` |
| `EX_FiscalYearEndDate(nepaliDateOrYear, [yearOffset], [format])` | End date of the fiscal year (Ashadh end). | `"2081-03-31"` |
| `EX_FiscalQuarterStartDate(nepaliDate, [quarter], [yearOffset], [format])` | Start date of specified quarter ($1–4$). | `"2080-04-01"` |
| `EX_FiscalQuarterEndDate(nepaliDate, [quarter], [yearOffset], [format])` | End date of specified quarter ($1–4$). | `"2080-06-31"` |

---

### 📊 8. Date Ranges & Working Days

| Function | Description | Notes |
| :--- | :--- | :--- |
| `EX_DaysBetween(startDate, endDate)` | Returns total calendar days in range (inclusive). | Inclusive interval: $[start, end]$. |
| `EX_WorkingDays(startDate, endDate, [excludeSunday])` | Calculates working days excluding Nepal's Saturday holiday. | Pass `excludeSunday = TRUE` to exclude 2-day weekends. |
| `EX_WeekendDays(startDate, endDate, [includeSunday])` | Counts weekend days in the date range. | Saturdays only by default; includes Sundays if specified. |
| `EX_RangeContains(startDate, endDate, checkDate)` | Checks if `checkDate` falls within $[startDate, endDate]$. | Returns `TRUE` / `FALSE`. |
| `EX_RangeOverlaps(start1, end1, start2, end2)` | Checks whether two date ranges overlap. | Returns `TRUE` / `FALSE`. |

---

## 🏷️ Formatting Options

Custom format patterns can be passed to `EX_Format`, `EX_NepaliDate`, `EX_Today`, and other formatting functions:

| Specifier | Description | Example Output |
| :---: | :--- | :--- |
| `yyyy` | 4-digit Nepali Year | `2080` |
| `yy` | 2-digit Nepali Year | `80` |
| `MM` | 2-digit Nepali Month with leading zero | `01` |
| `M` | 1 or 2-digit Nepali Month without leading zero | `1` |
| `dd` | 2-digit Nepali Day with leading zero | `05` |
| `d` | 1 or 2-digit Nepali Day without leading zero | `5` |
| `MMMM` | Full English name of the Nepali month | `Baishakh` |
| `MMM` | Abbreviated name of the Nepali month | `Bai` |
| `dddd` | Full weekday name | `Friday` |
| `ddd` | Abbreviated weekday name | `Fri` |

---

## 🛠️ Building from Source

To compile **ExNepDate** locally:

1. **Clone the repository**:
   ```bash
   git clone https://github.com/bishwokarmalaxman/ExNepDate.git
   cd ExNepDate
   ```

2. **Build with .NET CLI**:
   ```bash
   dotnet build -c Release
   ```

3. The compiled `.xll` files will be generated under `ExNepDate/bin/Release/net10.0-windows/`.

---

## 📄 License & Credits

- **License**: Released under the [MIT License](LICENSE).
- **Core Library**: Powered by [.NET NepDate]([https://github.com/RajuPrasai/NepDate]) by NepDate contributors.
- **Excel Integration**: Built using [Excel-DNA](https://github.com/Excel-DNA/ExcelDna) (Licensed under the Excel-DNA License).

---

<p align="center">
  Made with ❤️ for Excel power-users, accountants, and businesses working with Bikram Sambat in Nepal.
</p>
