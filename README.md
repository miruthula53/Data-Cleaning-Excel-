# Data Cleaning in Excel: HR Dataset

Cleaning a messy HR dataset in Excel, with the raw file and the cleaned file kept side by side so the before and after can be compared.

## Objective
Find the data quality problems in a raw HR dataset and fix them, so the data is ready for analysis and dashboards.

## Tools Used
Microsoft Excel: `TRIM`, `PROPER`, `UPPER`, `VALUE`, `SUBSTITUTE`, `RIGHT`, `IF`, `ISNUMBER`, `MEDIAN`, `COUNTBLANK`, Find & Replace, Text to Columns, Filters, Remove Duplicates.

## Columns
Emp ID, Name, Age, Sex, City, Department, Join Date, Salary, Phone, Email

## Issues Found and How I Fixed Them

| Column | Problem | Fix |
|---|---|---|
| Age | Blanks, "N/A", "-", numbers stored as text, outliers (250, -5, 0, 999) | `COUNTBLANK`, `VALUE`, filtered outliers, blanks kept as blank |
| Sex | M / F / Male / male / FEMALE, trailing spaces | `TRIM` + Find & Replace, standardised to Male / Female |
| City | Mixed case, extra spaces, typos (Chenai, Eroad, Selam), Bengaluru vs Bangalore, blanks, "NA" | `PROPER(TRIM())`, Find & Replace, blanks set to "Unknown", Tiruchirappalli changed to Trichy |
| Department | hr / HR / Human Resources, I.T., Fin, Mktg, Ops | `TRIM`, `UPPER`, Find & Replace |
| Join Date | Real dates mixed with text dates in 4 formats, "Unknown", blanks | Text to Columns, `DATEVALUE`, date number format |
| Salary | Text numbers, "Rs.45,000", ₹ sign, "-", "null", outliers (9999999, 5) | `SUBSTITUTE` to remove symbols, `VALUE`, filtered outliers |
| Phone | +91 prefix, spaces, dashes, leading 0 | `SUBSTITUTE`, `RIGHT(...,10)`, `TRIM` |
| Email | Blanks, "N/A", invalid emails | Filter, `ISNUMBER(SEARCH("@",...))` |
| Name | Extra spaces, UPPER/lower case, double spaces | `PROPER(TRIM())` |
| Whole sheet | About 11 duplicate rows, some with different name case | Remove Duplicates after text cleaning |

## Steps
1. Copied the raw sheet to a new sheet so the original stays untouched.
2. Checked each column for blanks, wrong data types and inconsistent values.
3. Cleaned text columns first (`TRIM`, `PROPER`), then standardised categories with Find & Replace.
4. Converted text numbers and dates into real numbers and dates.
5. Handled outliers by filtering and reviewing them.
6. Removed duplicate rows after text cleaning, so case differences didn't hide duplicates.
7. Validated the result: counted rows, checked unique values in each category column, and reviewed blanks.

## Result
| | Before | After |
|---|---|---|
| Duplicate rows | About 11 | 0 |
| Gender values | Many variants | 2 (Male, Female) |
| Department values | Many variants | 6 standard values |
| Salary and Age | Text and outliers | Numeric, valid range |
| Join Date | Mixed formats | One date format |
| Final records | | 290 |

## Decisions
- Remaining blanks (22 ages, 2 salaries, 8 join dates) are left blank and not filled with guessed values, so averages stay honest.
- Blank cities and Emails are labelled "Unknown" and Phone Numbers are labelled "Not Provided".
- Personal fields (Name, Phone, Email) should be removed before sharing data publicly.

## Author
**Miruthula**
Final-year ECE student | Aspiring Data Analyst
Skills: Python, SQL, MySQL, Excel, Power BI
GitHub: [miruthula53](https://github.com/miruthula53)
