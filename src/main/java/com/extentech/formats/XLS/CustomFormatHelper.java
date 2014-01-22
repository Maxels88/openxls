/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 * 
 * This file is part of OpenXLS.
 * 
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 * 
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
/**
 * CustomFormatHelper.java
 *
 *
 * Feb 26, 2010
 *
 *
 */
package com.extentech.formats.XLS;

/**
 * This class is used by FormatConstants to provide for handling of custom formats
 * <p/>
 * <p/>
 * From: http://support.microsoft.com/kb/264372
 * <p/>
 * Custom Number Formats
 * If one of the built-in number formats does not display the data in the format that you require, you can create your own custom number format. You can create these custom number formats by modifying the built-in formats or by combining the formatting symbols into your own combination.
 * <p/>
 * Before you create your own custom number format, you need to be aware of a few simple rules governing the syntax for number formats:
 * Each format that you create can have up to three sections for numbers and a fourth section for text.
 * <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>
 * <p/>
 * <p/>
 * The first section is the format for positive numbers, the second for negative numbers, and the third for zero values.
 * These sections are separated by semicolons.
 * If you have only one section, all numbers (positive, negative, and zero) are formatted with that format.
 * You can prevent any of the number types (positive, negative, zero) from being displayed by not typing symbols in the corresponding section. For example, the following number format prevents any negative or zero values from being displayed:
 * 0.00;;
 * To set the color for any section in the custom format, type the name of the color in brackets in the section. For example, the following number format formats positive numbers blue and negative numbers red:
 * [BLUE]#,##0;[RED]#,##0
 * Instead of the default positive, negative and zero sections in the format, you can specify custom criteria that must be met for each section. The conditional statements that you specify must be contained within brackets. For example, the following number format formats all numbers greater than 100 as green, all numbers less than or equal to -100 as yellow, and all other numbers as cyan:
 * [>100][GREEN]#,##0;[<=-100][YELLOW]#,##0;[CYAN]#,##0
 * For each part of the format, type symbols that represent how you want the number to look. See the table below for details on all the available symbols.
 * To create a custom number format, click Custom in the Category list on the Number tab in the Format Cells dialog box. Then, type your custom number format in the Type box.
 * <p/>
 * The following table outlines the different symbols available for use in custom number formats.
 * Format Symbol      Description/result
 * ------------------------------------------------------------------------
 * <p/>
 * 0                  Digit placeholder. For example, if you type 8.9 and
 * you want it to display as 8.90, then use the
 * format #.00
 * <p/>
 * #                  Digit placeholder. Follows the same rules as the 0
 * symbol except Excel does not display extra zeros
 * when the number you type has fewer digits on either
 * side of the decimal than there are # symbols in the
 * format. For example, if the custom format is #.## and
 * you type 8.9 in the cell, the number 8.9 is
 * displayed.
 * <p/>
 * ?                  Digit placeholder. Follows the same rules as the 0
 * symbol except Excel places a space for insignificant
 * zeros on either side of the decimal point so that
 * decimal points are aligned in the column. For
 * example, the custom format 0.0? aligns the decimal
 * points for the numbers 8.9 and 88.99 in a column.
 * <p/>
 * . (period)         Decimal point.
 * <p/>
 * %                  Percentage. If you enter a number between 0 and 1,
 * and you use the custom format 0%, Excel multiplies
 * the number by 100 and adds the % symbol in the cell.
 * <p/>
 * , (comma)          Thousands separator. Excel separates thousands by
 * commas if the format contains a comma surrounded by
 * '#'s or '0's. A comma following a placeholder
 * scales the number by a thousand. For example, if the
 * format is #.0,, and you type 12,200,000 in the cell,
 * the number 12.2 is displayed.
 * <p/>
 * E- E+ e- e+        Scientific format. Excel displays a number to the
 * right of the "E" symbol that corresponds to the
 * number of places the decimal point was moved. For
 * example, if the format is 0.00E+00 and you type
 * 12,200,000 in the cell, the number 1.22E+07 is
 * displayed. If you change the number format to #0.0E+0
 * the number 12.2E+6 is displayed.
 * <p/>
 * $-+/():space       Displays the symbol. If you want to display a
 * character that is different than one of these
 * symbols, precede the character with a backslash (\)
 * or enclose the character in quotation marks (" ").
 * For example, if the number format is (000) and you
 * type 12 in the cell, the number (012) is displayed.
 * <p/>
 * \                  Display the next character in the format. Excel does
 * not display the backslash. For example, if the number
 * format is 0\! and you type 3 in the cell, the value
 * 3! is displayed.
 * <p/>
 * Repeat the next character in the format enough times
 * to fill the column to its current width. You cannot
 * have more than one asterisk in one section of the
 * format. For example, if the number format is 0*x and
 * you type 3 in the cell, the value 3xxxxxx is
 * displayed. Note, the number of "x" characters
 * displayed in the cell vary based on the width of the
 * column.
 * <p/>
 * _ (underline)      Skip the width of the next character. This is useful
 * for lining up negative and positive values in
 * different cells of the same column. For example, the
 * number format _(0.0_);(0.0) align the numbers
 * 2.3 and -4.5 in the column even though the negative
 * number has parentheses around it.
 * <p/>
 * "text"             Display whatever text is inside the quotation marks.
 * For example, the format 0.00 "dollars" displays
 * "1.23 dollars" (without quotation marks) when you
 * type 1.23 into the cell.
 *
 * @ Text placeholder. If there is text typed in the
 * cell, the text from the cell is placed in the format
 * where the @ symbol appears. For example, if the
 * number format is "Bob "@" Smith" (including
 * quotation marks) and you type "John" (without
 * quotation marks) in the cell, the value
 * "Bob John Smith" (without quotation marks) is
 * displayed.
 * <p/>
 * DATE FORMATS
 * <p/>
 * m                  Display the month as a number without a leading zero.
 * <p/>
 * mm                 Display the month as a number with a leading zero
 * when appropriate.
 * <p/>
 * mmm                Display the month as an abbreviation (Jan-Dec).
 * <p/>
 * mmmm               Display the month as a full name (January-December).
 * <p/>
 * d                  Display the day as a number without a leading zero.
 * <p/>
 * dd                 Display the day as a number with a leading zero
 * when appropriate.
 * <p/>
 * ddd                Display the day as an abbreviation (Sun-Sat).
 * <p/>
 * dddd               Display the day as a full name (Sunday-Saturday).
 * <p/>
 * yy                 Display the year as a two-digit number.
 * <p/>
 * yyyy               Display the year as a four-digit number.
 * <p/>
 * TIME FORMATS
 * <p/>
 * h                  Display the hour as a number without a leading zero.
 * <p/>
 * [h]                Elapsed time, in hours. If you are working with a
 * formula that returns a time where the number of hours
 * exceeds 24, use a number format similar to
 * [h]:mm:ss.
 * <p/>
 * hh                 Display the hour as a number with a leading zero when
 * appropriate. If the format contains AM or PM, then
 * the hour is based on the 12-hour clock. Otherwise,
 * the hour is based on the 24-hour clock.
 * <p/>
 * m                  Display the minute as a number without a leading
 * zero.
 * <p/>
 * [m]                Elapsed time, in minutes. If you are working with a
 * formula that returns a time where the number of
 * minutes exceeds 60, use a number format similar to
 * [mm]:ss.
 * <p/>
 * mm                 Display the minute as a number with a leading zero
 * when appropriate. The m or mm must appear immediately
 * after the h or hh symbol, or Excel displays the
 * month rather than the minute.
 * <p/>
 * s                  Display the second as a number without a leading
 * zero.
 * <p/>
 * [s]                Elapsed time, in seconds. If you are working with a
 * formula that returns a time where the number of
 * seconds exceeds 60, use a number format similar to
 * [ss].
 * <p/>
 * ss                 Display the second as a number with a leading zero
 * when appropriate.
 * <p/>
 * NOTE: If you want to display fractions of a second,
 * use a number format similar to h:mm:ss.00.
 * <p/>
 * AM/PM              Display the hour using a 12-hour clock. Excel
 * am/pm              displays AM, am, A, or a for times from midnight
 * A/P                until noon, and PM, pm, P, or p for times from noon
 * a/p                until midnight.
 */
public class CustomFormatHelper
{

}
