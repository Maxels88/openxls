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
package com.extentech.formats.XLS.formulas;

//import java.text.SimpleDateFormat;

import com.extentech.ExtenXLS.DateConverter;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.TimeZone;

/**
 * DateTimeCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 * <p/>
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */

public class DateTimeCalculator
{
	/**
	 * utliity that takes either a PtgStr or a Number and converts it into a calendar
	 * for use in below functions
	 *
	 * @return
	 */
	private static GregorianCalendar getDateFromPtg( Ptg op )
	{
		Object o;
		if( op instanceof PtgStr )
		{
			o = calcDateValue( new Ptg[]{ op } ).getValue();
		}
		else if( op instanceof PtgRef )
		{
			o = op.getValue();
			if( o instanceof String )
			{
				o = calcDateValue( new Ptg[]{ new PtgStr( o.toString() ) } ).getValue();
			}
		}
		else if( op instanceof PtgName )
		{
			o = ((PtgName) op).getValue();    //getComponents()[0];
			o = op.getValue();
			if( o instanceof String )
			{
				o = calcDateValue( new Ptg[]{ new PtgStr( o.toString() ) } ).getValue();
			}
		}
		else
		{
			o = op.getValue();
		}

		return (GregorianCalendar) DateConverter.getCalendarFromNumber( o );
	}

	/**
	 * DATE
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcDate( Ptg[] operands )
	{
		long[] alloperands = PtgCalculator.getLongValueArray( operands );
		if( alloperands.length != 3 )
		{
			return PtgCalculator.getError();
		}
		int year = (int) alloperands[0];
		int month = (int) alloperands[1];
		month = month - 1;
		int day = (int) alloperands[2];
		GregorianCalendar c = new GregorianCalendar( year, month, day );
		double date = DateConverter.getXLSDateVal( c );
		int i = (int) date;
		PtgInt pi = new PtgInt( i );
		return pi;
	}

	/**
	 * DATEVALUE
	 * Returns the serial number of the date represented by date_text. Use DATEVALUE to convert a date represented by text to a serial number.
	 * <p/>
	 * Syntax
	 * DATEVALUE(date_text)
	 * <p/>
	 * Date_text   is text that represents a date in a Microsoft Excel date format. For example, "1/30/2008" or "30-Jan-2008" are text strings
	 * within quotation marks that represent dates. Using the default date system in Excel for Windows,
	 * date_text must represent a date from January 1, 1900, to December 31, 9999. Using the default date system in Excel for the Macintosh,
	 * date_text must represent a date from January 1, 1904, to December 31, 9999. DATEVALUE returns the #VALUE! error value if date_text is out of this range.
	 * <p/>
	 * If the year portion of date_text is omitted, DATEVALUE uses the current year from your computer's built-in clock. Time information in date_text is ignored.
	 * <p/>
	 * Remarks
	 * <p/>
	 * Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Excel for the Macintosh uses a different date system as its default.
	 * Most functions automatically convert date values to serial numbers.
	 *
	 * @param operands Ptg[]
	 * @return Ptg
	 */

	protected static Ptg calcDateValue( Ptg[] operands )
	{
		// TODO: there may be formats that need to be input 
		if( operands == null || operands[0].getString() == null )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}

		String dateString = operands[0].getString();
		Double d = DateConverter.calcDateValue( dateString );
		if( d == null )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		return new PtgNumber( d );

	}

	/**
	 * DAY
	 * Return the day of the month
	 */
	protected static Ptg calcDay( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		try
		{
			Object o = operands[0].getValue();
			GregorianCalendar c = (GregorianCalendar) DateConverter.getCalendarFromNumber( o );
			int retdate = c.get( Calendar.DAY_OF_MONTH );
			PtgInt pint = new PtgInt( retdate );
			return pint;
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * DAYS360
	 * Calculate the difference between 2 dates based on a 360
	 * day year, 12 mos, 30days each.
	 * <p/>
	 * first date is lower than second, otherwise a negative
	 * number is returned
	 * <p/>
	 * Seems pretty dumb to me, but what do I know?
	 */
	protected static Ptg calcDays360( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		try
		{
			Object o1 = operands[0].getValue();
			Object o2 = operands[1].getValue();
			GregorianCalendar dt1 = (GregorianCalendar) DateConverter.getCalendarFromNumber( o1 );
			GregorianCalendar dt2 = (GregorianCalendar) DateConverter.getCalendarFromNumber( o2 );
			int yr1 = dt1.get( Calendar.YEAR );
			int yr2 = dt2.get( Calendar.YEAR );
			int diff = yr2 - yr1;
			diff *= 360; // turn years to days.
			int mo1 = dt1.get( Calendar.MONTH );
			int mo2 = dt2.get( Calendar.MONTH );
			int mos = 0;
			if( mo2 > mo1 )
			{
				mos = mo2 - mo1;
			}
			else
			{
				diff -= 360;
				while( mo2 != mo1 )
				{
					mos++;
					mo1++;
					if( mo1 == 12 )
					{
						mo1 = 0;
					}
				}
			}
			diff += mos * 30;
			int dy1 = dt1.get( Calendar.DAY_OF_MONTH );
			int dy2 = dt2.get( Calendar.DAY_OF_MONTH );
			if( dy2 > dy1 )
			{
				diff += dy2 - dy1;
			}
			else
			{
				diff -= 30;
				while( dy2 != dy1 )
				{
					diff++;
					dy1++;
					if( dy1 == 30 )
					{
						dy1 = 0;
					}
				}
			}
			PtgInt pint = new PtgInt( diff );
			return pint;
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * EDATE
	 * Returns the serial number that represents the date that is the indicated number of months
	 * before or after a specified date (the start_date).
	 * Use EDATE to calculate maturity dates or due dates that fall on the same day of the month as the date of issue.
	 * <p/>
	 * EDATE(start_date,months)
	 * <p/>
	 * Start_date   is a date that represents the start date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
	 * <p/>
	 * Months   is the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
	 * <p/>
	 * If start_date is not a valid date, EDATE returns the #VALUE! error value.
	 * If months is not an integer, it is truncated.
	 */
	protected static Ptg calcEdate( Ptg[] operands )
	{
		try
		{
			GregorianCalendar startDate = getDateFromPtg( operands[0] );
			int inc = (int) PtgCalculator.getLongValue( operands[1] );
			int mm = startDate.get( Calendar.MONTH ) + inc;
			int y = startDate.get( Calendar.YEAR );
			int d = startDate.get( Calendar.DAY_OF_MONTH );
			if( mm < 0 )
			{
				mm += 12;    // 0-based
				y--;
			}
			else if( mm > 11 )
			{
				mm -= 12;
				y++;
			}
			GregorianCalendar resultDate;
			resultDate = new GregorianCalendar( y, mm, d );
			double retdate = DateConverter.getXLSDateVal( resultDate );
			int i = (int) retdate;
			return new PtgInt( i );
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * EOMONTH
	 * Returns the serial number for the last day of the month that is the indicated number of months
	 * before or after start_date. Use EOMONTH to calculate maturity dates or due dates that fall on
	 * the last day of the month.
	 * <p/>
	 * EOMONTH(start_date,months)
	 * <p/>
	 * Start_date   is a date that represents the starting date. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
	 * <p/>
	 * Months   is the number of months before or after start_date. A positive value for months yields a future date; a negative value yields a past date.
	 * <p/>
	 * If months is not an integer, it is truncated.
	 * <p/>
	 * If start_date is not a valid date, EOMONTH returns the #NUM! error value.
	 * If start_date plus months yields an invalid date, EOMONTH returns the #NUM! error value.
	 */
	protected static Ptg calcEOMonth( Ptg[] operands )
	{
		try
		{
			GregorianCalendar startDate = getDateFromPtg( operands[0] );
			int inc = operands[1].getIntVal();
			int mm = startDate.get( Calendar.MONTH ) + inc;
			int y = startDate.get( Calendar.YEAR );
			int d = startDate.get( Calendar.DAY_OF_MONTH );
			if( mm < 0 )
			{
				mm += 12;    // 0-based
				y--;
			}
			else if( mm > 11 )
			{
				mm -= 12;
				y++;
			}
			if( mm == 3 || mm == 5 || mm == 8 || mm == 10 )    // 0-based
			{
				d = 30;
			}
			else if( mm == 1 )
			{// february
				if( y % 4 == 0 )
				{
					d = 29;
				}
				else
				{
					d = 28;
				}
			}
			else
			{
				d = 31;
			}
			GregorianCalendar resultDate;
			resultDate = new GregorianCalendar( y, mm, d );
			double retdate = DateConverter.getXLSDateVal( resultDate );
			int i = (int) retdate;
			return new PtgInt( i );
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_NUM );
	}

	/**
	 * HOUR
	 * Converts a serial number to an hour
	 */
	protected static Ptg calcHour( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		GregorianCalendar dt = getDateFromPtg( operands[0] );
		int retdate = dt.get( Calendar.HOUR );
		PtgInt pint = new PtgInt( retdate );
		return pint;

	}

	/**
	 * MINUTE
	 * Converts a serial number to a minute
	 */
	protected static Ptg calcMinute( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		try
		{
			Object o = operands[0].getValue();
			GregorianCalendar dt = (GregorianCalendar) DateConverter.getCalendarFromNumber( o );
			int retdate = dt.get( Calendar.MINUTE );
			PtgInt pint = new PtgInt( retdate );
			return pint;
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * MONTH
	 * Converts a serial number to a month
	 */
	protected static Ptg calcMonth( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		try
		{
			GregorianCalendar dt = getDateFromPtg( operands[0] );
			int retdate = dt.get( Calendar.MONTH );
			retdate++; //month is ordinal
			PtgInt pint = new PtgInt( retdate );
			return pint;
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/**
	 * NETWORKDAYS
	 * Returns the number of whole working days between start_date and end_date. Working days exclude weekends and any dates identified in holidays.
	 * Use NETWORKDAYS to calculate employee benefits that accrue based on the number of days worked during a specific term.
	 * <p/>
	 * NETWORKDAYS(start_date,end_date,holidays)
	 * <p/>
	 * Start_date   is a date that represents the start date.
	 * End_date   is a date that represents the end date.
	 * Holidays   is an optional range of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contains the dates or an array constant of the serial numbers that represent the dates.
	 * <p/>
	 * Remarks
	 * If any argument is not a valid date, NETWORKDAYS returns the #VALUE! error value.
	 */
	protected static Ptg calcNetWorkdays( Ptg[] operands )
	{
		try
		{
			ArrayList holidays = new ArrayList();
			GregorianCalendar startDate = getDateFromPtg( operands[0] );
			GregorianCalendar endDate = getDateFromPtg( operands[1] );
			if( operands.length > 2 && operands[2] != null )
			{
				if( operands[2] instanceof PtgRef )
				{
					Ptg[] dts = ((PtgRef) operands[2]).getComponents();
					for( int i = 0; i < dts.length; i++ )
					{
						holidays.add( getDateFromPtg( dts[i] ) );
					}
				}
				else  // assume it's a string or a number rep of a date
				{
					holidays.add( getDateFromPtg( operands[2] ) );
				}
			}
			int count = 0;
			boolean countUp = endDate.after( startDate );
			while( !startDate.equals( endDate ) )
			{
				int d = startDate.get( Calendar.DAY_OF_WEEK );
				if( d != Calendar.SATURDAY && d != Calendar.SUNDAY )
				{
					boolean OKtoIncrement = true;
					// check if on a holidays
					if( holidays.size() > 0 )
					{
						for( int i = 0; i < holidays.size(); i++ )
						{
							if( startDate.equals( ((Calendar) holidays.get( i )) ) )
							{
								OKtoIncrement = false;
								break;
							}
						}
					}
					if( OKtoIncrement )
					{
						count++;
					}
				}
				if( countUp )
				{
					startDate.add( Calendar.DAY_OF_MONTH, 1 );
				}
				else
				{
					startDate.add( Calendar.DAY_OF_MONTH, -1 );
				}
			}
			return new PtgInt( count );
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * NOW
	 * Returns the serial number of the current date and time
	 */
	protected static Ptg calcNow( Ptg[] operands )
	{
		GregorianCalendar gc = new GregorianCalendar( TimeZone.getDefault() );//java.sql.Date dt = new java.sql.Date();
		// System.out.println(dt.toGMTString());
		double retdate = DateConverter.getXLSDateVal( gc );
		return new PtgNumber( retdate );
	}

	/**
	 * SECOND
	 * Converts a serial number to a second
	 */
	protected static Ptg calcSecond( Ptg[] operands )
	{
		GregorianCalendar dt = getDateFromPtg( operands[0] );
		int retdate = dt.get( Calendar.SECOND );
		PtgInt pint = new PtgInt( retdate );
		return pint;

	}

	/**
	 * TIME
	 * Returns the serial number of a particular time
	 * takes 3 arguments, hour, minute, second;
	 */
	protected static Ptg calcTime( Ptg[] operands )
	{
		Ptg o;
		if( operands[0] instanceof PtgStr )
		{
			o = calcDateValue( new Ptg[]{ operands[0] } );
		}
		else
		{
			o = operands[0];
		}
		int hour = o.getIntVal();
		int minute = operands[1].getIntVal();
		int second = operands[1].getIntVal();
		GregorianCalendar g = new GregorianCalendar( 2000, 1, 1, hour, minute, second );
		GregorianCalendar g2 = new GregorianCalendar( 2000, 1, 1, 0, 0, 0 );
		double dub = DateConverter.getXLSDateVal( g );
		double dub2 = DateConverter.getXLSDateVal( g2 );
		dub -= dub2;
		return new PtgNumber( dub );
	}

	/**
	 * TIMEVALUE
	 * Converts a time in the form of text to a serial number
	 * Returns the decimal number of the time represented by a text string.
	 * The decimal number is a value ranging from 0 (zero) to 0.99999999, representing the times from 0:00:00 (12:00:00 AM) to 23:59:59 (11:59:59 P.M.).
	 * TIMEVALUE(time_text)
	 * Time_text   is a text string that represents a time in any one of the Microsoft Excel time formats; for example, "6:45 PM" and "18:45" text strings within quotation marks that represent time.
	 */
	protected static Ptg calcTimevalue( Ptg[] operands )
	{
		double result = 0;
		try
		{
			GregorianCalendar d = getDateFromPtg( operands[0] );
			int h = d.get( Calendar.HOUR_OF_DAY );
			int m = d.get( Calendar.MINUTE );
			int s = d.get( Calendar.SECOND );
			double t = h + (m / 60.0) + (s / (60 * 60.0));
			result = (t / 24.0);
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		return new PtgNumber( result );
	}

	/**
	 * TODAY
	 * Returns the serial number of today's date
	 */
	protected static Ptg calcToday( Ptg[] operands )
	{
		java.sql.Date dt = new java.sql.Date( System.currentTimeMillis() );
		double retdate = DateConverter.getXLSDateVal( dt );
		int i = (int) retdate;
		return new PtgInt( i );
	}

	/**
	 * WEEKDAY  Converts a serial number to a day of the week
	 */
	protected static Ptg calcWeekday( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		GregorianCalendar dt = getDateFromPtg( operands[0] );
		int retdate = dt.get( Calendar.DAY_OF_WEEK );
		PtgInt pint = new PtgInt( retdate );
		return pint;
	}

	/**
	 * WEEKNUM
	 * Returns a number that indicates where the week falls numerically within a year.
	 * <p/>
	 * WEEKNUM(serial_num,return_type)
	 * <p/>
	 * Serial_num   is a date within the week. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
	 * <p/>
	 * Return_type   is a number that determines on which day the week begins. The default is 1.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcWeeknum( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int returnType = 1;
		GregorianCalendar dt = getDateFromPtg( operands[0] );
		if( operands[1] != null )
		{
			returnType = operands[1].getIntVal();
		}
		returnType -= 1;    // 1 is default, 2 =start on monday
		int retdate = dt.get( Calendar.WEEK_OF_YEAR ) - returnType;
		PtgInt pint = new PtgInt( retdate );
		return pint;
	}

	/**
	 * WORKDAY
	 * Returns a number that represents a date that is the indicated number of working days before or after a date (the starting date).
	 * Working days exclude weekends and any dates identified as holidays. Use WORKDAY to exclude weekends or holidays when you calculate
	 * invoice due dates, expected delivery times, or the number of days of work performed.
	 * <p/>
	 * WORKDAY(start_date,days,holidays)
	 * <p/>
	 * Important   Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
	 * Start_date   is a date that represents the start date.
	 * Days   is the number of nonweekend and nonholiday days before or after start_date. A positive value for days yields a future date; a negative value yields a past date.
	 * Holidays   is an optional list of one or more dates to exclude from the working calendar, such as state and federal holidays and floating holidays. The list can be either a range of cells that contain the dates or an array constant of the serial numbers that represent the dates.
	 * <p/>
	 * Remarks
	 * <p/>
	 * If any argument is not a valid date, WORKDAY returns the #VALUE! error value.
	 * If start_date plus days yields an invalid date, WORKDAY returns the #NUM! error value.
	 * If days is not an integer, it is truncated.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcWorkday( Ptg[] operands )
	{
		int days;
		ArrayList holidays = new ArrayList();
		try
		{
			GregorianCalendar dt = getDateFromPtg( operands[0] );
			days = operands[1].getIntVal();
			if( operands.length > 2 && operands[2] != null )
			{    // holidays
				if( operands[2] instanceof PtgRef )
				{
					Ptg[] dts = ((PtgRef) operands[2]).getComponents();
					for( int i = 0; i < dts.length; i++ )
					{
						holidays.add( getDateFromPtg( dts[i] ) );
					}
				}
				else  // assume it's a string or a number rep of a date
				{
					holidays.add( getDateFromPtg( operands[2] ) );
				}
			}
			for( int absDays = Math.abs( days ); absDays > 0; )
			{
				int d = dt.get( Calendar.DAY_OF_WEEK );
				if( d != Calendar.SATURDAY && d != Calendar.SUNDAY )
				{
					boolean OKtoIncrement = true;
					// check if on a holidays
					if( holidays.size() > 0 )
					{
						for( int i = 0; i < holidays.size(); i++ )
						{
							if( dt.equals( ((Calendar) holidays.get( i )) ) )
							{
								OKtoIncrement = false;
								break;
							}
						}
					}
					if( OKtoIncrement )
					{
						absDays--;
					}
				}
				if( days > 0 )
				{
					dt.add( Calendar.DAY_OF_MONTH, 1 );
				}
				else
				{
					dt.add( Calendar.DAY_OF_MONTH, -1 );
				}
			}
			return new PtgNumber( DateConverter.getXLSDateVal( dt ) );

		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * YEAR
	 * Converts a serial number to a year
	 */
	protected static Ptg calcYear( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		try
		{
			GregorianCalendar dt = getDateFromPtg( operands[0] );
			int retdate = dt.get( Calendar.YEAR );
			PtgInt pint = new PtgInt( retdate );
			return pint;
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/**
	 * YEARFRAC function
	 * Calculates the fraction of the year represented by the number of whole days between two dates (the start_date and the end_date). Use the YEARFRAC worksheet function to identify the proportion of a whole year's benefits or obligations to assign to a specific term.
	 * YEARFRAC(start_date,end_date,basis)
	 * Important  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.
	 * Start_date   is a date that represents the start date.
	 * End_date   is a date that represents the end date.
	 * Basis   is the type of day count basis to use.
	 * 0 or omitted US (NASD) 30/360
	 * 1 Actual/actual
	 * 2 Actual/360
	 * 3 Actual/365
	 * 4 European 30/360
	 */
	protected static Ptg calcYearFrac( Ptg[] operands )
	{
		long startDate, endDate;
		try
		{
			GregorianCalendar d = getDateFromPtg( operands[0] );
			startDate = (new Double( DateConverter.getXLSDateVal( d ) )).longValue();
			d = getDateFromPtg( operands[1] );
			endDate = (new Double( DateConverter.getXLSDateVal( d ) )).longValue();
		}
		catch( Exception e )
		{
			//If start_date or end_date are not valid dates, YEARFRAC returns the #VALUE! error value.
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int basis = 0;
		if( operands.length > 2 )
		{
			basis = operands[2].getIntVal();
		}
		//If basis < 0 or if basis > 4, YEARFRAC returns the #NUM! error value.
		if( basis < 0 || basis > 4 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double yf = FinancialCalculator.yearFrac( basis, startDate, endDate );
		if( yf < 0 )
		{
			yf *= -1;    // =# days between dates, no negatives
		}
		return new PtgNumber( yf );
	}
}	