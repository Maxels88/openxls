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

import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/*
    EngineeringCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
public class EngineeringCalculator
{
	private static final Logger log = LoggerFactory.getLogger( EngineeringCalculator.class );
/*
	BESSELI
	 Returns the modified Bessel function In(x)
 
	BESSELJ
	 Returns the Bessel function Jn(x)
 
	BESSELK
	 Returns the modified Bessel function Kn(x)
 
	BESSELY
	 Returns the Bessel function Yn(x)
*/

	/**
	 * BIN2DEC
	 * Converts a binary number to decimal
	 */
	protected static Ptg calcBin2Dec( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "BIN2DEC" );
		String bString = operands[0].getString().trim();
		// 10 bits at most
		if( bString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// must det. manually if binary string is negative because parseInt does not
		// handle two's complement input!!!
		boolean bIsNegative = ((bString.length() == 10) && bString.substring( 0, 1 ).equalsIgnoreCase( "1" ));

		int dec = 0;
		try
		{
			dec = Integer.parseInt( bString, 2 );
			if( bIsNegative )
			{
				dec -= 1024;    // 2^10 (= signed 11 bits)
			}
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		PtgNumber pnum = new PtgNumber( dec );
		log.debug( "Result from BIN2DEC= " + pnum.getVal() );
		return pnum;
	}

	/**
	 * BIN2HEX
	 * Converts a binary number to hexadecimal
	 */
	protected static Ptg calcBin2Hex( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "BIN2HEX" );
		String bString = operands[0].getString().trim();
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		// 10 bits at most
		if( bString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// must det. manually if binary string is negative because parseInt does not
		// handle two's complement input!!!
		boolean bIsNegative = ((bString.length() == 10) && bString.substring( 0, 1 ).equalsIgnoreCase( "1" ));

		long dec;
		String hString;
		try
		{
			dec = Long.parseLong( bString, 2 );
			if( bIsNegative )
			{
				dec -= 1024;        // 2^10 (= signed 11 bits)
			}
			hString = Long.toHexString( dec ).toUpperCase();
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			hString = hString.substring( Math.max( hString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( hString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			hString = ("0000000000" + hString);    // maximum= 10 bits
			hString = hString.substring( hString.length() - places );
		}

		PtgStr pstr = new PtgStr( hString );
		log.debug( "Result from BIN2HEX= " + pstr.getString() );
		return pstr;
	}

	/**
	 * BIN2OCT
	 * Converts a binary number to octal
	 */
	protected static Ptg calcBin2Oct( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "Bin2Oct" );
		String bString = operands[0].getString().trim();
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		// 10 bits at most
		if( bString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// must det. manually if binary string is negative because parseInt does not
		// handle two's complement input!!!
		boolean bIsNegative = ((bString.length() == 10) && bString.substring( 0, 1 ).equalsIgnoreCase( "1" ));

		int dec;
		String oString;
		try
		{
			dec = Integer.parseInt( bString, 2 );
			if( bIsNegative )
			{
				dec -= 1024;        // 2^10 (= signed 11 bits)
			}
			oString = Integer.toOctalString( dec );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			oString = oString.substring( Math.max( oString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( oString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			oString = ("0000000000" + oString);    // maximum= 10 bits
			oString = oString.substring( oString.length() - places );
		}
		PtgStr pstr = new PtgStr( oString );
		log.debug( "Result from BIN2OCT= " + pstr.getString() );
		return pstr;
	}

	/**
	 * COMPLEX
	 * Converts real and imaginary coefficients into a complex number
	 */
	protected static Ptg calcComplex( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "Complex" );
		int real = operands[0].getIntVal();
		int imaginary = operands[1].getIntVal();
		String suffix = "i";
		if( operands.length > 2 )
		{
			suffix = operands[2].getString().trim();
			if( !(suffix.equals( "i" ) || suffix.equals( "j" )) )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
		}
		String complexString = "";

		// result:  real + imaginary suffix
		// 			real - imaginary suffix
		//			real				(if imaginary==0)
		//			real + suffix 		(if imaginary==1)
		//			imaginary suffix	(if (real==0)
		if( real != 0 )
		{
			complexString = String.valueOf( real );
			if( imaginary > 0 )
			{
				complexString += " + ";
			}
		}

		if( imaginary != 0 )
		{
			complexString += ((Math.abs( imaginary ) != 1) ? String.valueOf( imaginary ) : "") + suffix;
		}

		if( complexString.equals( "" ) )
		{
			complexString = "0";
		}

		PtgStr pstr = new PtgStr( complexString );
		log.debug( "Result from COMPLEX= " + pstr.getString() );
		return pstr;
	}

	/**
	 * CONVERT
	 * Converts a number from one measurement system to another
	 * Weight and mass 		From_unit or to_unit
	 * Gram 					"g"
	 * Slug 					"sg"
	 * Pound mass (avoirdupois) "lbm"
	 * U (atomic mass unit) 	"u"
	 * Ounce mass (avoirdupois) "ozm"
	 * <p/>
	 * Distance 				From_unit or to_unit
	 * Meter 					"m"
	 * Statute mile 			"mi"
	 * Nautical mile 			"Nmi"
	 * Inch 					"in"
	 * Foot 					"ft"
	 * Yard 					"yd"
	 * Angstrom 				"ang"
	 * Pica (1/72 in.) 		"Pica"
	 * <p/>
	 * Time 					From_unit or to_unit
	 * Year 					"yr"
	 * Day 					"day"
	 * Hour 					"hr"
	 * Minute 					"mn"
	 * Second 					"sec"
	 * <p/>
	 * Pressure 				From_unit or to_unit
	 * Pascal 					"Pa"
	 * Atmosphere 				"atm"
	 * mm of Mercury 			"mmHg"
	 * <p/>
	 * Force 					From_unit or to_unit
	 * Newton 					"N"
	 * Dyne 					"dyn"
	 * Pound force 			"lbf"
	 * <p/>
	 * Energy 					From_unit or to_unit
	 * Joule 					"J"
	 * Erg 					"e"
	 * Thermodynamic calorie 	"c"
	 * IT calorie 				"cal"
	 * Electron volt 			"eV"
	 * Horsepower-hour 		"HPh"
	 * Watt-hour 				"Wh"
	 * Foot-pound 				"flb"
	 * BTU 					"BTU"
	 * <p/>
	 * Power 					From_unit or to_unit
	 * Horsepower 				"HP"
	 * Watt 					"W"
	 * <p/>
	 * Magnetism 				From_unit or to_unit
	 * Tesla 					"T"
	 * Gauss 					"ga"
	 * <p/>
	 * Temperature 			From_unit or to_unit
	 * Degree Celsius 			"C"
	 * Degree Fahrenheit 		"F"
	 * Degree Kelvin 			"K"
	 * <p/>
	 * Liquid measure 			From_unit or to_unit
	 * Teaspoon 				"tsp"
	 * Tablespoon 				"tbs"
	 * Fluid ounce 			"oz"
	 * Cup 					"cup"
	 * U.S. pint 				"pt"
	 * U.K. pint 				"uk_pt"
	 * Quart 					"qt"
	 * Gallon 					"gal"
	 * Liter 					"l"
	 * <p/>
	 * <p/>
	 * The following abbreviated unit prefixes can be prepended to any metric
	 * from_unit or to_unit.
	 * <p/>
	 * Prefix Multiplier Abbreviation
	 * exa 1E+18 "E"
	 * peta 1E+15 "P"
	 * tera 1E+12 "T"
	 * giga 1E+09 "G"
	 * mega 1E+06 "M"
	 * kilo 1E+03 "k"
	 * hecto 1E+02 "h"
	 * dekao 1E+01 "e"
	 * deci 1E-01 "d"
	 * centi 1E-02 "c"
	 * milli 1E-03 "m"
	 * micro 1E-06 "u"
	 * nano 1E-09 "n"
	 * pico 1E-12 "p"
	 * femto 1E-15 "f"
	 * atto 1E-18 "a"
	 * <p/>
	 * <p/>
	 * Remarks
	 * <p/>
	 * If the input data types are incorrect, CONVERT returns the #VALUE! error value.
	 * If the unit does not exist, CONVERT returns the #N/A error value.
	 * If the unit does not support an abbreviated unit prefix, CONVERT returns the #N/A error value.
	 * If the units are in different groups, CONVERT returns the #N/A error value.
	 * Unit names and prefixes are case-sensitive.
	 */
	private static int findUnits( String u, String[] units )
	{
		boolean bFound = false;
		for( int i = 0; i < units.length; i++ )
		{
			if( u.equals( units[i] ) )
			{
				return i;
			}
		}
		return -1;
	}

	private static double prefixMultiplier( String p, String[] prefixes )
	{
		double multiplier = 1.0;
		if( p.equals( "" ) )
		{
			return multiplier;
		}
		for( int i = 0; i < prefixes.length; i++ )
		{
			if( p.equals( prefixes[i] ) )
			{
				switch( i )
				{
					case 0:    // "E" 	exa
						multiplier = 1E+18;
						break;
					case 1: // "P" 	peta
						multiplier = 1E+15;
						break;
					case 2: // "T" 	tera
						multiplier = 1E+12;
						break;
					case 3: // "G" 	giga
						multiplier = 1E+09;
						break;
					case 4: // "M"	mega
						multiplier = 1E+06;
						break;
					case 5: // "k"	kilo
						multiplier = 1E+03;
						break;
					case 6: // "h"	hecto
						multiplier = 1E+02;
						break;
					case 7: // "e"	dekao
						multiplier = 1E+01;
						break;
					case 8: // "d" 	deci
						multiplier = 1E-01;
						break;
					case 9: // "c"	centi
						multiplier = 1E-02;
						break;
					case 10: // "m"	milli
						multiplier = 1E-03;
						break;
					case 11: // "u"	micro
						multiplier = 1E-06;
						break;
					case 12: // "n"	nano
						multiplier = 1E-09;
						break;
					case 13: // "p"	pico
						multiplier = 1E-12;
						break;
					case 14: // "f"	femto
						multiplier = 1E-15;
						break;
					case 15: // "a"	atto
						multiplier = 1E-18;
						break;
				}
				return multiplier;
			}
		}

		return multiplier;
	}

	protected static Ptg calcConvert( Ptg[] operands )
	{
		if( operands.length < 3 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "CONVERT" );
		double number = operands[0].getDoubleVal();
		String fromUnits = operands[1].getString().trim();
		String toUnits = operands[2].getString().trim();

		String[] allUnits = {
				"g",
				"sg",
				"lbm",
				"u",
				"ozm",
				"m",
				"mi",
				"Nmi",
				"in",
				"ft",
				"yd",
				"ang",
				"Pica",
				"yr",
				"day",
				"hr",
				"mn",
				"sec",
				"Pa",
				"atm",
				"mmHg",
				"N",
				"dyn",
				"lbf",
				"J",
				"e",
				"c",
				"cal",
				"eV",
				"HPh",
				"Wh",
				"flb",
				"BTU",
				"HP",
				"W",
				"T",
				"ga",
				"C",
				"F",
				"K",
				"tsp",
				"tbs",
				"oz",
				"cup",
				"pt",
				"uk_pt",
				"qt",
				"gal",
				"l"
		};
		String[] weightUnits = { "g", "sg", "lbm", "u", "ozm" };
		String[] distanceUnits = { "m", "mi", "Nmi", "in", "ft", "yd", "ang", "Pica" };
		String[] timeUnits = { "yr", "day", "hr", "mn", "sec" };
		String[] pressureUnits = { "Pa", "atm", "mmHg" };
		String[] forceUnits = { "N", "dyn", "lbf" };
		String[] energyUnits = { "J", "e", "c", "cal", "eV", "HPh", "Wh", "flb", "BTU" };
		String[] powerUnits = { "HP", "W" };
		String[] magnetismUnits = { "T", "ga" };
		String[] temperatureUnits = { "C", "F", "K" };
		String[] liquidMeasureUnits = { "tsp", "tbs", "oz", "cup", "pt", "uk_pt", "qt", "gal", "l" };

		// for any metric unit, may be prefixed with 
		String fromPrefix = "";
		String toPrefix = "";
		String[] metricPrefixes = { "E", "P", "T", "G", "M", "k", "h", "e", "d", "c", "m", "u", "n", "p", "f", "a" };

		// first, see if fromUnits and toUnits are in list of acceptable units
		if( findUnits( fromUnits, allUnits ) < 0 )
		{ // doesn't match; strip prefix and try again
			if( fromUnits.length() > 1 )
			{
				fromPrefix = fromUnits.substring( 0, 1 );
				fromUnits = fromUnits.substring( 1 );
				// now recheck
				if( findUnits( fromUnits, allUnits ) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}
				// make sure that prefix is acceptable
				if( findUnits( fromPrefix, metricPrefixes ) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}
		if( findUnits( toUnits, allUnits ) < 0 )
		{// doesn't match; strip prefix and try again
			if( toUnits.length() > 1 )
			{
				toPrefix = toUnits.substring( 0, 1 );
				toUnits = toUnits.substring( 1 );
				// now recheck
				if( findUnits( toUnits, allUnits ) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}
				// make sure that prefix is acceptable
				if( findUnits( toPrefix, metricPrefixes ) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// at here, we know that the prefixes and units are found, but we don't know if they match or make sense ...
		double result = 0;
		double from = 0;
		int i;
		int j = -1;

		// WEIGHT conversion
		if( (i = findUnits( fromUnits, weightUnits )) >= 0 )
		{
			j = findUnits( toUnits, weightUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				// { "g", "sg", "lbm", "u", "ozm" }
				// get fromUnit in grams
				switch( i )
				{
					// from:
					case 0: // "g"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: // "sg"
						from = 14593.84241892870000000000 * number;
						break;
					case 2: // "lbm"
						from = 453.5923097488115 * number;
						break;
					case 3: // "u"
						from = 1.660531004604650E-24 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 4: // "ozm"
						from = 28.349515207973 * number;
						break;
				}
				// now convert
				switch( j )
				{
					case 0:    // "g"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: // "sg"
						result = from * 0.00006852205000534780;
						break;
					case 2: // "lbm"
						result = from * 0.00220462291469134000;
						break;
					case 3: // "u"
						result = from * 6.02217000000000000000E+23 / prefixMultiplier( toPrefix, metricPrefixes );
						;
						break;
					case 4: // "ozm"
						result = from * 0.03527397180036270000;
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// DISTANCE conversion
		if( (j == -1) && ((i = findUnits( fromUnits, distanceUnits )) >= 0) )
		{
			j = findUnits( toUnits, distanceUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//  "m", "mi", "Nmi", "in", "ft", "yd", "ang", "Pica" ;
				// get fromUnits in m
				switch( i )
				{
					case 0:    //"m"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: //"mi"
						from = 1609.344000000000 * number;
						break;
					case 2: //"Nmi"
						from = 1852.000000000000 * number;
						break;
					case 3: //"in"
						from = 0.025400000000 * number;
						break;
					case 4: // "ft"
						from = 0.304800000000 * number;
						break;
					case 5: // "yd"
						from = 0.914400000300 * number;
						break;
					case 6: // "ang"
						from = 0.000000000100 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 7: // "Pica"
						from = 0.00035277777777780000 * number;
						break;
				}
				switch( j )
				{
					case 0:    //"m"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: //"mi"
						result = 0.00062137119223733400 * from;
						break;
					case 2: //"Nmi"
						result = 0.00053995680345572400 * from;
						break;
					case 3: //"in"
						result = 39.37007874015750000000 * from;
						break;
					case 4: // "ft"
						result = 3.28083989501312000000 * from;
						break;
					case 5: // "yd"
						result = 1.09361329797891000000 * from;
						break;
					case 6: // "ang"
						result = 10000000000.000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 7: // "Pica"
						result = 2834.64566929116000000000 * from;
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// TIME conversion
		if( (j == -1) && ((i = findUnits( fromUnits, timeUnits )) >= 0) )
		{
			j = findUnits( toUnits, timeUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"yr", "day", "hr", "mn", "sec"
				switch( i )
				{
					case 0:    // "yr"
						from = number;
						break;
					case 1: // "day"
						from = 0.00273785078713210000 * number;
						break;
					case 2: // "hr"
						from = 0.00011407711613050400 * number;
						break;
					case 3: // "mn"
						from = 0.00000190128526884174 * number;
						break;
					case 4: // "sec"
						from = 0.00000003168808781403 * number;
						break;
				}
				switch( j )
				{
					case 0:    // "yr"
						result = from;
						break;
					case 1: // "day"
						result = 365.250000000000 * from;
						break;
					case 2: // "hr"
						result = 8766.000000000000 * from;
						break;
					case 3: // "mn"
						result = 525960.000000000000 * from;
						break;
					case 4: // "sec"
						result = 31557600.000000000000 * from;
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// PRESSURE conversion
		if( (j == -1) && ((i = findUnits( fromUnits, pressureUnits )) >= 0) )
		{
			j = findUnits( toUnits, pressureUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"Pa", "atm", "mmHg"
				switch( i )
				{
					case 0:    // "Pa"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: // "atm"
						from = 101324.99658300000000000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 2: // "mmHg"
						from = 133.32236392500000000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
				}
				switch( j )
				{
					case 0:    // "Pa"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: // "atm"
						result = 0.00000986923299998193 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 2: // "mmHg"
						result = 0.00750061707998627000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// FORCE conversion
		if( (j == -1) && ((i = findUnits( fromUnits, forceUnits )) >= 0) )
		{
			j = findUnits( toUnits, forceUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"N", "dyn", "lbf"
				switch( i )
				{
					case 0:    // "N"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: // "dyn"
						from = 0.000010000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 2: // "lbf"
						from = 4.448222000000 * number;
						break;
				}
				switch( j )
				{
					case 0:    // "N"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: // "dyn"
						result = 100000.000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 2: // "lbf"
						result = 0.22480892365533900000 * from;
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// ENERGY conversion
		if( (j == -1) && ((i = findUnits( fromUnits, energyUnits )) >= 0) )
		{
			j = findUnits( toUnits, energyUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"J", "e", "c", "cal", "eV", "HPh", "Wh", "flb", "BTU"
				switch( i )
				{
					case 0: // "J"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: // "e"
						from = 0.00000010000004806570 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 2: // "c"
						from = 4.18399101363672000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 3: // "cal"
						from = 4.18679484613929000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 4: // "eV"
						from = 1.60217646E-19 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 5:    // "HPh"
						from = 2684517.41316170000000000000 * number;
						break;
					case 6:    // "Wh"
						from = 3599.99820554720000000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 7: // "flb"
						from = 0.04214000032364240000 * number;
						break;
					case 8: // "BTU"
						from = 1055.05813786749000000000 * number;
						break;
				}
				switch( j )
				{
					case 0: // "J"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: // "e"
						result = 9999995.19343231000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 2: // "c"
						result = 0.23900624947346700000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 3: // "cal"
						result = 0.23884619064201700000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 4: // "eV"
						result = 6241457000000000000.00000000000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 5:    // "HPh"
						result = 0.00000037250643080100 * from;
						break;
					case 6:    // "Wh"
						result = 0.00027777791623871100 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 7: // "flb"
						result = 23.73042221926510000000 * from;
						break;
					case 8: // "BTU"
						result = 0.00094781506734901500 * from;
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// POWER conversion
		if( (j == -1) && ((i = findUnits( fromUnits, powerUnits )) >= 0) )
		{
			j = findUnits( toUnits, powerUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"HP", "W"
				switch( i )
				{
					case 0: // "HP"
						from = number;
						break;
					case 1: // "W"
						from = 0.00134102006031908000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
				}
				switch( j )
				{
					case 0: // "HP"
						result = from;
						break;
					case 1: // "W"
						result = 745.70100000000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// MAGNETISM conversion
		if( (j == -1) && ((i = findUnits( fromUnits, magnetismUnits )) >= 0) )
		{
			j = findUnits( toUnits, magnetismUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"T", "ga"
				switch( i )
				{
					case 0: // "T"
						from = number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
					case 1: // "ga"
						from = 0.000100000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
				}
				switch( j )
				{
					case 0: // "T"
						result = from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
					case 1: // "ga"
						result = 10000.000000000000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// TEMPERATURE conversion
		if( (j == -1) && ((i = findUnits( fromUnits, temperatureUnits )) >= 0) )
		{
			j = findUnits( toUnits, temperatureUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"C", "F", "K"
				switch( i )
				{
					case 0: // "C"
						from = number;
						break;
					case 1: // "F"
						from = (number - 32) / 1.8;
						break;
					case 2: // "K"
						from = (number * prefixMultiplier( fromPrefix, metricPrefixes )) - 273.15;
						break;
				}
				switch( j )
				{
					case 0: // "C"
						result = from;
						break;
					case 1: // "F"
						result = from * 1.8 + 32;
						break;
					case 2: // "K"
						result = (273.15000000000000000000 + from) / prefixMultiplier( toPrefix, metricPrefixes );
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		// LIQUID MEASURE conversion
		if( (j == -1) && ((i = findUnits( fromUnits, liquidMeasureUnits )) >= 0) )
		{
			j = findUnits( toUnits, liquidMeasureUnits );
			if( j > -1 )
			{ // both fromUnits and toUnits are in same family
				//"tsp", "tbs", "oz", "cup", "pt", "uk_pt", "qt", "gal", "l"
				switch( i )
				{
					case 0:    // "tsp"
						from = number;
						break;
					case 1:    // "tbs"
						from = 3.00000000000000000000 * number;
						break;
					case 2:  // "oz"
						from = 6.00000000000000000000 * number;
						break;
					case 3: // "cup"
						from = 48.00000000000000000000 * number;
						break;
					case 4: // "pt"
						from = 96.00000000000000000000 * number;
						break;
					case 5:    // "uk_pt"
						from = 115.26600000000000000000 * number;
						break;
					case 6:    // "qt"
						from = 192.00000000000000000000 * number;
						break;
					case 7:    // "gal"
						from = 768.00000000000000000000 * number;
						break;
					case 8: // "l"
						from = 202.84000000000000000000 * number * prefixMultiplier( fromPrefix, metricPrefixes );
						break;
				}
				switch( j )
				{
					case 0:    // "tsp"
						result = from;
						break;
					case 1:    // "tbs"
						result = 0.33333333333333300000 * from;
						break;
					case 2:  // "oz"
						result = 0.16666666666666700000 * from;
						break;
					case 3: // "cup"
						result = 0.02083333333333330000 * from;
						break;
					case 4: // "pt"
						result = 0.01041666666666670000 * from;
						break;
					case 5:    // "uk_pt"
						result = 0.00867558516821960000 * from;
						break;
					case 6:    // "qt"
						result = 0.00520833333333333000 * from;
						break;
					case 7:    // "gal"
						result = 0.00130208333333333000 * from;
						break;
					case 8: // "l"
						result = 0.00492999408400710000 * from / prefixMultiplier( toPrefix, metricPrefixes );
						break;
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
		}

		PtgNumber pnum = new PtgNumber( result );
		log.debug( "Result from CONVERT= " + pnum.getString() );
		return pnum;
	}

	/**
	 * DEC2BIN
	 * Converts a decimal number to binary
	 */
	protected static Ptg calcDec2Bin( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "DEC2BIN" );
		int dec = operands[0].getIntVal();
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		if( (dec < -512) || (dec > 511) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		String bString = Integer.toBinaryString( dec );
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			bString = bString.substring( Math.max( bString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( bString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			bString = ("0000000000" + bString);    // maximum= 10 bits
			bString = bString.substring( bString.length() - places );
		}
		PtgStr pstr = new PtgStr( bString );
		log.debug( "Result from DEC2BIN= " + pstr.getString() );
		return pstr;
	}

	/**
	 * DEC2HEX
	 * Converts a decimal number to hexadecimal
	 */
	protected static Ptg calcDec2Hex( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "DEC2HEX" );
		long dec = PtgCalculator.getLongValue( operands[0] );
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}
		if( (dec < -549755813888L) || (dec > 549755813887L) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		String hString = Long.toHexString( dec ).toUpperCase();
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			hString = hString.substring( Math.max( hString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( hString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			hString = ("0000000000" + hString);    // maximum= 10 places
			hString = hString.substring( hString.length() - places );

		}
		PtgStr pstr = new PtgStr( hString );
		log.debug( "Result from DEC2HEX= " + pstr.getString() );
		return pstr;
	}

	/**
	 * DEC2OCT
	 * Converts a decimal number to octal
	 */
	protected static Ptg calcDec2Oct( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "DEC2OCT" );
		long dec = PtgCalculator.getLongValue( operands[0] );
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}
		if( (dec < -536870912L) || (dec > 536870911L) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		String oString = Long.toOctalString( dec );
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			oString = oString.substring( Math.max( oString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( oString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			oString = ("0000000000" + oString);    // maximum= 10 places
			oString = oString.substring( oString.length() - places );

		}
		PtgStr pstr = new PtgStr( oString );
		log.debug( "Result from DEC2OCT= " + pstr.getString() );
		return pstr;
	}

	/**
	 * DELTA
	 * Tests whether two values are equal
	 */
	protected static Ptg calcDelta( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "DELTA" );

		double number1 = operands[0].getDoubleVal();
		double number2 = 0;
		if( operands.length > 1 )
		{
			number2 = operands[1].getDoubleVal();
		}
		int result = 0;
		if( number1 == number2 )
		{
			result = 1;
		}

		PtgNumber pnum = new PtgNumber( result );
		log.debug( "Result from DELTA= " + pnum.getString() );
		return pnum;
	}

	/**
	 * helper erf calc - seems to work *ALRIGHT* for values over 1
	 * NOTE: not accurate to 9 digits for every case
	 *
	 * @param x
	 * @return
	 */
	private static double erf_try1( double x )
	{
		double t = x;
		double x2 = Math.pow( x, 2 );
		for( double n = 1000; n >= 0.5; n -= 0.5 )
		{
			t = x + n / t;
		}
		t = 1.0 / t;
		double tt = Math.exp( -x2 ) / Math.sqrt( Math.PI );
		return (1 - (tt * t));
	}

	/**
	 * ERF
	 * Returns the error function integrated between lower_limit and upper_limit.
	 * ERF(lower_limit,upper_limit)
	 * <p/>
	 * Lower_limit     is the lower bound for integrating ERF.
	 * Upper_limit     is the upper bound for integrating ERF. If omitted, ERF integrates between zero and lower_limit.
	 * <p/>
	 * With a single argument ERF returns the error function, defined as
	 * erf(x) = 2/sqrt(pi)* integral from 0 to x of exp(-t*t) dt.
	 * If two arguments are supplied, they are the lower and upper limits of the integral.
	 * NOTE: Accuracy is not always to 9 digits
	 * NOTE: Version with two parameters is NOT supported
	 */
	protected static Ptg calcErf( Ptg[] operands ) throws FunctionNotSupportedException
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		try
		{
			double lower_limit = operands[0].getDoubleVal();
			double upper_limit = Double.NaN;    //lower_limit;
			if( operands.length == 2 )
			{
				upper_limit = operands[1].getDoubleVal();
			}
			// If lower_limit is negative, ERF returns the #NUM! error value.
			// If upper_limit is negative, ERF returns the #NUM! error value.
//			if (lower_limit < 0 /*|| upper_limitupper_limit < 0*/) return new PtgErr(PtgErr.ERROR_NUM);
			boolean neg = (lower_limit < 0);
			lower_limit = Math.abs( lower_limit );

			double result;
			double limit = lower_limit;
			/*
			// try this: from "Computation of the error function erf in arbitrary precision with correct rounding"
			double r= 0;
			double r1= 0;
			double estimate= (2/Math.sqrt(Math.PI))*(limit - Math.pow(limit, 3)/3.0);
			double convergence= Math.pow(2, estimate-15); 
			for (int i= 0, n= 0; n < 100; i++, n++) {
				double factor= 2.0*n + 1.0;
				double z= Math.pow(limit, factor);
				double zz= (MathFunctionCalculator.factorial(n)*factor);
				double zzz= z/zz;
				if ((i % 2)!=0)
					r= r-zzz;
				else
					r= r+zzz;
				
				if (Math.abs(r)-r1) <
				
				r1= Math.abs(r1);
			}
			result= r*(2.0/Math.sqrt(Math.PI));*/

			if( limit < 0.005 )
			{
				/* A&S 7.1.1 - good to at least 6 digts ... but not for larger values ... sigh ...*/
				double r = 0;
				for( int i = 0, n = 0; n < 12; i++, n++ )
				{
					double factor = (2.0 * n) + 1.0;
					double z = Math.pow( limit, factor );
					double zz = (MathFunctionCalculator.factorial( n ) * factor);
					double zzz = z / zz;
					if( (i % 2) != 0 )
					{
						r = r - zzz;
					}
					else
					{
						r = r + zzz;
					}
				}
				result = r * (2.0 / Math.sqrt( Math.PI ));
			}
			else
			{
				result = erf_try1( limit );
			}

			if( neg )
			{
				result *= -1;
			}

			if( !Double.isNaN( upper_limit ) )
			{    // Erf(upper)-Erf(lower)
				Ptg result2 = calcErf( new Ptg[]{ operands[1] } );
				if( result2 instanceof PtgNumber )
				{
					result = result2.getDoubleVal() - result;
				}
				else
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}

			return new PtgNumber( result );

		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	static final double PI_SQRT = Math.sqrt( Math.PI );
	static final double TSQPI = 2. / PI_SQRT;
	private static boolean firstCall = true;

	/**
	 * Calculate the error function erf.
	 *
	 * @param x the argument
	 * @return the value erf(x)
	 */
	public static double erf( final double x )
	{
		//if (type==1) {
		return erfAS( x );
		//}
		//return erfCody(x);
	}

	/**
	 * Calculate the remaining error function erfc.
	 *
	 * @param x the argument
	 * @return the value erfc(x)
	 */
	public static double erfc( final double x )
	{
		//if (type==1) {
		return erfcAS( x );
		//}
		//return erfcCody(x);
	}

	/**
	 * Internal helper method to calculate the error function at value x.
	 * This code is  based on a Fortran implementation from
	 * <a href="http://www.netlib.org/specfun/erf">W. J. Cody</a>.
	 * Refactored by N.Wulff for Java.
	 * <p/>
	 * Cody vs AS algiorithm
	 *
	 * @param x the argument
	 * @return the approximation of erf(x)
	 */
	private static double erfCody( final double x )
	{
		double result = 0;
		double y = Math.abs( x );
		if( firstCall )
		{
			firstCall = false;
		}
		if( y <= THRESHOLD )
		{
			result = x * calcLower( y );
		}
		else
		{
			result = calcUpper( y );
			result = (0.5 - result) + 0.5;
			if( x < 0 )
			{
				result = -result;
			}
		}
		return result;
	}

	/**
	 * Internal helper method to calculate the erfc functions.
	 * This code is  based on a Fortran implementation from
	 * <a href="http://www.netlib.org/specfun/erf">W. J. Cody</a>.
	 * Refactored by N.Wulff for Java.
	 *
	 * @param x the argument
	 * @return the approximation erfc(x)
	 */
	private static double erfcCody( final double x )
	{
		double result = 0;
		double y = Math.abs( x );
		if( firstCall )
		{
			firstCall = false;
		}
		if( y <= THRESHOLD )
		{
			result = x * calcLower( y );
			result = 1 - result;
		}
		else
		{
			result = calcUpper( y );
			if( x < 0 )
			{
				result = 2.0 - result;
			}
		}
		return result;
	}

	/**
	 * Internal helper method to calculate the erf/erfc functions.
	 * This code is  based on a Fortran implementation from
	 * <a href="http://www.netlib.org/specfun/erf">W. J. Cody</a>.
	 * Refactored by N.Wulff for Java.
	 *
	 * @param y the value y=abs(x)<=THRESHOLD
	 * @return the series expansion
	 */
	private static double calcLower( final double y )
	{
		double result;
		double ySq;
		double xNum;
		double xDen;
		ySq = 0.0;
		if( y > X_SMALL )
		{
			ySq = y * y;
		}
		xNum = ERF_A[4] * ySq;
		xDen = ySq;
		for( int i = 0; i < 3; i++ )
		{
			xNum = (xNum + ERF_A[i]) * ySq;
			xDen = (xDen + ERF_B[i]) * ySq;
		}
		result = (xNum + ERF_A[3]) / (xDen + ERF_B[3]);
		return result;
	}

	/**
	 * Internal helper method to calculate the erf/erfc functions.
	 * This code is  based on a Fortran implementation from
	 * <a href="http://www.netlib.org/specfun/erf">W. J. Cody</a>.
	 * Refactored by N.Wulff for Java.
	 *
	 * @param y the value y=abs(x)>THRESHOLD
	 * @return the series expansion
	 */
	private static double calcUpper( final double y )
	{
		double result;
		double ySq;
		double xNum;
		double xDen;
		if( y <= 4.0 )
		{
			xNum = ERF_C[8] * y;
			xDen = y;
			for( int i = 0; i < 7; i++ )
			{
				xNum = (xNum + ERF_C[i]) * y;
				xDen = (xDen + ERF_D[i]) * y;
			}
			result = (xNum + ERF_C[7]) / (xDen + ERF_D[7]);
		}
		else
		{
			result = 0.0;
			if( y >= X_HUGE )
			{
				result = SQRPI / y;
			}
			else
			{
				ySq = 1.0 / (y * y);
				xNum = ERF_P[5] * ySq;
				xDen = ySq;
				for( int i = 0; i < 4; i++ )
				{
					xNum = (xNum + ERF_P[i]) * ySq;
					xDen = (xDen + ERF_Q[i]) * ySq;
				}
				result = ySq * (xNum + ERF_P[4]) / (xDen + ERF_Q[4]);
				result = (SQRPI - result) / y;
			}
		}
		ySq = Math.round( y * 16.0 ) / 16.0;
		double del = (y - ySq) * (y + ySq);
		result = Math.exp( -ySq * ySq ) * Math.exp( -del ) * result;
		return result;
	}

	/**
	 * Calculate the error function at value x.
	 * AS 7.1.5/7.1.26
	 *
	 * @param x the argument
	 * @return the value erf(x)
	 */
	private static double erfAS( final double x )
	{
		if( firstCall )
		{
			firstCall = false;
		}
		if( x < 0 )
		{
			return -erfAS( -x );
		}
		if( x < 2 )
		{
			return erfSeries( x );
		}
		return erfRational( x );
	}

	/**
	 * Calculate the remaining erfc error function at value x.
	 * AS 7.1.5/7.1.26
	 *
	 * @param x the argument
	 * @return the value erfc(x)
	 */
	private static double erfcAS( final double x )
	{
		if( firstCall )
		{
			firstCall = false;
		}
		return 1 - erfAS( x );
	}

	/**
	 * Series expansion from A&S 7.1.5.
	 *
	 * @param x the argument
	 * @return erf(x)
	 */
	private static double erfSeries( final double x )
	{
		final double eps = 1.E-8; // we want only ~1.E-7
		final int kmax = 50; // this can be reached with ~30-40
		double an;
		double ak = x;
		double erfo;
		double erf = ak;
		int k = 1;
		do
		{
			erfo = erf;
			ak *= -x * x / k;
			an = ak / ((2.0 * k) + 1.0);
			erf += an;
		} while( !hasConverged( erf, erfo, eps, ++k, kmax ) );
		return TSQPI * erf;
	}

	/**
	 * Indicate if an iterative algorithm has RELATIVE converged.
	 * <hr/>
	 * <b>Note</b>:
	 * HasConverged throws an ArithmeticException if more than max calls
	 * have been made. Choose hasReacherAccuracy if this is not desired.
	 * <hr/>
	 *
	 * @param xn  the actual argument x[n]
	 * @param xo  the older argument x[n-1]
	 * @param eps the accuracy to reach
	 * @param n   the actual iteration counter
	 * @param max the maximal number of iterations
	 * @return flag indicating if accuracy is reached.
	 */
	public static boolean hasConverged( final double xn, final double xo, final double eps, final int n, final int max )
	{
		if( hasReachedAccuracy( xn, xo, eps ) )
		{
			return true;
		}
		if( n >= max )
		{
			throw new ArithmeticException();
		}
		return false;
	}

	/**
	 * Indicate if xn and xo have the relative/absolute accuracy epsilon.
	 * In case that the true value is less than one this is based
	 * on the absolute difference, otherwise on the relative difference:
	 * <pre>
	 *     2*|x[n]-x[n-1]|/|x[n]+x[n-1]| < eps
	 * </pre>
	 *
	 * @param xn  the actual argument x[n]
	 * @param xo  the older argument x[n-1]
	 * @param eps accuracy to reach
	 * @return flag indicating if accuracy is reached.
	 */
	public static boolean hasReachedAccuracy( final double xn, final double xo, final double eps )
	{
		double z = Math.abs( xn + xo ) / 2;
		double error = Math.abs( xn - xo );
		if( z > 1 )
		{
			error /= z;
		}
		return error <= eps;
	}

	/**
	 * Rational approximation A&S 7.1.26 with accuracy 1.5E-7.
	 *
	 * @param x the argument
	 * @return erf(x)
	 */
	private static double erfRational( final double x )
	{
	         /*  coefficients for A&S 7.1.26. */
		final double[] a = {
				.254829592, -.284496736, 1.421413741, -1.453152027, 1.061405429
		};
	         /*  constant for A&S 7.1.26 */
		final double p = .3275911;
		double erf;
		double r = 0;
		double t = 1.0 / (1 + (p * x));
		for( int i = 4; i >= 0; i-- )
		{
			r = a[i] + r * t;
		}
		erf = 1 - t * r * Math.exp( -x * x );
		return erf;
	}

	// ===========================================================================
	/**
	 * Nominator coefficients for approximation to erf in first interval.
	 */
	private static final double[] ERF_A = {
			3.16112374387056560E00, 1.13864154151050156E02, 3.77485237685302021E02, 3.20937758913846947E03, 1.85777706184603153E-1
	};
	/**
	 * Denominator coefficients for approximation to erf in first interval.
	 */
	private static final double[] ERF_B = {
			2.36012909523441209E01, 2.44024637934444173E02, 1.28261652607737228E03, 2.84423683343917062E03
	};
	// ===========================================================================
	/**
	 * Nominator coefficients for approximation to erfc in second interval.
	 */
	private static final double[] ERF_C = {
			5.64188496988670089E-1,
			8.88314979438837594E0,
			6.61191906371416295E01,
			2.98635138197400131E02,
			8.81952221241769090E02,
			1.71204761263407058E03,
			2.05107837782607147E03,
			1.23033935479799725E03,
			2.15311535474403846E-8
	};
	/**
	 * Denominator coefficients for approximation to erfc in second interval.
	 */
	private static final double[] ERF_D = {
			1.57449261107098347E01,
			1.17693950891312499E02,
			5.37181101862009858E02,
			1.62138957456669019E03,
			3.29079923573345963E03,
			4.36261909014324716E03,
			3.43936767414372164E03,
			1.23033935480374942E03
	};
	// ===========================================================================
	/**
	 * Nominator coefficients for approximation to erfc in third interval.
	 */
	private static final double[] ERF_P = {
			3.05326634961232344E-1,
			3.60344899949804439E-1,
			1.25781726111229246E-1,
			1.60837851487422766E-2,
			6.58749161529837803E-4,
			1.63153871373020978E-2
	};
	/**
	 * Denominator coefficients for approximation to erfc in third interval.
	 */
	private static final double[] ERF_Q = {
			2.56852019228982242, 1.87295284992346047, 5.27905102951428412E-1, 6.05183413124413191E-2, 2.33520497626869185E-3
	};
	// ===========================================================================
	static final double THRESHOLD = 0.46875;
	static final double SQRPI = 1. / PI_SQRT;
	static final double X_INF = Double.MAX_VALUE;
	static final double X_MIN = Double.MIN_VALUE;
	// private static final double X_NEG = -9.38241396824444;
	static final double X_NEG = -Math.sqrt( Math.log( X_INF / 2 ) );
	static final double X_SMALL = getDEPS();
	static final double X_HUGE = 1.0 / (2.0 * Math.sqrt( X_SMALL ));
	static final double X_MAX = Math.min( X_INF, (1.0 / (PI_SQRT * X_MIN)) );
	// static final double X_BIG = 9.194E0;
	static final double X_BIG = 26.543;

	private static final float FEPS_START = 2.E-6f;
	/** double valued machine precision. */
	/**
	 * float valued machine precision.
	 */
	//public static final float FEPS;
	static double getDEPS()
	{
		/**
		 * Calculate the machine accuracy,
		 * which is the smallest eps with
		 * 1<1+eps
		 */
		float feps = FEPS_START;
		float fy = 1.0f + feps;
		while( fy > 1.0f )
		{
			feps /= 2.0f;
			fy = 1.0f + feps;
		}
		//FEPS = feps;
		double deps = feps * FEPS_START;
		double dy = 1.0 + deps;
		while( dy > 1.0 )
		{
			deps /= 2.0;
			dy = 1.0 + deps;
		}
		//DEPS =  deps;
		//if (DEBUG)

		//  Logger.logInfo(format("feps:%8.2E  deps:%8.3G", FEPS, DEPS));
		return deps;
	}	
/*
	ERFC
	 Returns the complementary error function
*/

	/**
	 * GESTEP
	 * Tests whether a number is greater than a threshold value
	 */
	protected static Ptg calcGEStep( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "GESTEP" );

		double number = operands[0].getDoubleVal();
		double step = 0;
		if( operands.length > 1 )
		{
			step = operands[1].getDoubleVal();
		}
		int result = 0;
		if( number >= step )
		{
			result = 1;
		}

		PtgNumber pnum = new PtgNumber( result );
		log.debug( "Result from GESTEP= " + pnum.getString() );
		return pnum;
	}

	/**
	 * HEX2BIN
	 * Converts a hexadecimal number to binary
	 */
	protected static Ptg calcHex2Bin( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "HEX2BIN" );
		String hString = operands[0].getString().trim();
		// 10 digits (40 bits) at most
		if( hString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		long dec;
		String bString;
		try
		{
			dec = Long.parseLong( hString, 16 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 549755813888L )        // 2^39 (= signed 40 bits)
			{
				dec -= 1099511627776L;        // 2^40 (= signed 41 bits)
			}
			bString = Long.toBinaryString( dec );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		if( (dec < -512) || (dec > 0x1FF) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			bString = bString.substring( Math.max( bString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( bString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			bString = ("0000000000" + bString);    // maximum= 10 places
			bString = bString.substring( bString.length() - places );
		}
		PtgStr pstr = new PtgStr( bString );
		log.debug( "Result from HEX2BIN= " + pstr.getString() );
		return pstr;
	}

	/**
	 * HEX2DEC
	 * Converts a hexadecimal number to decimal
	 */
	protected static Ptg calcHex2Dec( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "HEX2DEC" );
		String hString = operands[0].getString().trim();
		// 10 digits (40 bits) at most
		if( hString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		long dec;
		String oString;
		try
		{
			dec = Long.parseLong( hString, 16 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 549755813888L )        // 2^39 (= signed 40 bits)
			{
				dec -= 1099511627776L;        // 2^40 (= signed 41 bits)
			}
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		PtgNumber pnum = new PtgNumber( dec );
		log.debug( "Result from HEX2DEC= " + pnum.getVal() );
		return pnum;
	}

	/**
	 * HEX2OCT
	 * Converts a hexadecimal number to octal
	 */
	protected static Ptg calcHex2Oct( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "HEX2OCT" );
		String hString = operands[0].getString().trim();
		// 10 digits (40 bits) at most
		if( hString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		long dec;
		String oString;
		try
		{
			dec = Long.parseLong( hString, 16 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 549755813888L )        // 2^39 (= signed 40 bits)
			{
				dec -= 1099511627776L;        // 2^40 (= signed 41 bits)
			}
			oString = Long.toOctalString( dec );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( (dec < -536870912L) || (dec > 0x1FFFFFFF) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			oString = oString.substring( Math.max( oString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( oString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			oString = ("0000000000" + oString);    // maximum= 10 places
			oString = oString.substring( oString.length() - places );
		}
		PtgStr pstr = new PtgStr( oString );
		log.debug( "Result from HEX2OCT= " + pstr.getString() );
		return pstr;
	}

	/**
	 * imParseComplexNumber
	 * <p/>
	 * used in the following Imaginary-based formulas to parse a complex number into its
	 * real and imaginary components.
	 * Throws a numberformat exception if the complex number is not in the format of:
	 * real + imaginary
	 * real - imaginary
	 * imaginary			(real= 0)
	 * real				(imaginary= 0)
	 * where imaginary coefficient n= ni or nj
	 * *
	 */
	private static Complex imParseComplexNumber( String complexNumber ) throws NumberFormatException
	{
		Complex c = new Complex();
		if( complexNumber.length() > 0 )
		{
			try
			{
				int i = complexNumber.length();
				if( complexNumber.substring( i - 1, i ).equals( "i" ) || complexNumber.substring( i - 1, i ).equals( "j" ) )
				{
					c.suffix = complexNumber.substring( i - 1, i );
					i -= 2;
					while( (i >= 0) && !(complexNumber.substring( i, i + 1 ).equals( "+" ) || complexNumber.substring( i, i + 1 ).equals(
							"-" )) )
					{
						i--;
					}
					if( i < 0 )
					{ // case of "#i" or "#j" i.e. no real and no sign
						complexNumber = "+" + complexNumber;
						i++;
					}
					// get imaginary coefficient + sign
					String s = complexNumber.substring( i, complexNumber.length() - 1 );
					if( s.length() == 1 )
					{    // only a sign; means that the coefficient==1 eg. real-j or real+i
						s += "1";
					}
					c.imaginary = Double.parseDouble( s );
				}
				if( i > 0 )
				{
					c.real = Double.parseDouble( complexNumber.substring( 0, i ) );
				}
			}
			catch( Exception e )
			{
				throw new NumberFormatException();
			}
		}
		return c;
	}

	/**
	 * imGetStr
	 *
	 * @param operands
	 * @return double formatted as an integer if no precision, otherwise rounds to 15
	 */
	private static String imGetExcelStr( double d, int precision )
	{
		String s;
		if( d == (int) d )
		{
			if( (int) d == 1 )
			{
				return "";
			}
			return String.valueOf( (int) d );
		}
		// round to precision - default= 15
		double r = Math.pow( 10, precision );
		d *= r;
		d = Math.round( d );
		d /= r;
		return String.valueOf( d );
	}

	private static String imGetExcelStr( double d )
	{
		return imGetExcelStr( d, 15 );
	}

	/**
	 * IMABS
	 * Returns the absolute value (modulus) of a complex number
	 */
	protected static Ptg calcImAbs( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "IMABS" );
		String complexString = StringTool.allTrim( operands[0].getString() );

		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		// Absolute of a complex number is:
		// square root( real^2 + imaginary^2)		
		double result = Math.sqrt( Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 ) );

		PtgNumber pnum = new PtgNumber( result );
		log.debug( "Result from IMABS= " + pnum.getString() );
		return pnum;
	}

	/**
	 * IMAGINARY
	 * Returns the imaginary coefficient of a complex number
	 */
	protected static Ptg calcImaginary( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "Imaginary" );
		String complexString = StringTool.allTrim( operands[0].getString() );

		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		PtgNumber pnum = new PtgNumber( c.imaginary );
		log.debug( "Result from IMAGINARY= " + pnum.getString() );
		return pnum;
	}

	/**
	 * IMARGUMENT
	 * Returns the argument theta, an angle expressed in radians
	 */
	protected static Ptg calcImArgument( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImArgument" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		double result = Math.atan( c.imaginary / c.real );

		PtgNumber pnum = new PtgNumber( result );
		log.debug( "Result from IMARGUMENT= " + pnum.getString() );
		return pnum;
	}

	/**
	 * IMCONJUGATE
	 * Returns the complex conjugate of a complex number
	 */
	protected static Ptg calcImConjugate( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImCongugate" );
		String complexString = StringTool.allTrim( operands[0].getString() );

		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		String congugate;
		if( (c.real != 0) && (c.imaginary != 0) )
		{
			congugate = imGetExcelStr( c.real ) + ((c.imaginary < 0) ? "+" : "-") + imGetExcelStr( Math.abs( c.imaginary ) ) + c.suffix;
		}
		else if( c.real == 0 )
		{
			congugate = imGetExcelStr( Math.abs( c.imaginary ) ) + c.suffix;
		}
		else
		{
			congugate = imGetExcelStr( c.real );
		}
		PtgStr pstr = new PtgStr( congugate );
		log.debug( "Result from IMCONGUGATE= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMCOS
	 * Returns the cosine of a complex number
	 */
	protected static Ptg calcImCos( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImCos" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// cos(a + bi)= cos(a)cosh(b) - sin(a)sinh(b)i
		double cosh = (Math.pow( Math.E, c.imaginary ) + Math.pow( Math.E, c.imaginary * -1 )) / 2;
		double sinh = (Math.pow( Math.E, c.imaginary ) - Math.pow( Math.E, c.imaginary * -1 )) / 2;
		double a = Math.cos( c.real ) * cosh;
		double b = Math.sin( c.real ) * sinh;

		String imCos;
		if( b < 0 )
		{
			imCos = imGetExcelStr( a ) + "+" + imGetExcelStr( Math.abs( b ) ) + c.suffix;
		}
		else
		{
			imCos = imGetExcelStr( a ) + "-" + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imCos );
		log.debug( "Result from IMCOS= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMDIV
	 * Returns the quotient of two complex numbers, defined as (deep breath):
	 * <p/>
	 * (r1 + i1j)/(r2 + i2j) == ( r1r2 + i1i2 + (r2i1 - r1i2)i ) / (r2^2 + i2^2)
	 */
	protected static Ptg calcImDiv( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "IMDIV" );
		String complexString1 = StringTool.allTrim( operands[0].getString() );
		String complexString2 = StringTool.allTrim( operands[1].getString() );

		Complex c1;
		Complex c2;

		try
		{
			c1 = imParseComplexNumber( complexString1 );
			c2 = imParseComplexNumber( complexString2 );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		double divisor = Math.pow( c2.real, 2 ) + Math.pow( c2.imaginary, 2 );
		double a = (c1.real * c2.real) + (c1.imaginary * c2.imaginary);
		double b = (c1.imaginary * c2.real) - (c1.real * c2.imaginary);
		double c = a / divisor;
		double d = b / divisor;

		String imDiv;
		if( d > 0 )
		{
			imDiv = imGetExcelStr( c ) + "+" + imGetExcelStr( d ) + "i";
		}
		else
		{
			imDiv = imGetExcelStr( c ) + "-" + imGetExcelStr( d ) + "i";
		}

		PtgStr pstr = new PtgStr( imDiv );
		log.debug( "Result from IMDIV= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMEXP
	 * Returns the exponential of a complex number
	 */
	protected static Ptg calcImExp( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImExp" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// Exponential of a complex number x+yi = e^x(cos(y) + sin(y)i)
		double e_x = Math.pow( Math.E, c.real );
		double a = e_x * Math.cos( c.imaginary );
		double b = e_x * Math.sin( c.imaginary );

		String imExp;
		if( b > 0 )
		{
			imExp = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imExp = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imExp );
		log.debug( "Result from IMEXP= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMLN
	 * Returns the natural logarithm of a complex number
	 */
	protected static Ptg calcImLn( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImLn" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// Natural log of a complex number=
		// 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
		double a = Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 );
		a = Math.sqrt( a );
		a = Math.log( a );
		double b = Math.atan( c.imaginary / c.real );

		String imLn;
		if( b > 0 )
		{
			imLn = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imLn = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imLn );
		log.debug( "Result from IMLN= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMLOG10
	 * Returns the base-10 logarithm of a complex number
	 */
	protected static Ptg calcImLog10( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImLog10" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// Natural log of a complex number=
		// 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
		double a = Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 );
		a = Math.sqrt( a );
		a = Math.log( a );
		double b = Math.atan( c.imaginary / c.real );
		// now, convert to base 10 log:
		double logE = Math.log( Math.E ) / Math.log( 10 );
		a = a * logE;
		b = b * logE;

		String imLog10;
		if( b > 0 )
		{
			imLog10 = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imLog10 = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imLog10 );
		log.debug( "Result from IMLOG10= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMLOG2
	 * Returns the base-2 logarithm of a complex number
	 */
	protected static Ptg calcImLog2( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImLog2" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// Natural log of a complex number=
		// 	IMLN(x + yi)= ln(sqrt(x2+y2)) + atan(y/x)i
		double a = Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 );
		a = Math.sqrt( a );
		a = Math.log( a );
		double b = Math.atan( c.imaginary / c.real );
		// now, convert to base 2 log:
		double logE = Math.log( Math.E ) / Math.log( 2 );
		a = a * logE;
		b = b * logE;
		// TODO: Results only correct to 8th precision: WHY???
		String imLog2;
		if( b > 0 )
		{
			imLog2 = imGetExcelStr( a, 8 ) + "+" + imGetExcelStr( b, 8 ) + c.suffix;
		}
		else
		{
			imLog2 = imGetExcelStr( a, 8 ) + imGetExcelStr( b, 8 ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imLog2 );
		log.debug( "Result from IMLOG2= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMPOWER
	 * Returns a complex number raised to any power
	 */
	protected static Ptg calcImPower( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImPower" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		double n = operands[1].getDoubleVal();
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		// A complex number (x + yi) raised to a power n is:
		// sqrt(x^2 + y^2)*cos(n*atan(y/x)) + sqrt(x^2 + y^2)*sin(n*atan(y/x))i 
		double r = Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 );
		r = Math.sqrt( r );
		r = Math.pow( r, n );
		double t = Math.atan( c.imaginary / c.real );
		double a = r * Math.cos( n * t );
		double b = r * Math.sin( n * t );

		String imPower;
		if( b > 0 )
		{
			imPower = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imPower = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imPower );
		log.debug( "Result from IMPOWER= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMPRODUCT
	 * Returns the product of two complex numbers
	 */
	protected static Ptg calcImProduct( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}

		debugOperands( operands, "IMPRODUCT" );
		Ptg[] ops = PtgCalculator.getAllComponents( operands );
		String[] complexStrings = new String[ops.length];
		for( int i = 0; i < ops.length; i++ )
		{
			complexStrings[i] = StringTool.allTrim( ops[i].getString() );
		}

		Complex[] c = new Complex[complexStrings.length];
		for( int i = 0; i < complexStrings.length; i++ )
		{
			try
			{
				c[i] = imParseComplexNumber( complexStrings[i] );
			}
			catch( Exception e )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
		}

		// basically, linear binomial multiplication over n terms
		// (a + bi)(c + di) = (ac-bd) + (ad+bc)i   for n terms
		for( int i = 1; i < c.length; i++ )
		{
			double a = c[0].real;
			double b = c[0].imaginary;
			c[0].real = a * c[i].real - b * c[i].imaginary;
			c[0].imaginary = a * c[i].imaginary + b * c[i].real;
		}

		// Format Result
		String imSum;
		if( c[0].imaginary > 0 )
		{
			imSum = imGetExcelStr( c[0].real ) + "+" + imGetExcelStr( c[0].imaginary ) + c[0].suffix;
		}
		else
		{
			imSum = imGetExcelStr( c[0].real ) + imGetExcelStr( c[0].imaginary ) + c[0].suffix;
		}

		PtgStr pstr = new PtgStr( imSum );
		log.debug( "Result from IMSPRODUCT= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMREAL
	 * Returns the real coefficient of a complex number
	 */
	protected static Ptg calcImReal( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "IMREAL" );
		String complexString = StringTool.allTrim( operands[0].getString() );

		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		PtgNumber pnum = new PtgNumber( c.real );
		log.debug( "Result from IMREAL= " + pnum.getString() );
		return pnum;
	}

	/**
	 * IMSIN
	 * Returns the sine of a complex number
	 */
	protected static Ptg calcImSin( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImSin" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// sin(a + bi)= sin(a)cosh(b) + cos(a)sinh(b)i (Excel doc is wrong!)
		double cosh = (Math.pow( Math.E, c.imaginary ) + Math.pow( Math.E, c.imaginary * -1 )) / 2;
		double sinh = (Math.pow( Math.E, c.imaginary ) - Math.pow( Math.E, c.imaginary * -1 )) / 2;
		double a = Math.sin( c.real ) * cosh;
		double b = Math.cos( c.real ) * sinh;

		String imSin;
		if( b > 0 )
		{
			imSin = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imSin = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imSin );
		log.debug( "Result from IMSIN= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMSQRT
	 * Returns the square root of a complex number
	 */
	protected static Ptg calcImSqrt( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "ImSqrt" );
		String complexString = StringTool.allTrim( operands[0].getString() );
		Complex c;
		try
		{
			c = imParseComplexNumber( complexString );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		// The square root of a complex number (x + yi) is:
		// sqrt(sqrt(x^2 + y^2))*cos(atan(y/x)/2) + 
		//   sqrt(sqrt(x^2 + y^2))*sin(atan(y/x)/2)i 
		double r = Math.pow( c.real, 2 ) + Math.pow( c.imaginary, 2 );
		r = Math.sqrt( r );
		r = Math.sqrt( r );
		double t = Math.atan( c.imaginary / c.real );
		double a = r * Math.cos( t / 2 );
		double b = r * Math.sin( t / 2 );

		String imSqrt;
		if( b > 0 )
		{
			imSqrt = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c.suffix;
		}
		else
		{
			imSqrt = imGetExcelStr( a ) + imGetExcelStr( b ) + c.suffix;
		}

		PtgStr pstr = new PtgStr( imSqrt );
		log.debug( "Result from IMSQRT= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMSUB
	 * Returns the difference of two complex numbers
	 */
	protected static Ptg calcImSub( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "IMSUB" );
		String complexString1 = StringTool.allTrim( operands[0].getString() );
		String complexString2 = StringTool.allTrim( operands[1].getString() );

		Complex c1;
		Complex c2;
		try
		{
			c1 = imParseComplexNumber( complexString1 );
			c2 = imParseComplexNumber( complexString2 );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		// basically, linear binomial subtraction:
		// (a + bi) - (c + di)= (a-c) + (b-d)i
		double a = c1.real - c2.real;
		double b = c1.imaginary - c2.imaginary;

		String imSub;
		if( b > 0 )
		{
			imSub = imGetExcelStr( a ) + "+" + imGetExcelStr( b ) + c1.suffix;    // should have the same suffix
		}
		else
		{
			imSub = imGetExcelStr( a ) + imGetExcelStr( b ) + c1.suffix;
		}

		PtgStr pstr = new PtgStr( imSub );
		log.debug( "Result from IMSUB= " + pstr.getString() );
		return pstr;
	}

	/**
	 * IMSUM
	 * Returns the sum of complex numbers
	 */
	protected static Ptg calcImSum( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}

		debugOperands( operands, "IMSUM" );
		Ptg[] ops = PtgCalculator.getAllComponents( operands );
		String[] complexStrings = new String[ops.length];
		for( int i = 0; i < ops.length; i++ )
		{
			complexStrings[i] = StringTool.allTrim( ops[i].getString() );
		}

		Complex[] c = new Complex[complexStrings.length];
		for( int i = 0; i < complexStrings.length; i++ )
		{
			try
			{
				c[i] = imParseComplexNumber( complexStrings[i] );
			}
			catch( Exception e )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
		}

		// basically, linear binomial addition over n terms
		// (a + bi)+(c + di) = (a+c) + (b+d)i   for n terms
		for( int i = 1; i < c.length; i++ )
		{
			c[0].real = c[0].real + c[i].real;
			c[0].imaginary = c[0].imaginary + c[i].imaginary;
		}

		// Format Result
		String imSum;
		if( c[0].imaginary > 0 )
		{
			imSum = imGetExcelStr( c[0].real ) + "+" + imGetExcelStr( c[0].imaginary ) + c[0].suffix;
		}
		else
		{
			imSum = imGetExcelStr( c[0].real ) + imGetExcelStr( c[0].imaginary ) + c[0].suffix;
		}

		PtgStr pstr = new PtgStr( imSum );
		log.debug( "Result from IMSUM= " + pstr.getString() );
		return pstr;
	}

	/**
	 * OCT2BIN
	 * Converts an octal number to binary
	 */
	protected static Ptg calcOct2Bin( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "OCT2BIN" );
		long l = (long) operands[0].getDoubleVal();    // avoid sci notation
		String oString = String.valueOf( l ).trim();
		// 10 digits at most (=30 bits)
		if( oString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		long dec;
		String bString;
		try
		{
			dec = Long.parseLong( oString, 8 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 536870912L )        // 2^29 (= 30 bits, signed)
			{
				dec -= 1073741824L;        // 2^30 (= 31 bits, signed)
			}
			bString = Long.toBinaryString( dec );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		if( (dec < -512) || (dec > 0777) || (places < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			bString = bString.substring( Math.max( bString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( bString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			bString = ("0000000000" + bString);    // maximum= 10 places
			bString = bString.substring( bString.length() - places );
		}
		PtgStr pstr = new PtgStr( bString );
		log.debug( "Result from OCT2BIN= " + pstr.getString() );
		return pstr;
	}

	/**
	 * OCT2DEC
	 * Converts an octal number to decimal
	 */
	protected static Ptg calcOct2Dec( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "OCT2DEC" );
		long l = (long) operands[0].getDoubleVal(); // avoid sci notation
		String oString = String.valueOf( l ).trim();
		// 10 digits at most (=30 bits)
		if( oString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		long dec;
		try
		{
			dec = Long.parseLong( oString, 8 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 536870912L )        // 2^29 (= 30 bits, signed)
			{
				dec -= 1073741824L;        // 2^30 (= 31 bits, signed)
			}
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		PtgNumber pnum = new PtgNumber( dec );
		log.debug( "Result from OCT2DEC= " + pnum.getVal() );
		return pnum;
	}

	/**
	 * OCT2HEX
	 * Converts an octal number to hexadecimal
	 */
	protected static Ptg calcOct2Hex( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		debugOperands( operands, "OCT2HEX" );
		long l = (long) operands[0].getDoubleVal(); // avoid sci notation
		String oString = String.valueOf( l ).trim();
		// 10 digits at most (=30 bits)
		if( oString.length() > 10 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int places = 0;
		if( operands.length > 1 )
		{
			places = operands[1].getIntVal();
		}

		long dec;
		String hString;
		try
		{
			dec = Long.parseLong( oString, 8 );
			// must det. manually if binary string is negative because parseInt/parseLong does not
			// handle two's complement input!!!
			if( dec >= 536870912L )        // 2^29 (= 30 bits, signed)
			{
				dec -= 1073741824L;        // 2^30 (= 31 bits, signed)
			}
			hString = Long.toHexString( dec ).toUpperCase();
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		if( dec < 0 )
		{    // truncate to 10 digits automatically (should already be two's complement)
			hString = hString.substring( Math.max( hString.length() - 10, 0 ) );
		}
		else if( places > 0 )
		{
			if( hString.length() > places )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			hString = ("0000000000" + hString);    // maximum= 10 bits
			hString = hString.substring( hString.length() - places );
		}

		PtgStr pstr = new PtgStr( hString );
		log.debug( "Result from OCT2HEX= " + pstr.getString() );
		return pstr;
	}

	/*	
	SQRTPI
	 Returns the square root of (number * PI)
*/
	static void debugOperands( Ptg[] operands, String f )
	{
		if( !log.isDebugEnabled() )
		{
			return;
		}
			log.debug( "Operands for " + f );
			for( int i = 0; i < operands.length; i++ )
			{
				String s = operands[i].getString();
				if( !(operands[i] instanceof PtgMissArg) )
				{
					String v = operands[i].getValue().toString();
					log.debug( "\tOperand[" + i + "]=" + s + " " + v );
				}
				else
				{
					log.debug( "\tOperand[" + i + "]=" + s + " is Missing" );
				}
			}
	}

}

class Complex
{
	public double real;
	public double imaginary;
	public String suffix;

	Complex()
	{
		suffix = "i";
	}

	Complex( double r, double i )
	{
		real = r;
		imaginary = i;
		suffix = "i";
	}
}	

