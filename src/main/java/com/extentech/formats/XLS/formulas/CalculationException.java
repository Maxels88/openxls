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

/**
 * Indicates an error occurred during formula calculation.
 */
public class CalculationException extends Exception
{
	private static final long serialVersionUID = 2028428287133817627L;
	static String[][] errorStrings = {
			{ "#DIV/0!", "7" },
			{ "#N/A", "42" },
			{ "#NAME?", "29" },
			{ "#NULL!", "0" },
			{ "#NUM!", "36" },
			{ "#REF!", "23" },
			{ "#VALUE!", "15" },
			{ "#CIR_ERR!", "15" }
			// false error with code for value, output for circular ref exceptions
	};

	/**
	 * Excel #NULL! error.
	 * Indicates that a range intersection returned no cells.
	 */
	public static final byte NULL = (byte) 0x00;

	/**
	 * Excel #DIV/0! error.
	 * Indicates that the formula attempted to divide by zero.
	 */
	public static final byte DIV0 = (byte) 0x07;

	/**
	 * Excel #VALUE! error.
	 * Indicates that there was an operand type mismatch.
	 */
	public static final byte VALUE = (byte) 0x0F;

	/**
	 * Excel #REF! error.
	 * Indicates that a reference was made to a cell that doesn't exist.
	 */
	public static final byte REF = (byte) 0x17;

	/**
	 * Excel #NAME? error.
	 * Indicates an unknown string was encountered in the formula.
	 */
	public static final byte NAME = (byte) 0x1D;

	/**
	 * Excel #NUM! error.
	 * Indicates that a calculation result overflowed the number storage.
	 */
	public static final byte NUM = (byte) 0x24;

	/**
	 * Excel #N/A error.
	 * Indicates that a lookup (e.g. VLOOKUP) returned no results.
	 */
	public static final byte NA = (byte) 0x2A;

	/**
	 * Custom circular exception error, internally stores as a #VALUE
	 */
	public static final byte CIR_ERR = (byte) 0xFF;

	/**
	 * The error code for this error.
	 */
	private final byte error;

	/**
	 * Creates a new CaluculationException.
	 *
	 * @param error the error code. must be one of the defined error constants.
	 */
	public CalculationException( byte error )
	{
		this.error = error;
	}

	/**
	 * Gets the BIFF8 error code for this error.
	 */
	public byte getErrorCode()
	{
		if( error == CIR_ERR )
		{
			return VALUE;
		}
		return error;
	}

	/**
	 * static version, takes String error code and returns the correct error code
	 *
	 * @param error String
	 * @return
	 */
	public static byte getErrorCode( String error )
	{
		if( error == null )
		{
			return 0;    // unknown
		}
		for( int i = 0; i < errorStrings.length; i++ )
		{
			if( error.equals( errorStrings[i][0] ) )
			{
				return new Byte( errorStrings[i][1] ).byteValue();
			}
		}
		return 0;
	}

	/**
	 * Gets a human-readable message describing this error.
	 */
	@Override
	public String getMessage()
	{
		switch( error )
		{
			case NULL:
				return "a range intersection returned no cells";
			case DIV0:
				return "attempted to divide by zero";
			case VALUE:
				return "operand type mismatch";
			case REF:
				return "reference to a cell that doesn't exist";
			case NAME:
				return "reference to an unknown function or defined name";
			case NUM:
				return "number storage overflow";
			case NA:
				return "lookup returned no value for the given criteria";
			case CIR_ERR:
				return "circular reference error";
			default:
				return "unknown error occurred";
		}
	}

	/**
	 * Gets the string name of this error.
	 */
	public String getName()
	{
		switch( error )
		{
			case NULL:
				return "#NULL!";
			case DIV0:
				return "#DIV/0!";
			case VALUE:
				return "#VALUE!";
			case REF:
				return "#REF!";
			case NAME:
				return "#NAME?";
			case NUM:
				return "#NUM!";
			case NA:
				return "#N/A";
			case CIR_ERR:
				return "#CIR_ERR!";
			default:
				return null;
		}
	}

	public String toString()
	{
		return getName();
	}
}
