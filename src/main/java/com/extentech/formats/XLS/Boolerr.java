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
package com.extentech.formats.XLS;

import com.extentech.toolkit.ByteTools;

/**
 * <b>Boolerr: BiffRec Value, Boolean or Error (0x205)</b><br>
 * Describes a cell that contains a constant Boolean or error
 * value.
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row
 * 6       col         2       Column
 * 8       ixfe        2       Index to XF record
 * 10      bBoolErr    1       Boolean value or error value
 * 11      fError      1       Boolean/error flag (0 = Boolean, 1 = Error)
 * <p/>
 * Boolean vals are 1 = true, 0 = false.
 * <p/>
 * </p></pre>
 */

public final class Boolerr extends XLSCellRecord
{

	private static final long serialVersionUID = 39663492256953223L;
	private boolean val, iserr = false;
	private byte errorval;

	/**
	 * Returns whether the value is a Boolean or
	 * an error.
	 */
	boolean getIsErr()
	{
		return iserr;
	}

	/**
	 * get the int val
	 */
	@Override
	public int getIntVal()
	{
		if( this.getBooleanVal() )
		{
			return 1;
		}
		return 0;
	}

	/**
	 * return boolean value in float version 0 or 1
	 */
	@Override
	public float getFloatVal()
	{
		if( this.getBooleanVal() )
		{
			return 1;
		}
		return 0;
	}

	/**
	 * return the boolean value in double version 0 or 1
	 */
	@Override
	public double getDblVal()
	{
		if( this.getBooleanVal() )
		{
			return 1;
		}
		return 0;
	}

	/**
	 * get the String val
	 */
	@Override
	public String getStringVal( String encoding )
	{
		if( this.getIsErr() )
		{
			return this.getErrorCode();
		}
		return String.valueOf( val );
	}

	/**
	 * Returns the valid error code for this valrec
	 */
	public String getErrorCode()
	{
		if( !this.getIsErr() )
		{
			return String.valueOf( iserr );
		}
		String retval = "";
		if( (errorval & 0x0) == 0x0 )
		{
			retval = "#NULL!";
		}
		if( (errorval & 0x7) == 0x7 )
		{
			retval = "#DIV/0!";
		}
		if( (errorval & 0xF) == 0xF )
		{
			retval = "#VALUE!";
		}
		if( (errorval & 0x17) == 0x17 )
		{
			retval = "#REF!";
		}
		if( (errorval & 0x1D) == 0x1D )
		{
			retval = "#NAME?";
		}
		if( (errorval & 0x24) == 0x24 )
		{
			retval = "#NUM!";
		}
		if( (errorval & 0x2A) == 0x2A )
		{
			retval = "#N/A";
		}
		return retval;
	}

	/**
	 * get the String val
	 */
	@Override
	public String getStringVal()
	{
		return this.getStringVal( null ); // char encoding of true/false should be irrelevant -
		//! NOPE, 'cause this also has error codes in it -NR 10/03
	}

	/**
	 * Get the value of the record as a Boolean.
	 * Value must be parseable as a Boolean.
	 */
	@Override
	public boolean getBooleanVal()
	{
		return val;
	}

	@Override
	public void setBooleanVal( boolean newv )
	{
		if( newv )
		{
			this.getData()[6] = 1;
		}
		else
		{
			this.getData()[6] = 0;
		}
		this.val = newv;
	}

	@Override
	public void init()
	{
		super.init();
		// get the row information
		rw = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		;
		col = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		;
		ixfe = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		// get the value
		int num = this.getByteAt( 6 );
		if( num == 0 )
		{
			val = false;
		}
		if( num == 1 )
		{
			val = true;
		}
		num = this.getByteAt( 7 );
		if( num == 0 )
		{
			iserr = false;
		}
		if( num == 1 )
		{
			iserr = true;
		}
		errorval = this.getByteAt( 6 );
		this.setIsValueForCell( true );
		if( !iserr )
		{
			isBoolean = true;
		}
		else
		{
			isString = true;
		}
	}

	// these bytes are from a simple chart, 2 ranges a1:a2, b1:b2 - all default.  Likely will need to be modified
	// when we figure out wtf.
	private byte[] PROTOTYPE_BYTES = { 0, 0, 0, 0, 21, 0, 0, 0 };

	protected static XLSRecord getPrototype()
	{
		Boolerr be = new Boolerr();
		be.setOpcode( BOOLERR );
		be.setData( be.PROTOTYPE_BYTES );
		be.init();
		return be;
	}
}