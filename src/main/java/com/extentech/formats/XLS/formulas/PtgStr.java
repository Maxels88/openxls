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

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;

/**
 * PTG that stores a unicode string
 * <p/>
 * Offset  Name       Size     Contents
 * ------------------------------------
 * 0       cch          1      Length of the string
 * 1       rgch         var    The string
 * <p/>
 * *  I think the string includes a grbit itself, see UnicodeString.  Internationalization issues
 * may exist here!!!
 * <p/>
 * -- Yes, it did include grbit, all handled now.
 *
 * @see Ptg
 * @see Formula
 */
public class PtgStr extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1427051673654768400L;

	public boolean getIsOperand()
	{
		return true;
	}

	short cch;
	byte grbit;
	boolean negativeCch = false;

	public String getString()
	{
		String strVal = null;
		try
		{
			if( (grbit & 0x1) == 0x1 )
			{    // hits on Japanese strings in formulas
				byte[] barr = new byte[cch * 2];
				System.arraycopy( record, 3, barr, 0, cch * 2 );
				strVal = new String( barr, UNICODEENCODING );
			}
			else
			{
				byte[] barr = new byte[cch];
				System.arraycopy( record, 3, barr, 0, cch );
				strVal = new String( barr, DEFAULTENCODING );
			}
		}
		catch( Exception e )
		{
			byte[] barr = new byte[cch];
			System.arraycopy( record, 3, barr, 0, cch );
			strVal = new String( barr );
		}
		return strVal;
	}

	public String toString()
	{
		return getString();
	}

	/**
	 * return the human-readable String representation of
	 */
	public String getTextString()
	{
		try
		{
			Double d = new Double( getString() );
		}
		catch( NumberFormatException e )
		{
		}
		return "\"" + getString() + "\"";
	}

	public Object getValue()
	{
		return getString();
	}

	public PtgStr()
	{
		// default constructor
	}

	public PtgStr( String s )
	{
		ptgId = 0x17;
		setVal( s );
	}

	public void init( byte[] b )
	{
		grbit = b[2];
		cch = (short) (b[1] & 0xff); // this is the cch
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	/**
	 * Constructer to create these on the fly, this is needed
	 * for value storage in calculations of formulas.
	 */
	private void populateVals()
	{
		// no longer does anything, no String value stored
	}

	public String getVal()
	{
		return getString();
	}

	private String tempstr = null;

	public void setVal( String s )
	{
		tempstr = s;
		this.updateRecord();
	}

	public void updateRecord()
	{
		String ts = tempstr;
		if( ts == null )
		{
			return;
		}

		if( ByteTools.isUnicode( ts ) )
		{
			grbit = (byte) (grbit | 0x1);
		}
		try
		{
			byte[] strbytes = null;
			if( (grbit & 0x1) == 0x1 )
			{
				strbytes = ts.getBytes( UNICODEENCODING );
			}
			else
			{
				strbytes = ts.getBytes( DEFAULTENCODING );
			}

			short strbytelen = (short) strbytes.length;
			cch = strbytelen;
			if( (grbit & 0x1) == 0x1 )
			{
				cch = (short) (strbytelen / 2);
			}
			//cch = (short)( getString().length() + 3);
			record = new byte[strbytelen + 3];
			//record = new byte[(cch*times) + 3];
			record[0] = 0x17;
			record[1] = (byte) cch;
			record[2] = grbit;
			System.arraycopy( strbytes, 0, record, 3, strbytelen );

		}
		catch( UnsupportedEncodingException e )
		{
			Logger.logInfo( "decoding formula string failed: " + e );
		}
	}

	public int getLength()
	{
		return record.length;
	}

}