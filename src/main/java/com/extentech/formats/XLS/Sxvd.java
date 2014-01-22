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
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;

/**
 * SXVD b[1]h: View Fields
 * The Sxvd record specifies pivot field properties and specifies the
 * beginning of a collection of records.
 * This collection of records specifies details for a pivot field.
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       sxaxis      2       0x0 = no axis
 * 0x1 = row
 * 0x2 = col
 * 0x4 = page
 * 0x8 = data
 * 6       cSub        2       Number of subtotals attached
 * 8       grbitSub    2       Item subtotal type (see subtable)
 * 10      cItm        2       Number of items
 * 12      cchName     2       Length of the name if name= 0xFFFF, rgch is null and name in cache used.
 * 14      rgch        var     name
 * <p/>
 * name        contents
 * --------------------------------
 * bitFNone    0000
 * bitFDefault 0001
 * bitFSum     0002
 * bitFCounta  0004
 * bitFAvg     0008
 * bitFMax     0010
 * bitFMin     0020
 * <p/>
 * </p></pre>
 * <p/>
 * MORE INFO:
 * <p/>
 * sxaxis (2 bytes): An SXAxis structure that specifies the PivotTable axis that this pivot field is on.
 * If the sxaxis.sxaxisData field equals 1, there MUST be a corresponding SXDI record with an isxvd field that specifies this Sxvd record.
 * <p/>
 * cSub (2 bytes): An unsigned integer that specifies the number of subtotal functions used for this pivot field.
 * MUST equal the count of subtotal fields of this record whose value is 1.
 * The subtotal fields of this record are fDefault, fSum, fCounta, fAverage, fMax, fMin, fProduct, fCount, fStdev, fStdevp, fVariance, and fVariancep.
 * For more information, see Subtotalling.
 * <p/>
 * A - fDefault (1 bit): A bit that specifies whether the default subtotal function is applied.
 * The default subtotal is separately determined for each data item. If the fDefault field equals 1
 * and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage
 * field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 1.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The default subtotal function is not applied.
 * 1			    The default subtotal function is applied.
 * <p/>
 * B - fSum (1 bit): A bit that specifies whether the sum subtotal function is displayed. If the fDefault field equals 1,
 * this value MUST be zero. If the fSum field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol
 * field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the
 * SXVI record equal to 2. MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The sum subtotal function is not displayed.
 * 1				The sum subtotal function is displayed.
 * <p/>
 * C - fCounta (1 bit): A bit that specifies whether the count subtotal function is displayed. If the fDefault field equals 1,
 * this value MUST be zero. If the fCounta field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol
 * field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 3.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The count subtotal function is not displayed.
 * 1			    The count subtotal function is displayed.
 * <p/>
 * D - fAverage (1 bit): A bit that specifies whether the average subtotal function is displayed. If the fDefault field equals 1, this value MUST be zero.
 * If the fAverage field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage
 * field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 4.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0		    	The average subtotal function is not displayed.
 * 1			    The average subtotal function is displayed.
 * <p/>
 * E - fMax (1 bit): A bit that specifies whether the max subtotal function is displayed. If the fDefault field equals 1, this value MUST be zero.
 * If the fMax field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage
 * field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 5.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The max subtotal function is not displayed.
 * 1			    The max subtotal function is displayed.
 * <p/>
 * F - fMin (1 bit): A bit that specifies whether the min subtotal function is displayed. If the fDefault field equals 1, this value MUST be zero.
 * If the fMin field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage
 * field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 6.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The min subtotal function is not displayed.
 * 1			    The min subtotal function is displayed.
 * <p/>
 * G - fProduct (1 bit): A bit that specifies whether the product subtotal function is displayed. If the fDefault field is 1, this value MUST be zero.
 * If the fProduct field is 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage
 * field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 7.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0				The product subtotal function is not displayed.
 * 1			    The product subtotal function is displayed.
 * <p/>
 * H - fCount (1 bit): A bit that specifies whether the count numbers subtotal function is displayed. If the fDefault field is 1,
 * this value MUST be zero. If the fCount field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol
 * field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 8.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The count numbers subtotal function is not displayed.
 * 1			    The count numbers subtotal function is displayed.
 * <p/>
 * I - fStdev (1 bit): A bit that specifies whether the standard deviation subtotal function is displayed. If the fDefault field is 1,
 * this value MUST be zero. If the fStdev field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field
 * equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 9.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0				The standard deviation subtotal function is not displayed.
 * 1			    The standard deviation subtotal function is displayed.
 * <p/>
 * J - fStdevp (1 bit): A bit that specifies whether the standard deviation population subtotal function is displayed. If the fDefault field equals 1,
 * this value MUST be zero. If the fStdevp field equals 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field
 * equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 10.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The standard deviation population subtotal function is not displayed.
 * 1			    The standard deviation population subtotal function is displayed.
 * <p/>
 * K - fVariance (1 bit): A bit that specifies whether the variance subtotal function is displayed. If the fDefault field is 1,
 * this value MUST be zero. If the fVariance field is 1 and the sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol
 * field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 11.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0			    The variance subtotal function is not displayed.
 * 1			    The variance subtotal function is displayed.
 * <p/>
 * L - fVariancep (1 bit): A bit that specifies whether the variance population subtotal function is displayed. If the fDefault field is 1,
 * the value MUST be zero. If the fVariancep field equals 1 and sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field
 * equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 12.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0				The variance population subtotal function is not displayed.
 * 1			    The variance population subtotal function is displayed.
 * <p/>
 * M - reserved (4 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * cItm (2 bytes): A signed integer that specifies the number of pivot items for this pivot field. This value MUST match the number of
 * SXVI records following this record and MUST be less than or equal to the following formula:
 * 32500 + the cSub field
 * <p/>
 * cchName (2 bytes): An unsigned integer that specifies the length, in characters, of the stName field.
 * If the value is 0xFFFF then stName is NULL. The value MUST be 0xFFFF or greater than zero and less than or equal to 255.
 * <p/>
 * stName (variable): An XLUnicodeStringNoCch structure that specifies the caption of this pivot field.
 * A non-NULL value specifies that this string is used to override the stFieldName field in
 * SXFDB record from the associated cache field, as specified in pivot fields. The length is specified in cchName.
 * This field exists only if the value of cchName is not 0xFFFF. If this PivotTable view is not an OLAP PivotTable view
 * and this string is non-NULL, then stName MUST be unique within all Sxvd records in this PivotTable view.
 */
public class Sxvd extends XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6537376162863865578L;

	public static final short AXIS_NONE = 0;
	public static final short AXIS_ROW = 1;
	public static final short AXIS_COL = 2;
	public static final short AXIS_PAGE = 4;
	public static final short AXIS_DATA = 8;

	private short cSub = -1;
	private short cItm = -1;
	private short axis = -1;
	private short cchName = -1;
	private String caption = null;

	// flags:
	boolean fDefault, fSum, fCounta, fAverage, fMax, fMin, fProduct, fCount, fStdev, fStdevp, fVariance, fVariancep;

	@Override
	public void init()
	{
		super.init();
		axis = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		cSub = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		cItm = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		cchName = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		if( cchName != -1 )
		{
			byte encoding = this.getByteAt( 10 );
			byte[] tmp = this.getBytesAt( 11, (cchName) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					caption = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					caption = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding PivotTable caption name in Sxvd: " + e );
			}
		}
		// type of subtotal funtion
		byte b0 = this.getByteAt( 4 );
		byte b1 = this.getByteAt( 5 );
		fDefault = ((b0 & 0x1) == 0x1);    // default subtotal -- if 1, the rest of these flags are ignored
		fSum = ((b0 & 0x2) == 0x2);    //  sum subtotal function -if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 2.
		fCounta = ((b0 & 0x4) == 0x4);    // count subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 3.
		fAverage = ((b0 & 0x8) == 0x8);    // average subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 4.
		fMax = ((b0 & 0x10) == 0x10);    // max subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 5.
		fMin = ((b0 & 0x20) == 0x20);    // min subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 6.
		fProduct = ((b0 & 0x40) == 0x40);    // product subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 7.
		fCount = ((b0 & 0x80) == 0x80);    // count numbers function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 8.
		fStdev = ((b1 & 0x1) == 0x1);    // standard deviation function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 9.
		fStdevp = ((b1 & 0x2) == 0x2);    // standard deviation population function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 10.
		fVariance = ((b1 & 0x4) == 0x4);// variance subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 11.
		fVariancep = ((b1 & 0x8) == 0x8);// variance population subtotal function - if so: if sxaxis.sxaxisRw field equals 1 or if the sxaxis.sxaxisCol field equals 1 or if the sxaxis.sxaxisPage field equals 1, there MUST be one SXVI record with the itmType field of the SXVI record equal to 12.

		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXVD - axis:" + getAxisTypeName() + " cSub:" + cSub + " cItm:" + cItm + " default?" + fDefault + " sum?" + fSum + " caption:" + caption );
		}
	}

	/**
	 * returns the String representation of the subtotal function for the axis defined in this record
	 *
	 * @return
	 */
	public String getSubTotalFunction()
	{
		if( fDefault )
		{
			return "Default";
		}
		if( fSum )
		{
			return "Sum";
		}
		if( fCounta )
		{
			return "Count";
		}
		if( fAverage )
		{
			return "Average";
		}
		if( fMax )
		{
			return "Max";
		}
		if( fMin )
		{
			return "Min";
		}
		if( fProduct )
		{
			return "Product";
		}
		//if (fCount)
		if( fStdev )
		{
			return "Stddev";
		}
		if( fStdevp )
		{
			return "StddevP";
		}
		if( fVariance )
		{
			return "Variance";
		}
		if( fVariancep )
		{
			return "VarianceP";
		}
		return "Default";
	}

	/**
	 * Sets the subtotal function(s) for this pivot field.
	 * <li>"Default";
	 * <li>"Sum";
	 * <li>"Count";
	 * <li>"Average";
	 * <li>"Max";
	 * <li>"Min";
	 * <li>"Product";
	 * <li>"Stddev";
	 * <li>"StddevP";
	 * <li>"Variance";
	 * <li>"VarianceP";
	 *
	 * @param f
	 * @return
	 */
	public void setSubTotalFunction( String[] f )
	{
		if( f == null )
		{
			f = new String[]{ "Default" };
		}
		cSub = (short) f.length;
		for( int i = 0; i < cSub; i++ )
		{
			fDefault = f[i].equalsIgnoreCase( "Default" );
			fSum = f[i].equalsIgnoreCase( "Sum" );
			fCounta = f[i].equalsIgnoreCase( "Count" );
			fAverage = f[i].equalsIgnoreCase( "Average" );
			fMax = f[i].equalsIgnoreCase( "Max" );
			fMin = f[i].equalsIgnoreCase( "Min" );
			fProduct = f[i].equalsIgnoreCase( "Product" );
			//if (fCount)
			fStdev = f[i].equalsIgnoreCase( "Stddev" );
			fStdevp = f[i].equalsIgnoreCase( "StddevP" );
			fVariance = f[i].equalsIgnoreCase( "Variance" );
			fVariancep = f[i].equalsIgnoreCase( "VarianceP" );
		}
		if( fDefault )
		{
			cSub = 1;    // is this correct?
		}
		// update record
		byte[] b = ByteTools.shortToLEBytes( cSub );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];

		b[0] = 0;
		b[1] = 0;
		if( fDefault )
		{
			b[0] |= 0x1;
		}
		if( fSum )
		{
			b[0] |= 0x2;
		}
		if( fCounta )
		{
			b[0] |= 0x4;
		}
		if( fAverage )
		{
			b[0] |= 0x8;
		}
		if( fMax )
		{
			b[0] |= 0x10;
		}
		if( fMin )
		{
			b[0] |= 0x20;
		}
		if( fProduct )
		{
			b[0] |= 0x40;
		}
		if( fCount )
		{
			b[0] |= (byte) 0x80;
		}
		if( fStdev )
		{
			b[1] |= 0x1;
		}
		if( fStdevp )
		{
			b[1] |= 0x2;
		}
		if( fVariance )
		{
			b[1] |= 0x4;
		}
		if( fVariancep )
		{
			b[1] |= 0x8;
		}

		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	/**
	 * returns the axis type int
	 *
	 * @return
	 */
	int getAxisType()
	{
		return axis;
	}

	/**
	 * returns the Pivot Table Axis this Sxvd record defines
	 *
	 * @return
	 * @see AXIS_ constants
	 */
	String getAxisTypeName()
	{
		String ret = "";
		switch( axis )
		{
			case AXIS_NONE:
				ret = "NONE";
				break;
			case Sxvd.AXIS_ROW:
				ret = "ROW";
				break;
			case Sxvd.AXIS_COL:
				ret = "COL";
				break;
			case Sxvd.AXIS_PAGE:
				ret = "PAGE";
				break;
			case Sxvd.AXIS_DATA:
				ret = "DATA";
				break;
		}
		return ret;
	}

	/**
	 * sets the axis for this pivot field
	 * <li>1= row
	 * <li>2= col
	 * <li>4= page
	 * <li>8= data
	 *
	 * @param axis
	 * @see Sxvd.AXIS_ constants
	 */
	public void setAxis( int axis )
	{
		if( !(axis == 1 || axis == 2 || axis == 4 || axis == 8) )
		{
			return;
		}
		this.axis = (short) axis;
		byte[] b = ByteTools.shortToLEBytes( this.axis );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		// TODO if axis is data MUST have a SxDI otherwise if axis WAS a data axis must remove SxDI
	}

	/**
	 * returns the number of pivot items for this pivot field
	 *
	 * @return
	 */
	int getNumItems()
	{
		return cItm;
	}

	public void setNumItems( int n )
	{
		cItm = (short) n;
		byte[] b = ByteTools.shortToLEBytes( cItm );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
	}

	/**
	 * returns the number of subtotal functions set for this pivot field
	 *
	 * @return
	 */
	int getNumSubtotals()
	{
		return cSub;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 	/* axis */
			1, 0,	/* cSub */
			1, 0,   /* flags */
			0, 0,	/* cItm */
			-1, -1,	/* cchName */
	};

	public static XLSRecord getPrototype()
	{
		Sxvd sv = new Sxvd();
		sv.setOpcode( SXVD );
		sv.setData( sv.PROTOTYPE_BYTES );
		sv.init();
		return sv;
	}

	public String toString()
	{
		return "SXVD - axis:" + getAxisTypeName() + " cSub:" + cSub + " cItm:" + cItm + " default?" + fDefault + " sum?" + fSum + " caption:" + caption;
	}

}