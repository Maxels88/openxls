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
 * SXDI 0xC5
 * <p/>
 * The SXDI record specifies a data item for a PivotTable view.
 * <p/>
 * isxvdData (2 bytes): A signed integer that specifies a pivot field index as specified in Pivot Fields.
 * <p/>
 * If the PivotTable view is a non-OLAP PivotTable view, the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
 * <p/>
 * If the PivotTable view is an OLAP PivotTable view, the associated pivot hierarchy of the referenced pivot field specifies the OLAPmeasure for this data item and the iiftab field is ignored. See Association of Pivot Hierarchies and Pivot Fields and Cache Fields to determine the associated pivot hierarchy.
 * <p/>
 * MUST be greater than or equal to zero and less than the value of the cDim field of the preceding SxView record.
 * <p/>
 * The value of the sxaxis.sxaxisData field of the Sxvd record of the referenced pivot field MUST be 1.
 * <p/>
 * iiftab (2 bytes): A signed integer that specifies the aggregation function.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0x0000		    Sum of values
 * 0x0001		    Count of values
 * 0x0002		    Average of values
 * 0x0003		    Max of values
 * 0x0004		    Min of values
 * 0x0005		    Product of values
 * 0x0006		    Count of numbers
 * 0x0007		    Statistical standard deviation (sample)
 * 0x0008		    Statistical standard deviation (population)
 * 0x0009		    Statistical variance (sample)
 * 0x000A			Statistical variance (population)
 * <p/>
 * df (2 bytes): A signed integer that specifies the calculation used to display the value of this data item.
 * MUST be a value from the following table:
 * Value		    Meaning
 * 0x0000		    The data item value is displayed.
 * 0x0001		    Display as the difference between this data item value and the value of the pivot item specified by isxvi.
 * 0x0002		    Display as a percentage of the value of the pivot item specified by isxvi.
 * 0x0003		    Display as a percentage difference from the value of the pivot item specified by isxvi.
 * 0x0004		    Display as the running total for successive pivot items in the pivot field specified by isxvd.
 * 0x0005		    Display as a percentage of the total for the row containing this data item.
 * 0x0006		    Display as a percentage of the total for the column containing this data item.
 * 0x0007		    Display as a percentage of the grand total of the data item.
 * 0x0008		    Calculate the value to display using the following formula:
 * ((this data item value) * (grand total of grand totals)) / ((row grand total) * (column grand total))
 * <p/>
 * isxvd (2 bytes): A signed integer that specifies a pivot field index as specified in Pivot Fields.
 * The referenced pivot field is used in calculations as specified by the df field.
 * If df is 0x0001, 0x0002, 0x0003, or 0x0004 then the value of isxvd MUST be greater than or equal to zero and
 * less than the value of the cDim field in the preceding SxView record.
 * Otherwise, the value of isxvd is undefined and MUST be ignored.
 * <p/>
 * isxvi (2 bytes): A signed integer that specifies the pivot item used by df.
 * <p/>
 * If df is 0x0001, 0x0002, or 0x0003 then the value of this field MUST be a value from the following table:
 * Value		    Meaning
 * 0 to 0x7EFE	    A pivot item index, as specified by Pivot Items, that specifies a pivot item in the pivot field specified by isxvd.
 * MUST be less than the cItm field of the Sxvd record of the pivot field specified by isxvd.
 * 0x7FFB		    The previous pivot item in the pivot field specified by isxvd.
 * 0x7FFC		    The next pivot item in the pivot field specified by isxvd.
 * Otherwise, the value is undefined and MUST be ignored.
 * <p/>
 * ifmt (2 bytes): An IFmt structure that specifies the number format for this item.
 * <p/>
 * cchName (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stName field.
 * If the value is 0xFFFF then stName does not exist. Otherwise, the value MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * MUST NOT be 0xFFFF when the PivotCache functionality level is less than 3, or for non-OLAP PivotTable view .
 * <p/>
 * stName (variable): An XLUnicodeStringNoCch structure that specifies the name of this data item.
 * A value that is not NULL specifies that this string is used to override the name in the corresponding cache field.
 * <p/>
 * MUST NOT exist if cchName is 0xFFFF. Otherwise, MUST exist and the length MUST equal cchName.
 * <p/>
 * If this string is not NULL and the PivotTable view is a non-OLAP PivotTable view, this field MUST be unique within all SXDI records in this PivotTable view.
 */
public class SxDI extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	short isxvdData, isxvd, iiftab, df, isxvi, cchName, ifmt;
	String name = null;
	private static final long serialVersionUID = 2639291289806138985L;

	/**
	 * enum possible aggregation functions for Data axis fields
	 * <br>Sum, Counta, Average, Max, Min, Product, Count, StdDev, StdDevP, Var, VarP
	 */
	public enum AGGREGATIONFUNCTIONS
	{
		Sum( "sum" ),
		Count( "count" ),
		Average( "average" ),
		Max( "max" ),
		Min( "min" ),
		Product( "product" ),
		CountNums( "countnums" ),
		StdDev( "stdDev" ),
		StdDevP( "stdDevp" ),
		Var( "var" ),
		VarP( "varp" );
		private final String agf;

		AGGREGATIONFUNCTIONS( String s )
		{
			this.agf = s;
		}

		public static int get( String s )
		{
			for( AGGREGATIONFUNCTIONS c : values() )
			{
				if( c.agf.equals( s ) )
				{
					return c.ordinal();
				}
			}
			return 0;
		}
	}

	;

	/**
	 * enum possible display types for Data axis fields
	 * <li>value	-- The data item value is displayed.
	 * <li>difference --Display as the difference between this data item value and the value of the pivot item
	 * <li>percentageValue -- Display as a percentage of the value of the pivot item
	 * <li>percentageDifference	-- Display as a percentage difference from the value of the pivot item
	 * <li>runningTotal	-- Display as the running total for successive pivot items in the pivot field
	 * <li>percentageTotalRow	-- Display as a percentage of the total for the row containing this data item.
	 * <li>percentageTotalCol -- Display as a percentage of the total for the column containing this data item.
	 * <li>grandTotal -- Display as a percentage of the grand total of the data item.
	 * <li>calculated 	--Calculate the value to display using the following formula:
	 * <blockquote>((this data item value) * (grand total of grand totals)) / ((row grand total) * (column grand total))</blockquote>
	 */
	public enum DISPLAYTYPES
	{
		value, difference, percentageValue, percentageDifference, runningTotal, percentageTotalRow, percentageTotalCol, grandTotal, calculated
	}

	;

	public void init()
	{
		super.init();
		isxvdData = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );    // pivot field index
		iiftab = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );    // aggregation function -- see  aggregationfunctions
		df = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );        // display calculation
		isxvd = ByteTools.readShort( this.getByteAt( 6 ),
		                             this.getByteAt( 7 ) );    // specifies a pivot field index used in calculations as specified by the df field.
		isxvi = ByteTools.readShort( this.getByteAt( 8 ),
		                             this.getByteAt( 9 ) );    // A signed integer that specifies the pivot item used by df.
		ifmt = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );    // number format index
		cchName = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
		if( cchName != -1 )
		{
			byte encoding = this.getByteAt( 14 );
			byte[] tmp = this.getBytesAt( 15, (cchName) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					name = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					name = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding PivotTable caption name in Sxvd: " + e );
			}
		}
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXDI - isxvdData:" + isxvdData + " iiftab:" + iiftab + " df:" + df + " isxvd:" + isxvd + " isxvi:" + isxvi + " ifmt:" + ifmt + " name:" + name );
		}
	}

	/**
	 * returns the pivot field index for this data item;
	 * <br>the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
	 *
	 * @return
	 */
	public short getPivotFieldIndex()
	{
		return isxvdData;
	}

	/**
	 * sets the pivot field index for this data item;
	 * <br>the values in the source data associated with the associated cache field of the referenced pivot field are aggregated as specified in this record.
	 *
	 * @return
	 */
	public void setPivotFieldIndex( int fi )
	{
		isxvdData = (short) fi;
		byte[] b = ByteTools.shortToLEBytes( isxvdData );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	/**
	 * returns the aggregation function for this data field
	 *
	 * @return
	 * @see AGGREGATIONFUNCTIONS
	 */
	public int getAggregationFunction()
	{
		return iiftab;
	}

	/**
	 * sets the aggregation function for this data field
	 *
	 * @see AGGREGATIONFUNCTIONS
	 */
	public void setAggregationFunction( int af )
	{
		iiftab = (short) af;
		byte[] b = ByteTools.shortToLEBytes( iiftab );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
	}

	/**
	 * adds the data field to the DATA axis and it's aggregation function ("sum" is default)
	 *
	 * @param fieldIndex
	 * @param aggregationFunction
	 * @param name
	 */
	public void addDataField( int fieldIndex, String aggregationFunction, String name )
	{
		setPivotFieldIndex( fieldIndex );
		setAggregationFunction( AGGREGATIONFUNCTIONS.get( aggregationFunction ) );
		setName( name );
	}

	/**
	 * A signed integer that specifies the calculation used to display the value of this data item.
	 *
	 * @return
	 * @see DISPLAYTYPES
	 */
	public int getDisplayCalculation()
	{
		return df;
	}

	/**
	 * sets the Display Calculation for the Data item:  A signed integer that specifies the calculation used to display the value of this data item.
	 *
	 * @see DISPLAYTYPES
	 */
	public void setDisplayCalculation( int dc )
	{
		df = (short) dc;
		byte[] b = ByteTools.shortToLEBytes( df );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	/**
	 * specifies a pivot field index used in calculations as specified by the Display Calculation function
	 *
	 * @return
	 */
	public short getCalculationPivotFieldIndex()
	{
		return isxvd;
	}

	/**
	 * specifies a pivot field index used in calculations as specified by the Display Calculation function
	 *
	 * @return
	 */
	public void setCalculationPivotFieldIndex( int ci )
	{
		isxvd = (short) ci;
		byte[] b = ByteTools.shortToLEBytes( isxvd );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
	}

	/**
	 * specifies a pivot item index used in calculations as specified by the Display Calculation function
	 *
	 * @return
	 */
	public short getCalculationPivotItemIndex()
	{
		return isxvi;
	}

	/**
	 * specifies a pivot item index used in calculations as specified by the Display Calculation function
	 *
	 * @return
	 */
	public void setCalculationPivotItemIndex( int ci )
	{
		isxvi = (short) ci;
		byte[] b = ByteTools.shortToLEBytes( isxvi );
		this.getData()[8] = b[0];
		this.getData()[9] = b[1];
	}

	/**
	 * returns the index to the number format pattern for this data field
	 *
	 * @return
	 */
	public short getNumberFormat()
	{
		return ifmt;
	}

	/**
	 * sets the index to the number format pattern for this data field
	 */
	public void setNumberFormat( int i )
	{
		ifmt = (short) i;
		byte[] b = ByteTools.shortToLEBytes( ifmt );
		this.getData()[10] = b[0];
		this.getData()[11] = b[1];
	}

	/**
	 * returns the name of this data item; if not null, overrides the name in the corresponding cache field
	 *
	 * @return
	 */
	public String getName()
	{
		return name;
	}

	/**
	 * sets the name of this data item; if not null, overrides the name in the corresponding cache field
	 * for this pivot item
	 */
	public void setName( String name )
	{
		this.name = name;
		byte[] data = new byte[14];
		System.arraycopy( this.getData(), 0, data, 0, 13 );
		if( name != null )
		{
			byte[] strbytes = null;
			try
			{
				strbytes = this.name.getBytes( DEFAULTENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding pivot table name in SXVI: " + e );
			}

			//update the lengths:
			cchName = (short) strbytes.length;
			byte[] nm = ByteTools.shortToLEBytes( cchName );
			data[12] = nm[0];
			data[13] = nm[1];

			// now append variable-length string data
			byte[] newrgch = new byte[cchName + 1];    // account for encoding bytes
			System.arraycopy( strbytes, 0, newrgch, 1, cchName );

			data = ByteTools.append( newrgch, data );
		}
		else
		{
			data[12] = -1;
			data[13] = -1;
		}
		this.setData( data );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 	/* isxvdData */
			0, 0,	/* iiftab */
			0, 0,   /* df */
			0, 0,	/* isxvd */
			0, 0,	/* isxvi */
			0, 0,	/* ifmt */
			-1, -1,	/* cchname */
	};

	public static XLSRecord getPrototype()
	{
		SxDI di = new SxDI();
		di.setOpcode( SXDI );
		di.setData( di.PROTOTYPE_BYTES );
		di.init();
		return di;
	}
}
