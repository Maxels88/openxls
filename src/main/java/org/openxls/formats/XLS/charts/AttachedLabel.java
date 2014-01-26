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

package org.openxls.formats.XLS.charts;

import org.openxls.formats.XLS.XLSRecord;
import org.openxls.toolkit.ByteTools;

/**
 * <b> AttachedLabel: Series Data/Value Labels (0x100c) </b>
 * <p/>
 * <p/>
 * bit
 * 0	= show value label -- a bit that specifies whether the value, or the vertical value on bubble or scatter chart groups, is displayed in the data label.
 * 1   = show value as percentage -- (must be 0 for non-pie)
 * 2	= show cat/label as percentage -- A bit that specifies whether the category (3) name and value,
 * represented as a percentage of the sum of the values of the series with
 * which the data label is associated, are displayed in the data label.	(pie charts only)
 * 3	= unused
 * 4	= show cat or series label -- A bit that specifies whether the category, or the horizontal value on bubble or scatter chart groups,
 * is displayed in the data label on a non-area chart group, or the series name is displayed in the data label on an area chart group.
 * 5	= show bubble label	-- A bit that specifies whether the bubble size is displayed in the data label.
 * 6   = show series name -- A bit that specifies whether the data label contains the name of the series.
 * <p/>
 * ???? Series DataLabels do not appear to be done via AttachedLabel
 */
public class AttachedLabel extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2532517522176536995L;
	public static final int VALUE = 0x1;
	public static final int VALUEPERCENT = 0x2;
	public static final int CATEGORYPERCENT = 0x4;
	public static final int CATEGORYLABEL = 0x10;
	public static final int BUBBLELABEL = 0x20;
	public static final int VALUELABEL = 0x40;

	private short grbit = 0;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
	}

	public static XLSRecord getPrototype()
	{
		AttachedLabel al = new AttachedLabel();
		al.setOpcode( ATTACHEDLABEL );
		al.setData( al.PROTOTYPE_BYTES );
		al.init();
		return al;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	/**
	 * return a string of all the label options chosen
	 *
	 * @return
	 */
	public String getType()
	{
		String ret = "";
		if( (grbit & VALUE) == VALUE )        // bit 0
		{
			ret = "Value ";
		}
		if( (grbit & VALUEPERCENT) == VALUEPERCENT )        // bit 1
		{
			ret = ret + "ValuePerecentage ";
		}
		if( (grbit & CATEGORYPERCENT) == CATEGORYPERCENT )        // bit 2
		{
			ret = ret + "CategoryPercentage ";    // Pie only
		}
		if( (grbit & CATEGORYLABEL) == CATEGORYLABEL )        // bit 4
		{
			ret = ret + "CategoryLabel ";
		}
		if( (grbit & BUBBLELABEL) == BUBBLELABEL )        // bit 5
		{
			ret = ret + "BubbleLabel ";
		}
		if( (grbit & VALUELABEL) == VALUELABEL )        // bit 6
		{
			ret = ret + "SeriesLabel ";
		}
		return ret.trim();
	}

	/**
	 * return the data label options as an int
	 *
	 * @return a combination of data label options above or 0 if none
	 * @see AttachedLabel constants
	 * <br>SHOWVALUE= 0x1;
	 * <br>SHOWVALUEPERCENT= 0x2;
	 * <br>SHOWCATEGORYPERCENT= 0x4;
	 * <br>SHOWCATEGORYLABEL= 0x10;
	 * <br>SHOWBUBBLELABEL= 0x20;
	 * <br>SHOWSERIESLABEL= 0x40;
	 */
	public int getTypeInt()
	{
		return grbit;
	}

	public void setType( short type )
	{
		grbit = type;
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * Show or Hide AttachedLabel Option
	 * <br>ShowValueLabel,
	 * <br>ShowValueAsPercent,
	 * <br>ShowLabelAsPercent,
	 * <br>ShowLabel,
	 * <br>ShowBubbleLabel
	 * <br>ShowSeriesName
	 *
	 * @param val true or 1 to set
	 */
	public void setType( String type, String val )
	{
		boolean bSet = val.equals( "true" ) || val.equals( "1" );
		if( type.equals( "ShowValueLabel" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 0 );
		}
		if( type.equals( "ShowValueAsPercent" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 1 );
		}
		if( type.equals( "ShowLabelAsPercent" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 2 );
		}
		if( type.equals( "ShowLabel" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 4 );
		}
		if( type.equals( "ShowBubbleLabel" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 5 );
		}
		if( type.equals( "ShowSeriesName" ) )
		{
			grbit = ByteTools.updateGrBit( grbit, bSet, 6 );
		}
		byte[] bb = ByteTools.shortToLEBytes( grbit );
		getData()[0] = bb[0];
		getData()[1] = bb[1];
	}

	/**
	 * return the value of the specified option
	 *
	 * @param type Sting option
	 * @return string true or false
	 */
	public String getType( String type )
	{
		boolean b = false;
		String ret = "";
		if( type.equals( "ShowValueLabel" ) )
		{
			b = ((grbit & VALUE) == VALUE);        // bit 0
		}
		if( type.equals( "ShowValueAsPercent" ) )
		{
			b = ((grbit & VALUEPERCENT) == VALUEPERCENT);
		}
		if( type.equals( "ShowLabelAsPercent" ) )
		{
			b = (((grbit & CATEGORYPERCENT) == CATEGORYPERCENT));
		}
		if( type.equals( "ShowLabel" ) )
		{
			b = (((grbit & CATEGORYLABEL) == CATEGORYLABEL));
		}
		if( type.equals( "ShowBubbleLabel" ) )
		{
			b = (((grbit & BUBBLELABEL) == BUBBLELABEL));
		}
		if( type.equals( "ShowSeriesName" ) )
		{
			b = (((grbit & VALUELABEL) == VALUELABEL));
		}
		return (b) ? ("1") : ("0");
	}

	/**
	 * @param type
	 * @deprecated
	 */
	public void setType( String type )
	{
		short t = 0;
		if( type.equalsIgnoreCase( "Value" ) || type.equalsIgnoreCase( "Y Value" ) )
		{
			t = 1;
		}
		else if( type.equalsIgnoreCase( "ValuePercentage" ) )
		{
			t = 2;
		}
		else if( type.equalsIgnoreCase( "CategoryPercentage" ) )
		{
			t = 3;
		}
		else if( type.equalsIgnoreCase( "Category" ) || type.equalsIgnoreCase( "X Value" ) )
		{
			t = 16;
		}
		else if( type.equalsIgnoreCase( "CandP" ) )
		{
			t = 22;
		}
		else if( type.equalsIgnoreCase( "Bubble" ) )
		{
			t = 32;
		}
		grbit = t;
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public String toString()
	{
		return getType();
	}

}
