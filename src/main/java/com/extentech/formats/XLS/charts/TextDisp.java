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
package com.extentech.formats.XLS.charts;

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

import java.util.HashMap;

/**
 * <b>TextDisp: (Text) Defines Display of Text Fields (0x1025)</b>
 * <p/>
 * Used in conjunction with several other records to define alignment, color, position,
 * size, and so on, of text fiedls that appear on the chart.
 * The fields in this record have meaning according to the TEXT record's parent
 * (CHART, LEGEND or DEFAULTTEXT)
 * <p/>
 * 4		at		1		Horizontal alignment (1=left, 2= center, 3=bottom, 4= justify, 7=distributed)
 * 5		vat		1		Vertical alignment (1= top, 2= center, 3= bottom, 4= justify, 7=distributed)
 * 6		wBkgMode 2		Display mode of bg 1= transparent, 2= opaque
 * 8		rgbText	4		Color of the text
 * LongRGB Structure:      red (1 byte): An unsigned integer that specifies the relative intensity of red.
 * green (1 byte): An unsigned integer that specifies the relative intensity of green.
 * blue (1 byte): An unsigned integer that specifies the relative intensity of blue.
 * reserved (1 byte): MUST be zero, and MUST be ignored.
 * 12		x		4		in SPRC (see Pos); ignore if preceded by DefaultText OR when followed by Pos
 * 16		y		4		in SPRC (see Pos);	""
 * 20		dx		4		A signed integer that specifies the horizontal size of the text, relative to the chart area in SPRC.
 * This value MUST be ignored when this record is followed by a Pos record;
 * otherwise MUST be greater than or equal to 0 and less than or equal to 32767. SHOULD be less than or equal to 4000.
 * 24		dy		4		A signed integer that specifies the vertical size of the text, relative to the chart area in SPRC.
 * This value MUST be ignored when this record is followed by a Pos record;
 * otherwise MUST be greater than or equal to 0 and less than or equal to 32767. SHOULD<128> be less than or equal to 4000.
 * 28		grbit	2		Option flags
 * 30		icvText 2		Index to color value. icv structure.
 * 32		dlp		2		(4 bits): An unsigned integer that specifies the data label positioning of the text, relative to the graph object item the text is attached to.
 * For all data label text fields, MUST be a value from the following table:
 * Data label position		Value		Value for chart group type
 * Auto					0x0			Pie chart group
 * Right					0x0			Line, Bubble, or Scatter chart group
 * Outside					0x0			Bar or Column chart group with fStacked equal to 0
 * Center					0x0			Bar or Column chart group with fStacked equal to 1
 * Outside End				0x1			Bar, Column, or Pie chart group
 * Inside End				0x2			Bar, Column, or Pie chart group
 * Center					0x3			Bar, Column, Line, Bubble, Scatter, or Pie chart group
 * Inside Base				0x4			Bar or Column chart group
 * Above					0x5			Line, Bubble, or Scatter chart group
 * <p/>
 * Below					0x6			Line, Bubble, or Scatter chart group
 * Left					0x7			Line, Bubble, or Scatter chart group
 * Right					0x8			Line, Bubble, or Scatter chart group
 * Auto					0x9			Pie chart group
 * Moved by user			0xA			All
 * unused 10 bits
 * iReadingOrder 2 bits
 * 34		trot			An unsigned integer that specifies the text rotation. MUST be a value from the following table:
 * Value		Angle description
 * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
 * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
 * 255			Text top-to-bottom with letters upright
 * <p/>
 * grbit
 * 0		0x1		fAutoColor	1= auto color, 0= user-selected
 * 1		0x2		fShowKey	1= the text is attached to a legend key.
 * 2		0x4		fShowValue	1= the value, or the vertical value on bubble or scatter chart groups, is displayed in the data label.
 * If the current attached label contains a DataLabExtContents record and the fPercent field of the DataLabExtContents record equals 0, this field MUST equal the fValue field of the DataLabExtContents record.
 * If the current attached label does not contain a DataLabExtContents record and fShowLabelAndPerc equals 1, this field MUST equal 0.
 * This field MUST equal 0 if the current attached label does not contain a DataLabExtContents record and one or more of the following conditions are satisfied:
 * The fShowLabelAndPerc field equals 1.
 * The fShowPercent field equals 1.
 * 3		0x8	fVert		unused
 * 4		0x10	fAutoText	1= autogenerated text string
 * 5		0x20	fGenerated	1= default 0= modified
 * 6		0x40	zfDeleted	1= an automatic label has been deleted by user
 * 7		0x80	fAutoMode	1= Bg is set to auto
 * 10-8		0x700	unused
 * 11		0x800	fShowLblPct	1= show category name and the value, represented as a percentage of the sum of the values of the series the data label is associated with,
 * MUST equal 0 if the chart group type of the corresponding chart group, series, or data point, is not a bar of pie, doughnut, pie, or pie of pie chart group.
 * This field MUST equal 1 if the current attached label contains a DataLabExtContents record and both of the following conditions are satisfied:
 * The fCatName and fPercent fields of the DataLabExtContents record equal 1.
 * The fSerName, fValue, and fBubSizes fields of the DataLabExtContents record equal 0.
 * This field MUST equal 0 if the current attached label contains a DataLabExtContents record and one or more of the following conditions is satisfied:
 * The fCatName or fPercent fields of the DataLabExtContents record equal 0.
 * The fSerName, fValue, or fBubSizes fields of the DataLabExtContents record equal 1.
 * MUST be ignored if fAutoText equals 0.
 * 12		0x1000	fShowPct	A bit that specifies whether the value, represented as a percentage of the sum of the values of the series the data label is associated with, is displayed in the data label.
 * MUST equal 0 if the chart group type of the corresponding chart group, series, or data point is not a bar of pie, doughnut, pie, or pie of pie chart group.
 * If the current attached label contains a DataLabExtContents record, this field MUST equal the value of the fPercent field of the DataLabExtContents record.
 * If the current attached label does not contain a DataLabExtContents record and fShowLabelAndPerc equals 1, this field MUST equal 1.
 * MUST be ignored if fAutoText equals 0.
 * 13		0x2000	fShowBubbleSizes	1= show bubble sizes
 * 14		0x4000	fShowCatLabel	bit that specifies whether the category (3), or the horizontal value on bubble or scatter chart groups, is displayed in the data label on a non-area chart group, or the series name is displayed in the data label on an area chart group.
 * This field MUST equal the fCatNameLabel field of the DataLabExtContents record if the current attached label contains a DataLabExtContents record, the chart group is non-area, and both of the following conditions are satisfied:
 * The fValue field of the DataLabExtContents record equals 0.
 * The fShowLabelAndPerc field equals 1 or the fPercent field equals 0.
 * This field MUST equal the fCatNameLabel field of the DataLabExtContents record if the current attached label contains a DataLabExtContents record, the chart group is area or filled radar, and the following condition is satisfied:
 * The fValue field of the DataLabExtContents record equals 0.
 * If the current attached label contains a DataLabExtContents record and the fValue field of the DataLabExtContents record equals 1, this field MUST equal 0.
 * This field MUST equal 0 if the current attached label does not contain a DataLabExtContents record and one of the following conditions is satisfied:
 * The fShowValue field equals 1.
 * The fShowLabelAndPerc field equals 0 and the fShowPercent field equals 1.
 * MUST be ignored if fAutoText equals 0.
 */
public class TextDisp extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7862828186455339066L;
	private short grbit = 0, grbit2 = 0;
	private short at = 2, vat = 1, wBkgMode = 1, icvText = 0, trot = 0, dlp;
	private int x = 0, y = 0, dx = 0, dy = 0;
	private java.awt.Color rgbText = null;
	private boolean fAutoColor, fShowKey, fShowValue, fVert, fAutoText, fGenerated, fDeleted, fAutoMode;
	private boolean fShowLblPct, fShowPct, fShowBubbleSizes, fShowCatLabel;

	@Override
	public void init()
	{
		super.init();
		byte[] data = getData();
		at = data[0];
		vat = data[1];
		wBkgMode = ByteTools.readShort( data[2], data[3] );
		rgbText = new java.awt.Color( ((data[4] < 0) ? (255 + data[4]) : data[4]),
		                              ((data[5] < 0) ? (255 + data[5]) : data[5]),
		                              ((data[6] < 0) ? (255 + data[6]) : data[6]) );
//		rgbText= ByteTools.readInt(this.getBytesAt(4, 4));
		x = ByteTools.readInt( this.getBytesAt( 8, 4 ) );
		y = ByteTools.readInt( this.getBytesAt( 12, 4 ) );
		dx = ByteTools.readInt( this.getBytesAt( 16, 4 ) );
		dy = ByteTools.readInt( this.getBytesAt( 20, 4 ) );
		grbit = ByteTools.readShort( data[24], data[25] );
		fAutoColor = (grbit & 0x1) == 0x1;    // usually T
		fShowKey = (grbit & 0x2) == 0x2;
		fShowValue = (grbit & 0x4) == 0x4;                    // pie-specific shows value label
		fVert = (grbit & 0x8) == 0x8;
		fAutoText = (grbit & 0x10) == 0x10;    // usually T except for set labels
		fGenerated = (grbit & 0x20) == 0x20;    // usually T
		fDeleted = (grbit & 0x40) == 0x40;
		fAutoMode = (grbit & 0x80) == 0x80;    // usually T
		fShowLblPct = (grbit & 0x800) == 0x800;            // pie-specific
		fShowPct = (grbit & 0x1000) == 0x1000;            // pie-specific
		fShowBubbleSizes = (grbit & 0x2000) == 0x2000;        // bubble-specific
		fShowCatLabel = (grbit & 0x4000) == 0x4000;        // pie-specific shows category label
		icvText = ByteTools.readShort( data[26], data[27] );
		dlp = (short) (data[28] & 0xF);
		grbit2 = ByteTools.readShort( data[28], data[29] );
		trot = ByteTools.readShort( data[30], data[31] );
	}

	/**
	 * @return the type of this TextDisp (X, Y, Z axis, title, data series ...
	 */
	public int getType()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == OBJECTLINK )
			{
				return ((ObjectLink) b).getType();
			}
		}
		return -1;
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		sb.append( " Label=\"" + this.toString() + "\"" );
		sb.append( " TextRotation=\"" + trot + "\"" );
		// KSC: TODO: Handle more options
		if( fShowLblPct )
		{
			sb.append( " ShowLabelPct=\"" + true + "\"" );
		}
		if( fShowPct )
		{
			sb.append( " ShowPct=\"" + true + "\"" );
		}
		if( fShowCatLabel )
		{
			sb.append( " ShowCatLabel=\"" + true + "\"" );    // 20081223 KSC: try to figure out difference between AttachedLabel and TextDisp options
		}
		if( fShowBubbleSizes )
		{
			sb.append( " ShowBubbleSizes=\"" + true + "\"" );
		}
		sb.append( " Font=\"" + getFontId() + "\"" );
		return sb.toString();
	}

	/**
	 * return the id of the Fontx associated with this TextDisp
	 *
	 * @return FontX id
	 */
	public int getFontId()
	{
		int ret = 0;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == FONTX )
			{
				ret = ((Fontx) b).getIfnt();
				break;
			}
		}
		return ret;
	}

	/**
	 * set the id of the Fontx associated with this TextDisp
	 */
	public void setFontId( int id )
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == FONTX )
			{
				((Fontx) b).setIfnt( id );
				break;
			}
		}
	}

	/**
	 * return the font record that sets the font properties for this text element
	 *
	 * @param wb
	 * @return
	 */
	public com.extentech.formats.XLS.Font getFont( WorkBook wb )
	{
		return wb.getFont( getFontId() );
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "Label" ) )
		{
			this.setText( val );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowKey" ) )
		{
			fShowKey = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowKey, 1 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowValue" ) )
		{
			fShowValue = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowValue, 2 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowLabelPct" ) )
		{
			fShowLblPct = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowLblPct, 11 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowPct" ) )
		{
			fShowPct = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowPct, 12 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowCatLabel" ) )
		{    // 20081223 KSC: changed from ShowLabel
			fShowCatLabel = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowCatLabel, 14 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowBubbleSizes" ) )
		{
			fShowBubbleSizes = val.equals( "true" ) || val.equals( "1" );
			grbit = ByteTools.updateGrBit( grbit, fShowBubbleSizes, 13 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "TextRotation" ) )
		{
			trot = Short.parseShort( val );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Font" ) )
		{
			setFontId( Short.parseShort( val ) );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

	/**
	 * Return the string value of the specified option
	 */
	@Override
	public String getChartOption( String op )
	{
		if( op.equalsIgnoreCase( "Label" ) )
		{
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b.getOpcode() == SERIESTEXT )
				{
					SeriesText st = (SeriesText) b;
					return st.toString();
				}
			}
		}
		else if( op.equalsIgnoreCase( "ShowKey" ) )
		{
			return ((fShowKey) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "ShowValue" ) )
		{
			return ((fShowValue) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "ShowLabelPct" ) )
		{
			return ((fShowLblPct) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "ShowPct" ) )
		{
			return ((fShowPct) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "ShowCatLabel" ) )
		{
			return ((fShowCatLabel) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "ShowBubbleSizes" ) )
		{
			return ((fShowBubbleSizes) ? "1" : "0");
		}
		else if( op.equalsIgnoreCase( "TextRotation" ) )
		{
			return String.valueOf( trot );
		} /*else if (op.equalsIgnoreCase("Font")) {
			return setFontId(Short.parseShort(val));
    		bHandled= true;    		
    	}*/
		return "";
	}

	/**
	 * @return
	 */
	protected boolean isChartTitle()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == OBJECTLINK )
			{
				ObjectLink ol = (ObjectLink) b;
				return ol.isChartTitle();
			}
		}
		return false;
	}

	protected boolean isXAxisLabel()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == OBJECTLINK )
			{
				ObjectLink ol = (ObjectLink) b;
				return ol.isXAxisLabel();
			}
		}
		return false;
	}

	protected boolean isYAxisLabel()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == OBJECTLINK )
			{
				ObjectLink ol = (ObjectLink) b;
				return ol.isYAxisLabel();
			}
		}
		return false;
	}

	/**
	 * @return truth of "This TextDisp represents the Z axis"
	 */
	protected boolean isZAxisLabel()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == OBJECTLINK )
			{
				ObjectLink ol = (ObjectLink) b;
				return ol.getType() == ObjectLink.TYPE_ZAXIS;
			}
		}
		return false;
	}

	/**
	 * Return the string associated with this TextDisp object.
	 *
	 * @return
	 */
	public String toString()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == SERIESTEXT )
			{
				SeriesText st = (SeriesText) b;
				return st.toString();
			}
		}
		return "";
	}

	/**
	 * Set the text of this textDisplay object.
	 * Requires setting the child seriesText object
	 */
	public void setText( String newText )
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == SERIESTEXT )
			{
				SeriesText st = (SeriesText) b;
				st.setText( newText );
				break;
			}
		}
	}

	/**
	 * add new TextDisp rec with associated recs, according to type
	 *
	 * @see ObjectLink
	 */
	public static XLSRecord getPrototype( int type, String text, com.extentech.formats.XLS.WorkBook book )
	{
		TextDisp td = new TextDisp();
		td.setOpcode( TEXTDISP );
		td.setData( td.PROTOTYPE_BYTES );
		td.init();
		if( type != ObjectLink.TYPE_DATAPOINTS )
		{ // for all TextDisps for labels i.e. axes, labels, must set fAutogenerated and fAutoText off
			td.grbit = ByteTools.updateGrBit( td.grbit, false, 4 );
			td.grbit = ByteTools.updateGrBit( td.grbit, false, 5 );
			td.updateRecord();
			td.init();
		}
		td.setWorkBook( book );
		// add pos record
		Pos p = (Pos) Pos.getPrototype( Pos.TYPE_TEXTDISP );
		p.setWorkBook( book );
		if( type == ObjectLink.TYPE_XAXIS )
		{    // 20090309 KSC: set relative position
			p.setX( -4 );
			p.setY( -3 );
		}
		td.addChartRecord( p );
		// add fontx rec, if necesssary
		if( type != ObjectLink.TYPE_DATAPOINTS )
		{
			Fontx f = (Fontx) Fontx.getPrototype();
			f.setWorkBook( book );
			td.addChartRecord( f );
		}
		// add ai record
		Ai ai = null;
		if( type == ObjectLink.TYPE_TITLE )
		{
			ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_LEGEND );
		}
		else
		{
			ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_NULL_LEGEND );
		}
		ai.setWorkBook( book );
		td.addChartRecord( ai );
		// add seriestext rec, if necessary
		if( type != ObjectLink.TYPE_DATAPOINTS )
		{
			SeriesText st = (SeriesText) SeriesText.getPrototype( text );
			td.addChartRecord( st );
		}
		// add objectlink rec
		ObjectLink o = (ObjectLink) ObjectLink.getPrototype( type );
		o.setWorkBook( book );
		td.addChartRecord( o );
		return td;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			2, 2, 1, 0, 0, 0, 0, 0, -33, -1, -1, -1, -74, -1, -1, -1, 0, 0, 0, 0, 0, 0, 0, 0, -79, 0, 77, 0, -96, 9, 0, 0
	};

	public static XLSRecord getPrototype()
	{
		TextDisp td = new TextDisp();
		td.setOpcode( TEXTDISP );
		td.setData( td.PROTOTYPE_BYTES );
		td.init();
		return td;
	}

	private void updateRecord()
	{
		this.getData()[0] = (byte) at;
		this.getData()[1] = (byte) vat;
		byte[] b = ByteTools.shortToLEBytes( wBkgMode );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
		b = new byte[4];
		b[0] = (byte) rgbText.getRed();
		b[1] = (byte) rgbText.getGreen();
		b[2] = (byte) rgbText.getBlue();
		b[3] = 0;    // reserved/0
		System.arraycopy( b, 0, this.getData(), 0, 4 );
		// x
		// y
		// dx
		// dy
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[24] = b[0];
		this.getData()[25] = b[1];
		// icvText
		b = ByteTools.shortToLEBytes( grbit2 );
		this.getData()[28] = b[0];
		this.getData()[29] = b[1];
		b = ByteTools.shortToLEBytes( trot );
		this.getData()[30] = b[0];
		this.getData()[31] = b[1];
	}

	public static int convertType( int axis )
	{
		int t = axis;
		switch( axis )
		{
			case XAXIS:
				t = ObjectLink.TYPE_XAXIS;
				break;
			case YAXIS:
			case XVALAXIS:
				t = ObjectLink.TYPE_YAXIS;
				break;
			case ZAXIS:
				t = ObjectLink.TYPE_ZAXIS;
				break;
		}
		return t;
	}

	/**
	 * return the coordinates, in pixels, of the text area, if possible
	 *
	 * @return
	 */
	public float[] getCoords()
	{
		Pos p = (Pos) Chart.findRec( this.chartArr, Pos.class );
		if( p != null )
		{
			float[] coords = p.getCoords();
			return coords;
		}
		return new float[]{ 0, 0, 0, 0 };

	}

	/**
	 * return the rotation of this text
	 * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
	 * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
	 * 255			Text top-to-bottom with letters upright
	 */
	public int getRotation()
	{
		return trot;
	}

	/**
	 * sets the rotation of this text
	 * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
	 * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
	 * 255			Text top-to-bottom with letters upright
	 */
	public void setRotation( int rot )
	{
		trot = (short) rot;
		byte[] b = ByteTools.shortToLEBytes( trot );
		this.getData()[30] = b[0];
		this.getData()[31] = b[1];
	}

	/**
	 * return the text color of the text of this text display
	 *
	 * @return String hex color string
	 */
	public String getTextColor()
	{
		/*		Font fx = this.getFont(wkbook);
		if (fx!=null)
			return fx.getFontColor();
		return 0;    			
*/        // it seems that font color and TextDisp.rgbText are the same
		return FormatHandle.colorToHexString( rgbText );
	}

	/**
	 * return the SVG representation of this Text Display
	 *
	 * @return
	 */
	public StringBuffer getSVG( HashMap<String, Double> chartMetrics )
	{
		StringBuffer svg = new StringBuffer();

		float[] coords = this.getCoords();
		coords[0] = (int) Math.ceil( Pos.convertFromSPRC( coords[0], chartMetrics.get( "canvasw" ).floatValue(), 0 ) ) - 3;
		coords[1] = (int) Math.ceil( Pos.convertFromSPRC( coords[1], 0, chartMetrics.get( "canvash" ).floatValue() ) );
		Font fx = this.getFont( wkbook );

		svg.append( "<g>\r\n" );
		float fh = 10;
		if( fx != null )
		{
			fh = (float) fx.getFontHeightInPoints();
		}
		float x = (float) (chartMetrics.get( "x" ) + (chartMetrics.get( "w" ) / 2));
		float y = (float) (chartMetrics.get( "TITLEOFFSET" ).floatValue() + fh) / 2;
		Frame f = (Frame) Chart.findRec( this.chartArr, Frame.class );
		if( f != null )
		{
			coords[0] = coords[0] + x;
			coords[1] = y + coords[1] + fh;
			coords[3] = fh * 2;    // just a test really
			svg.append( f.getSVG( coords ) );
		}
		else
		{
			coords[0] = x;
			coords[1] = y + fh / 2;
		}

		svg.append( "<text " + GenericChartObject.getScript( "charttitle" ) + " x='" + (coords[0]) + "' y='" + (coords[1]) + "' style='text-anchor: middle;' alignment-baseline='text-after-edge' " );
		if( fx != null )
		{
			svg.append( " " + fx.getSVG() + ">" );
		}
		else
		{
			svg.append( " font-family='Arial' font-size='14pt' font-weight='bold'>" );
		}
		svg.append( this.toString() + "</text>\r\n" );
		svg.append( "</g>\r\n" );
		return svg;
	}

	/**
	 * create a bg frame with the specified settings
	 *
	 * @param lw      line width (-1 is none, 0= default ...)
	 * @param lclr    line color
	 * @param bgcolor frame bg color
	 * @param coords  frame coordinates
	 */
	public void setFrame( int lw, int lclr, int bgcolor, float[] coords )
	{
		Frame f = (Frame) Chart.findRec( this.chartArr, Frame.class );
		if( f == null )
		{
			f = (Frame) Frame.getPrototype();
			f.addBox( lw, lclr, bgcolor );
			f.setParentChart( this.getParentChart() );
			this.chartArr.add( f );
		}
	}
}
