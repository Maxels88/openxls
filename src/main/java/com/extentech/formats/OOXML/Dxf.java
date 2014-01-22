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
package com.extentech.formats.OOXML;

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.Font;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Arrays;

/**
 * dxf (Formatting) OOXML Element
 * <p/>
 * A single dxf record, expressing incremental formatting to be applied.
 * Used for Conditional Formatting, Tables, Sort Conditions, Color filters ...
 * <p/>
 * Differntial Formatting:
 * define formatting for all non-cell formatting in the workbook. Whereas xf records fully specify a particular aspect of formatting (e.g., cell borders)
 * by referencing those formatting definitions elsewhere in the Styles part, dxf records specify incremental (or
 * differential) aspects of formatting directly inline within the dxf element. The dxf formatting is to be applied on
 * top of or in addition to any formatting already present on the object using the dxf record.
 * <p/>
 * parent:  (StyleSheet styles.xml) dxfs
 * chilren: SEQUENCE:  font, numFmt, fill, alignment, border, protection
 */
// TODO: protection element
public class Dxf implements OOXMLElement
{

	private static final long serialVersionUID = -5999328795988018131L;
	private Font font = null;
	private NumFmt numFmt = null;
	private Fill fill = null;
	private Alignment alignment = null;
	private Border border = null;
	private WorkBookHandle wbh = null;

	public Dxf( Font fnt, NumFmt nf, Fill f, Alignment a, Border b, WorkBookHandle wbh )
	{
		this.font = fnt;
		this.numFmt = nf;
		this.fill = f;
		this.alignment = a;
		this.border = b;
		this.wbh = wbh;
	}

	public Dxf( Dxf d )
	{
		this.font = d.font;
		this.numFmt = d.numFmt;
		this.fill = d.fill;
		this.alignment = d.alignment;
		this.border = d.border;
		this.wbh = d.wbh;
	}

	public Dxf()
	{
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		Font fnt = null;
		NumFmt nf = null;
		Fill f = null;
		Alignment a = null;
		Border b = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "font" ) )
					{
						fnt = (Font) Font.parseOOXML( xpp, bk );
					}
					else if( tnm.equals( "numFmt" ) )
					{
						nf = (NumFmt) NumFmt.parseOOXML( xpp );
					}
					else if( tnm.equals( "fill" ) )
					{
						f = (Fill) Fill.parseOOXML( xpp, true, bk );
					}
					else if( tnm.equals( "alignment" ) )
					{
						a = (Alignment) Alignment.parseOOXML( xpp );
					}
					else if( tnm.equals( "border" ) )
					{
						b = (Border) Border.parseOOXML( xpp, bk );
					}
					else if( tnm.equals( "protection" ) )
					{
						// TODO: finish
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "dxf" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "dxf.parseOOXML: " + e.toString() );
		}
		Dxf d = new Dxf( fnt, nf, f, a, b, bk );
		return d;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<dxf>" );
		if( font != null )
		{
			ooxml.append( font.getOOXML() );
		}
		if( numFmt != null )
		{
			ooxml.append( numFmt.getOOXML() );
		}
		if( fill != null )
		{
			ooxml.append( fill.getOOXML( true ) );
		}
		if( alignment != null )
		{
			ooxml.append( alignment.getOOXML() );
		}
		if( border != null )
		{
			ooxml.append( border.getOOXML() );
		}
		ooxml.append( "</dxf>" );
		return ooxml.toString();
	}

	public int[] getBorderColors()
	{
		if( border != null )
		{
			return border.getBorderColorInts();
		}
		return null;
	}

	public int[] getBorderStyles()
	{
		if( border != null )
		{
			int[] styles = border.getBorderStyles();
			if( Arrays.equals( new int[]{ 0, 0, 0, 0 }, styles ) )
			{
				return null;
			}
			return styles;
		}
		return null;
	}

	public int[] getBorderSizes()
	{
		if( border != null )
		{
			return border.getBorderSizes();
		}
		return null;
	}

	public boolean isStriken()
	{
		if( font != null )
		{
			return font.getStricken();
		}
		return false;
	}

	public int getFg()
	{
		if( fill != null )
		{
			return fill.getFgColorAsInt( wbh.getWorkBook().getTheme() );
		}
		return -1;
	}

	public int getFillPatternInt()
	{
		if( fill != null )
		{
			return fill.getFillPatternInt();
		}
		return -1;
	}

	/**
	 * returns the OOXML Fill element
	 *
	 * @return
	 */
	public Fill getFill()
	{
		return fill;
	}

	public int getBg()
	{
		if( fill != null )
		{
			return fill.getBgColorAsInt( wbh.getWorkBook().getTheme() );
		}
		return -1;
	}

	public String getBgColorAsString()
	{
		if( fill != null )
		{
			return fill.getBgColorAsRGB( wbh.getWorkBook().getTheme() );
		}
		return null;
	}

	public String getHorizontalAlign()
	{
		if( alignment != null )
		{
			return alignment.getAlignment( "horizontal" );
		}
		return null;
	}

	public String getVerticalAlign()
	{
		if( alignment != null )
		{
			return alignment.getAlignment( "vertical" );
		}
		return null;
	}

	public String getNumberFormat()
	{
		if( numFmt != null )
		{
			return numFmt.getFormatId();
		}
		return null;
	}

	/**
	 * returns the Font for ths dxf, if any
	 *
	 * @return
	 */
	public Font getFont()
	{
		return font;
	}

	public int getFontHeight()
	{
		if( font != null )
		{
			return font.getFontHeight();
		}
		return -1;
	}

	public int getFontWeight()
	{
		if( font != null )
		{
			return font.getFontWeight();
		}
		return -1;
	}

	public String getFontName()
	{
		if( font != null )
		{
			return font.getFontName();
		}
		return null;
	}

	public int getFontColor()
	{
		if( font != null )
		{
			return font.getColor();
		}
		return -1;
	}

	public boolean isItalic()
	{
		if( font != null )
		{
			return font.getItalic();
		}
		return false;
	}

	public int getFontUnderline()
	{
		if( font != null )
		{
			return font.getUnderlineStyle();
		}
		return -1;
	}

	/**
	 * return a String representation of this Dxf in "style properties" notation
	 *
	 * @return String representation of this Dxf
	 * @see Cf.setStylePropsFromString
	 */
	public String getStyleProps()
	{
		StringBuffer props = new StringBuffer();

		// fill
		if( fill != null )
		{
			props.append( "pattern:" + fill.getFillPatternInt() + ";" );
			String s = fill.getFgColorAsRGB( wbh.getWorkBook().getTheme() );
			if( s != null )    // fg is pattern color
			{
				props.append( "patterncolor:#" + s + ";" );
			}
			s = fill.getBgColorAsRGB( wbh.getWorkBook().getTheme() );
			if( s != null )
			{
				props.append( "background:#" + s + ";" );
			}
		}

		// font
		if( font != null )
		{    // note: since this is differential, many of these may not be set
			if( !font.getFontName().equals( "" ) )
			{
				props.append( "font-name" + font.getFontName() + ";" );
			}
			if( font.getFontWeight() > -1 )
			{
				props.append( "font-weight:" + font.getFontWeight() + ";" );
			}
			if( font.getFontHeight() > -1 )
			{
				props.append( "font-Height:" + font.getFontHeight() + ";" );
			}
			props.append( "font-ColorIndex:" + font.getColor() + ";" );
			if( font.getStricken() )
			{
				props.append( "font-Striken:" + font.getStricken() + ";" );
			}
			if( font.getItalic() )
			{
				props.append( "font-italic:" + font.getItalic() + ";" );
			}
			if( font.getUnderlineStyle() != 0 )
			{
				props.append( "font-UnderlineStyle:" + font.getUnderlineStyle() + ";" );
			}
			// TODO: italic, bold
		}

		// borders
		if( border != null )
		{
			int[] sizes = border.getBorderSizes();
			int[] styles = border.getBorderStyles();
			String[] colors = border.getBorderColors();
			props.append( "border-top:" + sizes[0] + " " + FormatHandle.BORDER_NAMES[styles[0]] + " " + colors[0] + ";" );
			props.append( "border-left:" + sizes[1] + " " + FormatHandle.BORDER_NAMES[styles[1]] + " " + colors[1] + ";" );
			props.append( "border-bottom:" + sizes[2] + " " + FormatHandle.BORDER_NAMES[styles[2]] + " " + colors[2] + ";" );
			props.append( "border-right:" + sizes[3] + " " + FormatHandle.BORDER_NAMES[styles[3]] + " " + colors[3] + ";" );
		}

		// alignment
		if( alignment != null )
		{
			String s = alignment.getAlignment( "vertical" );
			if( s != null )
			{
				props.append( "alignment-vertical" + s + ";" );
			}
			s = alignment.getAlignment( "horizontal" );
			if( s != null )
			{
				props.append( "alignment-horizontal" + s + ";" );
			}
		}

		// number format
		if( numFmt != null )
		{
			String s = numFmt.getFormatId();
			props.append( "numberformat:" + s + ";" );

		}

		return props.toString();
	}

	/**
	 * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
	 */
	public void createFont( int w, boolean i, int ustyle, int cl, int h )
	{
		font = new Font( "", w, h );
		if( w == 700 )
		{
			font.setBold( true );    // why doesn't constructor do this?
		}
		if( i )
		{
			font.setItalic( i );
		}
		if( ustyle != 0 )
		{
			font.setUnderlineStyle( (byte) ustyle );
		}
		font.setColor( cl );
	}

	/**
	 * Sets the fill for this dxf from an existing Fill element
	 *
	 * @param f
	 */
	public void setFill( Fill f )
	{
		this.fill = (Fill) f.cloneElement();
	}

	/**
	 * Sts the Font for this dxf from an existing Font
	 *
	 * @param f
	 */
	public void setFont( Font f )
	{
		this.font = (Font) f;
	}

	/**
	 * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
	 */
	public void createFill( int fs, int fg, int bg, WorkBookHandle bk )
	{
		if( fs < 0 || fs > OOXMLConstants.patternFill.length )
		{
			this.fill = new Fill( null, fg, bg, bk.getWorkBook().getTheme() );    // meaning it's the default (solid bg) pattern
		}
		else
		{
			this.fill = new Fill( OOXMLConstants.patternFill[fs], fg, bg, bk.getWorkBook().getTheme() );
		}
	}

	/**
	 * for BIFF8->OOXML Compatiblity, create a dxf from Cf style info
	 */
	public void createBorder( WorkBookHandle bk, int[] styles, int[] colors )
	{
		border = new Border( bk, styles, colors );
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Dxf( this );
	}
}

