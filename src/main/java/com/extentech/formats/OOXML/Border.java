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
import com.extentech.formats.XLS.Xf;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;

/**
 * border OOXML element
 * <p/>
 * parent:  	styleSheet/borders element in styles.xml
 * children: 	SEQ: left, right, top, bottom, diagonal, vertical, horizontal
 */
public class Border implements OOXMLElement
{

	private static final long serialVersionUID = 4340789910636828223L;
	private HashMap<String, String> attrs = null;
	private HashMap<String, BorderElement> borderElements = null;

	public Border()
	{
	}

	/**
	 * @param styles int array {top, left, top, bottom, right, [diagonal]}
	 * @param colors int array {top, left, top, bottom, right, [diagonal]}
	 */
	public Border( HashMap<String, String> attrs, HashMap<String, BorderElement> borderElements )
	{
		this.attrs = attrs;
		this.borderElements = borderElements;
	}

	public Border( Border b )
	{
		this.attrs = b.attrs;
		this.borderElements = b.borderElements;
	}

	/**
	 * set borders
	 *
	 * @param bk
	 * @param styles t, l, b, r
	 * @param colors
	 */
	public Border( WorkBookHandle bk, int[] styles, int[] colors )
	{
		this.borderElements = new HashMap<String, BorderElement>();
		String[] borderElements = { "top", "left", "bottom", "right" };
		for( int i = 0; i < 4; i++ )
		{
			if( styles[i] > 0 )
			{
				String style = OOXMLConstants.borderStyle[styles[i]];
				this.borderElements.put( borderElements[i], new BorderElement( style, colors[i], borderElements[i], bk ) );
			}
		}
		// diagonal? vertical?  horizontal?
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		HashMap<String, BorderElement> borderElements = new HashMap<String, BorderElement>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "border" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "left" ) )
					{
						borderElements.put( "left", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "right" ) )
					{
						borderElements.put( "right", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "top" ) )
					{
						borderElements.put( "top", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "bottom" ) )
					{
						borderElements.put( "bottom", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "diagonal" ) )
					{
						borderElements.put( "diagonal", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "vertical" ) )
					{
						borderElements.put( "vertical", BorderElement.parseOOXML( xpp, bk ) );
					}
					else if( tnm.equals( "horizontal" ) )
					{
						borderElements.put( "horizontal", BorderElement.parseOOXML( xpp, bk ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "border" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "border.parseOOXML: " + e.toString() );
		}
		Border b = new Border( attrs, borderElements );
		return b;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<border" );
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = i.next();
				String val = attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		ooxml.append( ">" );
		if( borderElements.get( "left" ) != null )
		{
			ooxml.append( borderElements.get( "left" ).getOOXML() );
		}
		if( borderElements.get( "right" ) != null )
		{
			ooxml.append( borderElements.get( "right" ).getOOXML() );
		}
		if( borderElements.get( "top" ) != null )
		{
			ooxml.append( borderElements.get( "top" ).getOOXML() );
		}
		if( borderElements.get( "bottom" ) != null )
		{
			ooxml.append( borderElements.get( "bottom" ).getOOXML() );
		}
		if( borderElements.get( "diagonal" ) != null )
		{
			ooxml.append( borderElements.get( "diagonal" ).getOOXML() );
		}
		if( borderElements.get( "vertical" ) != null )
		{
			ooxml.append( borderElements.get( "vertical" ).getOOXML() );
		}
		if( borderElements.get( "horizontal" ) != null )
		{
			ooxml.append( borderElements.get( "horizontal" ).getOOXML() );
		}

		ooxml.append( "</border>" );
		return ooxml.toString();
	}

	/**
	 * returns an array representing the border sizes
	 * <br>top, left, bottom, right, diag
	 *
	 * @return int[5] representing border sizes
	 */
	public int[] getBorderSizes()
	{
		int[] sizes = new int[5];
		if( borderElements.get( "top" ) != null )
		{
			sizes[0] = borderElements.get( "top" ).getBorderSize();
		}
		if( borderElements.get( "left" ) != null )
		{
			sizes[1] = borderElements.get( "left" ).getBorderSize();
		}
		if( borderElements.get( "bottom" ) != null )
		{
			sizes[2] = borderElements.get( "bottom" ).getBorderSize();
		}
		if( borderElements.get( "right" ) != null )
		{
			sizes[3] = borderElements.get( "right" ).getBorderSize();
		}
		if( borderElements.get( "diagonal" ) != null )
		{
			sizes[4] = borderElements.get( "diagonal" ).getBorderSize();
		}
		return sizes;
	}

	/**
	 * returns an array representing the border styles
	 * translated from OOXML String value to 2003-int value
	 * <br>top, left, bottom, right, diag
	 *
	 * @return int[5]
	 */
	public int[] getBorderStyles()
	{
		int[] styles = new int[5];
		if( borderElements.get( "top" ) != null )
		{
			styles[0] = borderElements.get( "top" ).getBorderStyle();
		}
		if( borderElements.get( "left" ) != null )
		{
			styles[1] = borderElements.get( "left" ).getBorderStyle();
		}
		if( borderElements.get( "bottom" ) != null )
		{
			styles[2] = borderElements.get( "bottom" ).getBorderStyle();
		}
		if( borderElements.get( "right" ) != null )
		{
			styles[3] = borderElements.get( "right" ).getBorderStyle();
		}
		if( borderElements.get( "diagonal" ) != null )
		{
			styles[4] = borderElements.get( "diagonal" ).getBorderStyle();
		}
		return styles;
	}

	/**
	 * returns an array representing the border colors as rgb string
	 * <br>top, left, bottom, right, diag
	 *
	 * @return String[6]
	 */
	public String[] getBorderColors()
	{
		try
		{
			String[] clrs = new String[5];
			if( borderElements.get( "top" ) != null )
			{
				clrs[0] = borderElements.get( "top" ).getBorderColor();
			}
			if( borderElements.get( "left" ) != null )
			{
				clrs[1] = borderElements.get( "left" ).getBorderColor();
			}
			if( borderElements.get( "bottom" ) != null )
			{
				clrs[2] = borderElements.get( "bottom" ).getBorderColor();
			}
			if( borderElements.get( "right" ) != null )
			{
				clrs[3] = borderElements.get( "right" ).getBorderColor();
			}
			if( borderElements.get( "diagonal" ) != null )
			{
				clrs[4] = borderElements.get( "diagonal" ).getBorderColor();
			}
			return clrs;
		}
		catch( NullPointerException e )
		{
			return new String[5];
		}
	}

	/**
	 * returns an array representing the border colors as rgb string
	 * <br>top, left, bottom, right, diag
	 *
	 * @return String[6]
	 */
	public int[] getBorderColorInts()
	{
		try
		{
			int[] clrs = new int[5];
			if( borderElements.get( "top" ) != null )
			{
				clrs[0] = borderElements.get( "top" ).getBorderColorInt();
			}
			if( borderElements.get( "left" ) != null )
			{
				clrs[1] = borderElements.get( "left" ).getBorderColorInt();
			}
			if( borderElements.get( "bottom" ) != null )
			{
				clrs[2] = borderElements.get( "bottom" ).getBorderColorInt();
			}
			if( borderElements.get( "right" ) != null )
			{
				clrs[3] = borderElements.get( "right" ).getBorderColorInt();
			}
			if( borderElements.get( "diagonal" ) != null )
			{
				clrs[4] = borderElements.get( "diagonal" ).getBorderColorInt();
			}
			return clrs;
		}
		catch( NullPointerException e )
		{
			return new int[5];
		}
	}

	public OOXMLElement cloneElement()
	{
		return new Border( this );
	}

	public String toString()
	{
		if( borderElements != null )
		{
			return borderElements.toString();
		}
		return "<none>";
	}

	/**
	 * return an OOXML representation of this border based on this FormatHandle object
	 */
	public static String getOOXML( Xf xf )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<border>" );
		int[] lineStyles = new int[5];
		lineStyles[0] = xf.getLeftBorderLineStyle();
		lineStyles[1] = xf.getRightBorderLineStyle();
		lineStyles[2] = xf.getTopBorderLineStyle();
		lineStyles[3] = xf.getBottomBorderLineStyle();
		lineStyles[4] = xf.getDiagBorderLineStyle();

		int[] colors = new int[5];
		colors[0] = xf.getLeftBorderColor();
		colors[1] = xf.getRightBorderColor();
		colors[2] = xf.getTopBorderColor();
		colors[3] = xf.getBottomBorderColor();
		colors[4] = xf.getDiagBorderColor();

		if( lineStyles[0] > 0 )
		{
			ooxml.append( "<left" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[lineStyles[0]] + "\"" );
			if( colors[0] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[0]] )
				                                             .substring( 1 ) + "\"/></left>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}

		if( lineStyles[1] > 0 )
		{
			ooxml.append( "<right" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[lineStyles[1]] + "\"" );
			if( colors[1] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[1]] )
				                                             .substring( 1 ) + "\"/></right>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}

		if( lineStyles[2] > 0 )
		{
			ooxml.append( "<top" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[lineStyles[2]] + "\"" );
			if( colors[2] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[2]] )
				                                             .substring( 1 ) + "\"/></top>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}

		if( lineStyles[3] > 0 )
		{
			ooxml.append( "<bottom" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[lineStyles[3]] + "\"" );
			if( colors[3] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[3]] )
				                                             .substring( 1 ) + "\"/></bottom>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}

		if( lineStyles[4] > 0 )
		{
			ooxml.append( "<diagonal" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[lineStyles[4]] + "\"" );
			if( colors[4] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[4]] )
				                                             .substring( 1 ) + "\"/></diagonal>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		ooxml.append( "</border>" );
		return ooxml.toString();
	}

	/**
	 * return the OOMXL to define a border based on the below specifications:
	 *
	 * @param styles int array {top, left, top, bottom, right, [diagonal]}
	 * @param colors int array {top, left, top, bottom, right, [diagonal]}
	 */
	public static String getOOXML( int[] styles, int[] colors )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<border>" );
		// left
		if( styles[0] > 0 )
		{
			ooxml.append( "<left" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[styles[1]] + "\"" );
			if( colors[0] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[0]] )
				                                             .substring( 1 ) + "\"/></left>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		// right
		if( styles[1] > 0 )
		{
			ooxml.append( "<right" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[styles[3]] + "\"" );
			if( colors[1] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[1]] )
				                                             .substring( 1 ) + "\"/></right>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		// top
		if( styles[2] > 0 )
		{
			ooxml.append( "<top" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[styles[0]] + "\"" );
			if( colors[2] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[2]] )
				                                             .substring( 1 ) + "\"/></top>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		// bottom
		if( styles[3] > 0 )
		{
			ooxml.append( "<bottom" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[styles[2]] + "\"" );
			if( colors[3] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[3]] )
				                                             .substring( 1 ) + "\"/></top>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		// diagonal
		if( styles[4] > 0 )
		{
			ooxml.append( "<diagonal" );
			ooxml.append( " style=\"" + OOXMLConstants.borderStyle[styles[4]] + "\"" );
			if( colors[4] > 0 )
			{
				ooxml.append( "><color rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[colors[4]] )
				                                             .substring( 1 ) + "\"/></diagonal>" );
			}
			else
			{
				ooxml.append( "/>" );
			}
		}
		ooxml.append( "</border>" );
		return ooxml.toString();
	}
}

/**
 * one of:
 * left, right, top, bottom, diagonal, vertical, horizontal
 * <p/>
 * parent: 	border
 * children:  color
 */
class BorderElement implements OOXMLElement
{

	private static final long serialVersionUID = -8040551653089261574L;
	private String style;
	private Color color;
	private String borderElement;

	public BorderElement( String style, Color c, String borderElement )
	{
		this.style = style;
		this.color = c;
		this.borderElement = borderElement;
	}

	public BorderElement( BorderElement b )
	{
		this.style = b.style;
		this.color = b.color;
		this.borderElement = b.borderElement;
	}

	/**
	 * @return the border size for this border element, translated from OOXML string to 2003-style int
	 */
	public int getBorderSize()
	{
		int st = getBorderStyle();
		if( st <= 4 )
		{
			return st + 1;
		}
		// otherwise, interpret style --> size???
		if( st == 7 )    // hair
		{
			return 1;
		}
		if( st == 6 || st == 8 || st == 0xC )
		{
			return 3;
		}
		return 2;
	}

	/**
	 * return the border style for this border element, translated from OOXML string to 2003-style int
	 *
	 * @return
	 */
	public int getBorderStyle()
	{
		if( style == null || style.equals( "none" ) )
		{
			return -1;
		}
		else if( style.equals( "thin" ) )
		{
			return 1;
		}
		else if( style.equals( "medium" ) )
		{
			return 2;
		}
		else if( style.equals( "dashed" ) )
		{
			return 3;
		}
		else if( style.equals( "dotted" ) )
		{
			return 4;
		}
		else if( style.equals( "thick" ) )
		{
			return 5;
		}
		else if( style.equals( "double" ) )
		{
			return 6;
		}
		else if( style.equals( "hair" ) )
		{
			return 7;
		}
		else if( style.equals( "mediumDashed" ) )
		{
			return 8;
		}
		else if( style.equals( "dashDot" ) )
		{
			return 9;
		}
		else if( style.equals( "mediumDashDot" ) )
		{
			return 0xA;
		}
		else if( style.equals( "dashDotDot" ) )
		{
			return 0xB;
		}
		else if( style.equals( "mediumDashDotDot" ) )
		{
			return 0xC;
		}
		else if( style.equals( "slantDashDot" ) )
		{
			return 0xD;
		}
		return -1;
	}

	/**
	 * return the rgb color string for this border element
	 *
	 * @return
	 */
	public String getBorderColor()
	{
		if( color != null )
		{
			return this.color.getColorAsOOXMLRBG();
		}
		return null;
	}

	public int getBorderColorInt()
	{
		if( color != null )
		{
			return this.color.getColorInt();
		}
		return 0;
	}

	/**
	 * constructor from string representation of a border element
	 *
	 * @param style         "thin", "thick" ...
	 * @param val           rgb color value
	 * @param borderElement "left", "right", "top", "bottom", "diagonal"
	 */
	public BorderElement( String style, String val, String borderElement, WorkBookHandle bk )
	{
		this.style = style;
		if( style != null && val != null )
		{
			this.color = new Color( "color", false, Color.COLORTYPERGB, val, 0.0, (short) 0, bk.getWorkBook().getTheme() );
		}
		this.borderElement = borderElement;
	}

	/**
	 * create a new border element
	 *
	 * @param style         "thin", "thick" ...
	 * @param val           color int
	 * @param borderElement "left", "right", "top", "bottom", "diagonal"
	 */
	public BorderElement( String style, int val, String borderElement, WorkBookHandle bk )
	{
		this.style = style;
		if( style != null && val != -1 )
		{
			this.color = new Color( "color",
			                        false,
			                        Color.COLORTYPEINDEXED,
			                        String.valueOf( val ),
			                        0.0,
			                        (short) 0,
			                        bk.getWorkBook().getTheme() );
		}
		this.borderElement = borderElement;
	}

	public static BorderElement parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		String style = null;
		Color c = null;
		String borderElement = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "color" ) )
					{
						c = (Color) Color.parseOOXML( xpp, (short) 0, bk );
					}
					else
					{    // one of the border elements
						if( xpp.getAttributeCount() > 0 )    // style
						{
							style = xpp.getAttributeValue( 0 );
						}
						borderElement = tnm;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( borderElement ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "borderElement.parseOOXML: " + e.toString() );
		}
		BorderElement b = new BorderElement( style, c, borderElement );
		return b;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + borderElement );
		if( style != null )
		{
			ooxml.append( " style=\"" + style + "\">" );
			if( color != null )
			{
				ooxml.append( color.getOOXML() );
			}
			ooxml.append( "</" + borderElement + ">" );
		}
		else
		{
			ooxml.append( "/>" );
		}
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new BorderElement( this );
	}

	public String toString()
	{
		return ((style != null) ? style : "<none>") +
				" c:" + ((color != null) ? color.toString() : "<none>");

	}
}

