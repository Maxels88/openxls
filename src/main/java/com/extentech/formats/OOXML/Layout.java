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

import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * chart layout element
 * <p/>
 * <p/>
 * parent:  plotarea, title ...
 * children:  manualLayout
 */
public final class Layout implements OOXMLElement
{

	private static final long serialVersionUID = -6547994902298821138L;
	private ManualLayout ml;

	private Layout( ManualLayout ml )
	{
		this.ml = ml;
	}

	private Layout( Layout l )
	{
		this.ml = l.ml;
	}

	/**
	 * create a new plot area Layout/manual layout element
	 * <br>note that the layout is calculated via "edge" i.e the w and h are the bottom and right edges
	 *
	 * @param target "inner" or "outer" <br>inner specifies the plot area size does not include tick marks and axis labels<br>outer does
	 * @param offs   x, y, w, h as a fraction of the width or height of the actual chart
	 */
	public Layout( String target, double[] offs )
	{
		String[] modes = new String[]{ "edge", "edge", null, null };
		String[] soffs = new String[4];
		for( int i = 0; i < 4; i++ )
		{
			if( offs[i] > 0 )
			{
				soffs[i] = String.valueOf( offs[i] );
			}
		}
		this.ml = new ManualLayout( target, modes, soffs );
	}

	/**
	 * return manual layout coords
	 * TODO: interpret xMode, yMode, hMode, wMode
	 *
	 * @return
	 */
	public float[] getCoords()
	{
		float[] coords = new float[4];
		for( int i = 0; i < 4; i++ )
		{
			if( ml.offs[i] != null )
			{
				coords[i] = new Float( ml.offs[i] );
			}
		}
		return coords;
	}

	/**
	 * parse title OOXML element title
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spPr object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		ManualLayout ml = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "manualLayout" ) )
					{
						lastTag.push( tnm );
						ml = (ManualLayout) ManualLayout.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "layout" ) )
					{
						lastTag.pop();    // pop layout tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "layout.parseOOXML: " + e.toString() );
		}
		Layout l = new Layout( ml );
		return l;
	}

	/**
	 * generate ooxml to define a layout
	 *
	 * @return
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer looxml = new StringBuffer();
		if( ml != null )
		{
			looxml.append( "<c:layout>" );
			looxml.append( ml.getOOXML() );
			looxml.append( "</c:layout>" );
		}
		else
		{
			looxml.append( "<c:layout/>" );
		}
		return looxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Layout( this );
	}

	/**
	 * static version for 2003-style charts, takes offsets and creates a layout element
	 *
	 * @param offs int offsets of chart
	 * @return
	 */
	public static String getOOXML( double[] offs )
	{
		StringBuffer looxml = new StringBuffer();
		looxml.append( "<c:layout>" );
		looxml.append( "<c:manualLayout>" );
		looxml.append( "<c:layoutTarget val=\"inner\"/>" );
		looxml.append( "<c:xMode val=\"edge\"/>" );
		looxml.append( "<c:yMode val=\"edge\"/>" );
		looxml.append( "<c:x val=\"" + offs[0] + "\"/>" );
		looxml.append( "<c:y val=\"" + offs[1] + "\"/>" );
		looxml.append( "<c:w val=\"" + offs[2] + "\"/>" );
		looxml.append( "<c:h val=\"" + offs[3] + "\"/>" );
		looxml.append( "</c:manualLayout>" );
		looxml.append( "</c:layout>" );
		return looxml.toString();
	}
}

/**
 * manualLayout
 * specifies exact position of chart
 * <p/>
 * parent:  layout
 * children:  h, hMode, layoutTarget, w, wMode, x, xMode, y, yMode
 */
class ManualLayout implements OOXMLElement
{

	private static final long serialVersionUID = 6460833211809500902L;
	String[] modes; // xMode, yMode, wMode, hMode 
	String[] offs;    // x, y, w, h;
	String target;

	public ManualLayout( String target, String[] modes, String[] offs )
	{
		this.modes = modes;
		this.target = target;
		this.offs = offs;
	}

	public ManualLayout( ManualLayout ml )
	{
		this.modes = ml.modes;
		this.target = ml.target;
		this.offs = ml.offs;
	}

	public static ManualLayout parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		String[] modes = new String[4]; // xMode, yMode, wMode, hMode
		String[] offs = new String[4];    // x, y, w, h;
		String target = null;

		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "manualLayout" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "ATTR" ) )
							{
								//sz= Integer.valueOf(xpp.getAttributeValue(i)).intValue();
							}
						}
					}
					else if( tnm.equals( "layoutTarget" ) )
					{
						target = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "xMode" ) )
					{
						modes[0] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "yMode" ) )
					{
						modes[1] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "wMode" ) )
					{
						modes[2] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "hMode" ) )
					{
						modes[3] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "x" ) )
					{
						offs[0] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "y" ) )
					{
						offs[1] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "w" ) )
					{
						offs[2] = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "h" ) )
					{
						offs[3] = xpp.getAttributeValue( 0 );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "manualLayout" ) )
					{
						lastTag.pop();
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "manualLayout.parseOOXML: " + e.toString() );
		}
		ManualLayout l = new ManualLayout( target, modes, offs );
		return l;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:manualLayout>" );
		if( target != null )
		{
			ooxml.append( "<c:layoutTarget val=\"" + this.target + "\"/>" );
		}
		if( modes[0] != null )
		{
			ooxml.append( "<c:xMode val=\"" + this.modes[0] + "\"/>" );
		}
		if( this.modes[1] != null )
		{
			ooxml.append( "<c:yMode val=\"" + this.modes[1] + "\"/>" );
		}
		if( this.modes[2] != null )
		{
			ooxml.append( "<c:wMode val=\"" + this.modes[2] + "\"/>" );
		}
		if( this.modes[3] != null )
		{
			ooxml.append( "<c:hMode val=\"" + this.modes[3] + "\"/>" );
		}
		if( this.offs[0] != null )
		{
			ooxml.append( "<c:x val=\"" + this.offs[0] + "\"/>" );
		}
		if( this.offs[1] != null )
		{
			ooxml.append( "<c:y val=\"" + this.offs[1] + "\"/>" );
		}
		if( this.offs[2] != null )
		{
			ooxml.append( "<c:w val=\"" + this.offs[2] + "\"/>" );
		}
		if( this.offs[3] != null )
		{
			ooxml.append( "<c:h val=\"" + this.offs[3] + "\"/>" );
		}
		ooxml.append( "</c:manualLayout>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ManualLayout( this );
	}
}

