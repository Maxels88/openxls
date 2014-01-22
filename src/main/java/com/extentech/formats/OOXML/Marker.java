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

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * class holds OOXML to define markers for a chart (radar, scatter or line)
 * <p/>
 * parent:  dPt, Ser ...
 * children: symbol, size, spPr
 * NOTE: child elements symbol and size have only 1 attribute and no children, and so are treated as strings
 */
public final class Marker implements OOXMLElement
{

	private static final long serialVersionUID = -5070227633357072878L;
	private SpPr sp;
	private String size;
	private String symbol;

	public Marker( String s, String sz, SpPr sp )
	{
		this.symbol = s;
		this.size = sz;
		this.sp = sp;
	}

	public Marker( Marker m )
	{
		this.symbol = m.symbol;
		this.size = m.size;
		this.sp = m.sp;
	}

	/**
	 * parse marker OOXML element
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return marker object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		SpPr sp = null;
		String size = null;        // size element:  val is only attribute + no children
		String symbol = null;  // symbol element:  val is only attribute + no children
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "symbol" ) )
					{
						symbol = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "size" ) )
					{
						size = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
						//sp.setNS("c");
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "marker" ) )
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
			Logger.logErr( "marker.parseOOXML: " + e.toString() );
		}
		Marker m = new Marker( symbol, size, sp );
		return m;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:marker>" );
		if( this.symbol != null )
		{
			ooxml.append( "<c:symbol val=\"" + this.symbol + "\"/>" );
		}
		if( this.size != null )
		{
			ooxml.append( "<c:size val=\"" + this.size + "\"/>" );
		}
		if( this.sp != null )
		{
			ooxml.append( sp.getOOXML() );
		}
		ooxml.append( "</c:marker>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Marker( this );
	}
}
