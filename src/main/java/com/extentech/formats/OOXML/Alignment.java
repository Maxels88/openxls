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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;

/**
 * alignment (Alignment) OOMXL element
 * <p/>
 * Formatting information pertaining to text alignment in cells. There are a variety of choices for how text is
 * aligned both horizontally and vertically, as well as indentation settings, and so on.
 * <p/>
 * parent:		(styles.xml) xf, dxf
 * children: none
 */
public class Alignment implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Alignment.class );
	private static final long serialVersionUID = 995367747930839216L;
	private HashMap<String, String> attrs = null;

	public Alignment( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Alignment( Alignment a )
	{
		this.attrs = a.attrs;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "alignment" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "alignment" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "alignment.parseOOXML: " + e.toString() );
		}
		Alignment a = new Alignment( attrs );
		return a;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<alignment" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Alignment( this );
	}

	/**
	 * @param type horizontal vertical
	 * @return
	 */
	public String getAlignment( String type )
	{
		// attributes:
		// horizontal: center, centerContinuous, fill, general, justify, left, right, distributed
		// indent: int value
		// justifyLastLine: bool
		// readingOrder:	0=Context Dependent, 1=Left-to-Right, 2=Right-to-Left
		// relativeIndent: #
		// shrinkToFit: bool
		// textRotation:	degrees from 0-180
		// vertical: bottom, centered, distributed, justify, top
		// wrapText	- true/false
		if( attrs != null )
		{
			return attrs.get( type );
		}
		return null;
	}
}

