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

import java.util.ArrayList;
import java.util.Stack;

/**
 * avLst (List of Shape Adjust Values)
 * <p/>
 * This element specifies the adjust values that will be applied to the specified shape. An adjust value is simply a
 * guide that has a value based formula specified. That is, no calculation takes place for an adjust value guide.
 * Instead, this guide specifies a parameter value that is used for calculations within the shape guides.
 * <p/>
 * parent:	prstGeom, prstTxWarp, custGeom
 * children: gd (shape guide) (0 or more)
 */
public class AvLst implements OOXMLElement
{

	private static final long serialVersionUID = 4823524943145191780L;
	private ArrayList gds = null;

	public AvLst()
	{

	}

	public AvLst( ArrayList gds )
	{
		this.gds = gds;
	}

	public AvLst( AvLst av )
	{
		this.gds = av.gds;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		ArrayList<Gd> gds = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gd" ) )
					{
						lastTag.push( tnm );
						if( gds == null )
						{
							gds = new ArrayList<Gd>();
						}
						gds.add( (Gd) Gd.parseOOXML( xpp, lastTag ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "avLst" ) )
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
			Logger.logErr( "avLst.parseOOXML: " + e.toString() );
		}
		AvLst av = new AvLst( gds );
		return av;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:avLst>" );
		if( gds != null )
		{
			for( int i = 0; i < gds.size(); i++ )
			{
				ooxml.append( ((Gd) gds.get( i )).getOOXML() );
			}
		}
		ooxml.append( "</a:avLst>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new AvLst( this );
	}
}

