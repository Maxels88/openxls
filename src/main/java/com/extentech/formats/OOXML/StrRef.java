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
 * OOXML element strRef, string reference, child of tx (chart text) element or cat (category) element
 */
public class StrRef implements OOXMLElement
{

	private static final long serialVersionUID = -5992001371281543027L;
	private String stringRef = null;
	private StrCache strCache = null;

	public StrRef( String f, StrCache s )
	{
		this.stringRef = f;
		this.strCache = s;
	}

	/**
	 * parse strRef OOXML element
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spRef object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String f = null;
		StrCache s = null;

		/**
		 * contains (in Sequence) 
		 * f
		 * strRef
		 */
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "f" ) )
					{
						f = com.extentech.formats.XLS.OOXMLAdapter.getNextText( xpp );
					}
					else if( tnm.equals( "strCache" ) )
					{
						lastTag.push( tnm );
						s = (StrCache) StrCache.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "strRef" ) )
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
			Logger.logErr( "title.parseOOXML: " + e.toString() );
		}
		StrRef sr = new StrRef( f, s );
		return sr;
	}

	/**
	 * generate ooxml to define a strRef, part of tx element or cat element
	 *
	 * @return
	 */
	/**
	 * strRef contains f + strRef elements
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<c:strRef>" );
		if( this.stringRef != null )
		{
			tooxml.append( "<c:f>" + this.stringRef + "</c:f>" );
		}
		if( this.strCache != null )
		{
			tooxml.append( strCache.getOOXML() );
		}
		tooxml.append( "</c:strRef>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new StrRef( this.stringRef, this.strCache );
	}

}

/**
 * define OOXML strCache element
 */
class StrCache implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4914374179641060956L;
	private int ptCount = -1, idx = -1;
	private String pt = null;

	public StrCache( int ptCount, int idx, String pt )
	{
		this.ptCount = ptCount;
		this.idx = idx;
		this.pt = pt;
	}

	/**
	 * parse title OOXML element title
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spCache object
	 */
	public static StrCache parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		int ptCount = -1, idx = -1;
		String pt = null;

		/**
		 * contains (in Sequence) 
		 * ptCount
		 * pt
		 */
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "ptCount" ) )
					{
						ptCount = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "pt" ) )
					{
						idx = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						pt = com.extentech.formats.XLS.OOXMLAdapter.getNextText( xpp );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "strCache" ) )
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
			Logger.logErr( "strCache.parseOOXML: " + e.toString() );
		}
		StrCache sc = new StrCache( ptCount, idx, pt );
		return sc;
	}

	/**
	 * generate ooxml to define a strCache element, part of strRef element
	 *
	 * @return
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<c:strCache>" );
		tooxml.append( "<c:ptCount val=\"" + this.ptCount + "\"/>" );
		tooxml.append( "<c:pt idx=\"" + this.idx + "\">" );
		tooxml.append( "<c:v>" + this.pt + "</c:v>" );
		tooxml.append( "</c:pt>" );
		tooxml.append( "</c:strCache>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new StrCache( this.ptCount, this.idx, this.pt );
	}

}
