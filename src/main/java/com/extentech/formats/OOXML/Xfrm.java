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
import java.util.Stack;

class Xfrm implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Xfrm.class );
	private static final long serialVersionUID = 5383438744617393878L;
	private HashMap<String, String> attrs = null;
	private Off o = null;
	private Ext e = null;
	private String ns = null;

	public Xfrm()
	{
		ns = "xdr";    // set default
	}

	public Xfrm( HashMap<String, String> attrs, Off o, Ext e, String ns )
	{
		this.attrs = attrs;
		this.o = o;
		this.e = e;
		this.ns = ns;
	}

	public Xfrm( Xfrm x )
	{
		attrs = x.attrs;
		o = x.o;
		e = x.e;
		ns = x.ns;
	}

	/**
	 * set the namespace for xfrm element
	 * xdr (graphicFrame) or a(spPr)
	 *
	 * @param ns
	 */
	public void setNS( String ns )
	{
		this.ns = ns;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<>();
		Off o = null;
		Ext e = null;
		String ns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "xfrm" ) )
					{        // get attributes
						ns = xpp.getPrefix();
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "off" ) )
					{
						lastTag.push( tnm );
						o = Off.parseOOXML( xpp, lastTag );
						//o.setNS("a");
					}
					else if( tnm.equals( "ext" ) )
					{
						lastTag.push( tnm );
						e = (Ext) Ext.parseOOXML( xpp, lastTag );
						//e.setNS("a");
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "xfrm" ) )
					{
						lastTag.pop();
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception ex )
		{
			log.error( "xfrm.parseOOXML: " + ex.toString() );
		}
		Xfrm x = new Xfrm( attrs, o, e, ns );
		return x;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + ns + ":xfrm" );
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
		if( o != null )
		{
			ooxml.append( o.getOOXML() );
		}
		if( e != null )
		{
			ooxml.append( e.getOOXML() );
		}
		ooxml.append( "</" + ns + ":xfrm>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Xfrm( this );
	}

}

class Off implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Off.class );
	private static final long serialVersionUID = -7624630398053353694L;
	private HashMap<String, String> attrs = null;
	private String ns = "";

	public Off()
	{
		attrs = new HashMap<>();
		attrs.put( "x", new String( "0" ) );
		attrs.put( "y", new String( "0" ) );
	}

	public Off( HashMap<String, String> attrs, String ns )
	{
		this.attrs = attrs;
		this.ns = ns;
	}

	public Off( Off o )
	{
		attrs = o.attrs;
		ns = o.ns;
	}

	public void setNS( String ns )
	{
		this.ns = ns;
	}

	public static Off parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<>();
		String ns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "off" ) )
					{        // get attributes
						ns = xpp.getPrefix();
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "off" ) )
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
			log.error( "off.parseOOXML: " + e.toString() );
		}
		Off o = new Off( attrs, ns );
		return o;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + ns + ":off" );
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
		return new Off( this );
	}
}
	
