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

import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * cNvPr  (Non-Visual Drawing Properties)
 * <p/>
 * OOXML/DrawingML element specifies non-visual canvas properties. This allows for additional information that does not affect
 * the appearance of the picture to be stored.
 * <p/>
 * attributes: descr, hidden, id REQ, name REQ,
 * parents:     nvSpPr, nvPicPr ...
 * children:   hlinkClick, hlinkHover
 */
// TODO: Handle Child elements hlinkClick, hlinkHover
public class CNvPr implements OOXMLElement
{

	private static final long serialVersionUID = -3382139449400844949L;
	//private hlinkClick hc;
	//private hlinkHover hh;
	private String descr = null, name = null;
	private boolean hidden = false;
	private int id = -1;

	public CNvPr()
	{
	}

	public CNvPr(/*hlinkClick hc, hlinkHover hh, */int id, String name, String descr, boolean hidden )
	{
		//this.hc= hc;
		//this.hh= hh;
		this.id = id;
		this.name = name;
		this.descr = descr;
		this.hidden = hidden;
	}

	public CNvPr( CNvPr cnv )
	{
		//this.hc= cnv.hc;
		//this.hh= cnv.hh;
		this.id = cnv.id;
		this.name = cnv.name;
		this.descr = cnv.descr;
		this.hidden = cnv.hidden;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		//hlinkClick hc= null;
		//hlinkHover hh= null;
		String descr = null, name = null;
		boolean hidden = false;
		int id = -1;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cNvPr" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							String val = xpp.getAttributeValue( i );
							if( nm.equals( "id" ) )
							{
								id = Integer.valueOf( val ).intValue();
							}
							else if( nm.equals( "name" ) )
							{
								name = val;
							}
							else if( nm.equals( "descr" ) )
							{
								descr = val;
							}
							else if( nm.equals( "hidden" ) )
							{
								hidden = val.equals( "1" );
							}
						}
					}
					else if( tnm.equals( "hlinkClick" ) )
					{
						//            	hc= (hlinkClick) hlinkClick.parseOOXML(xpp).clone();
					}
					else if( tnm.equals( "hlinkHover" ) )
					{
						//				hh= (hlinkHover) hlinkHover.parseOOXML(xpp).clone();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cNvPr" ) )
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
			Logger.logErr( "cNvPr.parseOOXML: " + e.toString() );
		}
		CNvPr cnv = new CNvPr(/*hc, hh, */id, name, descr, hidden );
		return cnv;
	}

	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<xdr:cNvPr id=\"" + id + "\" name=\"" + OOXMLAdapter.stripNonAscii( name ) + "\"" );
		if( descr != null )
		{
			tooxml.append( " descr=\"" + descr + "\"" );
		}
		if( hidden )
		{
			tooxml.append( " hidden=\"" + ((hidden) ? "1" : "0") + "\"" );
		}
		tooxml.append( ">" );
		// TODO: HANDLE  if (hc!=null) tooxml.append(hc.getOOXML());
		// TODO: HANDLE  if (hh!=null) tooxml.append(hh.getOOXML());
		tooxml.append( "</xdr:cNvPr>" );
		return tooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new CNvPr( this );
	}

	/**
	 * get name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		return name;
	}

	/**
	 * set name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		this.name = name;
	}

	/**
	 * get descr attribute
	 *
	 * @return
	 */

	public String getDescr()
	{
		return descr;
	}

	/**
	 * set description attribute
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setDescr( String descr )
	{
		this.descr = descr;
	}

	/**
	 * set the id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		this.id = id;
	}

	/**
	 * return the id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		return this.id;
	}
}

