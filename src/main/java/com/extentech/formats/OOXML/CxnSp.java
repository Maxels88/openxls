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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * parents:  absoluteAnchor, oneCellAnchor, twoCellAnchor
 * children: nvCxnSpPr, spPr, style
 */
//TODO: finish nvCxnSpPr.cNvCxnSpPr element
public class CxnSp implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CxnSp.class );
	private static final long serialVersionUID = -8492664135843926551L;
	private HashMap<String, String> attrs = new HashMap<>();
	private NvCxnSpPr nvc;
	private SpPr spPr;
	private Style style;

	public CxnSp( HashMap<String, String> attrs, NvCxnSpPr nvc, SpPr sp, Style s )
	{
		this.attrs = attrs;
		this.nvc = nvc;
		this.spPr = sp;
		this.style = s;
	}

	public CxnSp( CxnSp c )
	{
		this.attrs = c.attrs;
		this.nvc = c.nvc;
		this.spPr = c.spPr;
		this.style = c.style;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<>();
		NvCxnSpPr nvc = null;
		SpPr sp = null;
		Style s = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cxnSp" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "nvCxnSpPr" ) )
					{
						lastTag.push( tnm );
						nvc = NvCxnSpPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
//	        			 sp.setNS("xdr");
					}
					else if( tnm.equals( "style" ) )
					{
						lastTag.push( tnm );
						s = (Style) Style.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cxnSp" ) )
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
			log.error( "cxnSp.parseOOXML: " + e.toString() );
		}
		CxnSp c = new CxnSp( attrs, nvc, sp, s );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:cxnSp" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( nvc != null )
		{
			ooxml.append( nvc.getOOXML() );
		}
		if( spPr != null )
		{
			ooxml.append( spPr.getOOXML() );
		}
		if( style != null )
		{
			ooxml.append( style.getOOXML() );
		}
		ooxml.append( "</xdr:cxnSp>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CxnSp( this );
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( nvc != null )
		{
			return nvc.getName();
		}
		return null;
	}

	/**
	 * set cNvPr name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		if( nvc != null )
		{
			nvc.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( nvc != null )
		{
			return nvc.getDescr();
		}
		return null;
	}

	/**
	 * set cNvPr description attribute
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setDescr( String descr )
	{
		if( nvc != null )
		{
			nvc.setDescr( descr );
		}
	}

	/**
	 * get Macro attribute
	 */
	public String getMacro()
	{
		if( attrs.get( "macro" ) != null )
		{
			return attrs.get( "macro" );
		}
		return null;
	}

	/**
	 * set Macro attribute
	 *
	 * @param macro
	 */
	public void setMacro( String macro )
	{
		attrs.put( "macro", macro );
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( nvc != null )
		{
			nvc.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( nvc != null )
		{
			return nvc.getId();
		}
		return -1;
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( spPr != null )
		{
			return spPr.getEmbed();
		}
		return null;
	}

	/**
	 * return the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( spPr != null )
		{
			return spPr.getLink();
		}
		return null;
	}

	/**
	 * set the embed attribute for this blip (the id for the embedded picture)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( spPr != null )
		{
			spPr.setEmbed( embed );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( spPr != null )
		{
			spPr.setLink( link );
		}
	}
}

/**
 * nvCxnSpPr (Non-Visual Properties for a Connection Shape)
 * This element specifies all non-visual properties for a connection shape. This element is a container for the non15
 * visual identification properties, shape properties and application properties that are to be associated with a
 * DrawingML Reference Material - DrawingML - SpreadsheetML Drawing connection shape.
 * This allows for additional information that does not affect 1 the appearance of the connection
 * shape to be stored.
 * <p/>
 * parent: cxnSp
 * children:  cNvPr, cNvCxnSpPr
 */
// TODO: finish cNvCxnSpPr
class NvCxnSpPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( NvCxnSpPr.class );
	private static final long serialVersionUID = -4808617992996239153L;
	private CNvPr cpr;
//	private cNvCxnSpPr sppr= null;

	public NvCxnSpPr( CNvPr cpr/*, cNvCxnSpPr sppr*/ )
	{
		this.cpr = cpr;
		//this.sppr= sppr;
	}

	public NvCxnSpPr( NvCxnSpPr n )
	{
		this.cpr = n.cpr;
		//this.sppr= n.sppr;
	}

	public static NvCxnSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		CNvPr cpr = null;
		;
//    	cNvCxnSpPr sppr= null;
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
						lastTag.push( tnm );
						cpr = (CNvPr) CNvPr.parseOOXML( xpp, lastTag );
/*		            } else if (tnm.equals("cNvCxnSpPr")) {		
	        			 lastTag.push(tnm);
	        			 sppr= (cNvCxnSpPr) cNvCxnSpPr.parseOOXML(xpp, lastTag).clone();		            	
*/
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "nvCxnSpPr" ) )
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
			log.error( "nvCxnSpPr.parseOOXML: " + e.toString() );
		}
		NvCxnSpPr n = new NvCxnSpPr( cpr/*, sppr*/ );
		return n;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:nvCxnSpPr>" );
		if( cpr != null )
		{
			ooxml.append( cpr.getOOXML() );
		}
		// TODO: finihs cNvCxnSpPr
		ooxml.append( "<xdr:cNvCxnSpPr/>" );
		ooxml.append( "</xdr:nvCxnSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NvCxnSpPr( this );
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( cpr != null )
		{
			return cpr.getName();
		}
		return null;
	}

	/**
	 * set cNvPr name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		if( cpr != null )
		{
			cpr.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( cpr != null )
		{
			return cpr.getDescr();
		}
		return null;
	}

	/**
	 * set cNvPr description attribute
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setDescr( String descr )
	{
		if( cpr != null )
		{
			cpr.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( cpr != null )
		{
			cpr.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( cpr != null )
		{
			return cpr.getId();
		}
		return -1;
	}
}


