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
 * pic (Picture)
 * This element specifies the existence of a picture object within the spreadsheet.
 * <p/>
 * parent:  absoluteAnchor, grpSp, oneCellAnchor, twoCellAnchor
 * children: nvPicPr, blipFill, spPr, style
 */
//TODO: handle nvPicPr.cNvPicPr.picLocks element
public class Pic implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Pic.class );
	private static final long serialVersionUID = -4929177274389163606L;
	private HashMap<String, String> attrs;
	private NvPicPr nvPicPr;
	private BlipFill blipFill;
	private SpPr spPr;
	private Style style;

	public Pic()
	{        // set common defaults
		nvPicPr = new NvPicPr();
		blipFill = new BlipFill();
		spPr = new SpPr( "xdr" );
		spPr.setNS( "xdr" );
		attrs = null;
	}

	public Pic( HashMap<String, String> attrs, NvPicPr nv, BlipFill bf, SpPr sp, Style s )
	{
		this.attrs = attrs;
		this.nvPicPr = nv;
		this.blipFill = bf;
		this.spPr = sp;
		this.style = s;
	}

	public Pic( Pic p )
	{
		this.attrs = p.attrs;
		this.nvPicPr = p.nvPicPr;
		this.blipFill = p.blipFill;
		this.spPr = p.spPr;
		this.style = p.style;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<>();
		NvPicPr nv = null;
		BlipFill bf = null;
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
					if( tnm.equals( "pic" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "nvPicPr" ) )
					{
						lastTag.push( tnm );
						nv = NvPicPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "blipFill" ) )
					{
						lastTag.push( tnm );
						bf = BlipFill.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
						//sp.setNS("xdr");
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
					if( endTag.equals( "pic" ) )
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
			log.error( "pic.parseOOXML: " + e.toString() );
		}
		Pic p = new Pic( attrs, nv, bf, sp, s );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:pic" );
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
		if( nvPicPr != null )
		{
			ooxml.append( nvPicPr.getOOXML() );
		}
		if( blipFill != null )
		{
			ooxml.append( blipFill.getOOXML() );
		}
		if( spPr != null )
		{
			ooxml.append( spPr.getOOXML() );
		}
		if( style != null )
		{
			ooxml.append( style.getOOXML() );
		}
		ooxml.append( "</xdr:pic>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Pic( this );
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( blipFill != null )
		{
			return blipFill.getEmbed();
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
		if( blipFill != null )
		{
			return blipFill.getLink();
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
		if( blipFill != null )
		{
			blipFill.setEmbed( embed );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( blipFill != null )
		{
			blipFill.setLink( link );
		}
	}

	/**
	 * return the name of this shape, if any
	 *
	 * @return
	 */
	public String getName()
	{
		if( nvPicPr != null )
		{
			return nvPicPr.getName();
		}
		return null;
	}

	/**
	 * set the name of this shape, if any
	 */
	public void setName( String name )
	{
		if( nvPicPr != null )
		{
			nvPicPr.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( nvPicPr != null )
		{
			return nvPicPr.getDescr();
		}
		return null;
	}

	/**
	 * set cNvPr desc attribute
	 *
	 * @param name
	 */
	public void setDescr( String descr )
	{
		if( nvPicPr != null )
		{
			nvPicPr.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( nvPicPr != null )
		{
			nvPicPr.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( nvPicPr == null )
		{
			nvPicPr = new NvPicPr();
		}
		return nvPicPr.getId();
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
	 * add a line to this image
	 *
	 * @param w   int line width in
	 * @param clr String color html string
	 * @return
	 */
	public void setLine( int w, String clr )
	{
		if( spPr != null )
		{
			spPr.setLine( w, clr );
		}
	}

	/**
	 * utility to return the shape properties element for this picture
	 * should be depreciated when OOXML is completely distinct from BIFF8
	 *
	 * @return
	 */
	public SpPr getSppr()
	{
		return spPr;
	}

	/**
	 * utility to return the shape properties element for this picture
	 * should be depreciated when OOXML is completely distinct from BIFF8
	 *
	 * @return
	 */
	public void setSppr( SpPr sp )
	{
		this.spPr = sp;
	}
}

/**
 * nvPicPr (Non-Visual Properties for a Picture)
 * This element specifies all non-visual properties for a picture. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a picture.
 * This allows for additional information that does not affect the appearance of the picture to be stored.
 * <p/>
 * parent: pic
 * children: cNvPr, cNvPicPr
 */
class NvPicPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( NvPicPr.class );
	private static final long serialVersionUID = -3722424348721713313L;
	private CNvPr cpr;
	private CNvPicPr ppr;

	public NvPicPr()
	{    // set common defaults
		this.cpr = new CNvPr();
		this.ppr = new CNvPicPr();
	}

	public NvPicPr( CNvPr cpr, CNvPicPr ppr )
	{
		this.cpr = cpr;
		this.ppr = ppr;
	}

	public NvPicPr( NvPicPr n )
	{
		this.cpr = n.cpr;
		this.ppr = n.ppr;
	}

	public static NvPicPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		CNvPr cpr = null;
		CNvPicPr ppr = null;
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
					}
					else if( tnm.equals( "cNvPicPr" ) )
					{
						lastTag.push( tnm );
						ppr = CNvPicPr.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "nvPicPr" ) )
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
			log.error( "nvPicPr.parseOOXML: " + e.toString() );
		}
		NvPicPr n = new NvPicPr( cpr, ppr );
		return n;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:nvPicPr>" );
		ooxml.append( cpr.getOOXML() );
		ooxml.append( ppr.getOOXML() );
		ooxml.append( "</xdr:nvPicPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NvPicPr( this );
	}

	/**
	 * return the name of this shape, if any
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
	 * set the name of this shape, if any
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
	 * set cNvPr desc attribute
	 *
	 * @param name
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

/**
 * cNvPicPr (Non-Visual Picture Drawing Properties)
 * This element describes the non-visual properties of a picture within a spreadsheet. These are the set of
 * properties of a picture which do not affect its display within a spreadsheet.
 */
// TODO: handle child picLocks
class CNvPicPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CNvPicPr.class );
	private static final long serialVersionUID = 3690228348761065940L;
	private String preferRelativeResize = null;

	public CNvPicPr()
	{

	}

	public CNvPicPr( String preferRelativeResize )
	{
		this.preferRelativeResize = preferRelativeResize;
	}

	public CNvPicPr( CNvPicPr c )
	{
		this.preferRelativeResize = c.preferRelativeResize;
	}

	public static CNvPicPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String preferRelativeResize = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cNvPicPr" ) )
					{        // get attributes
						if( xpp.getAttributeCount() > 0 )
						{
							preferRelativeResize = xpp.getAttributeValue( 0 );
						}
/* TODO: Finish		} else if (tnm.equals("picLocks")) {
	        			 lastTag.push(tnm);
	        			 //p = (picLocks) picLocks.parseOOXML(xpp, lastTag).clone();
*/
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cNvPicPr" ) )
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
			log.error( "cNvPicPr.parseOOXML: " + e.toString() );
		}
		CNvPicPr c = new CNvPicPr( preferRelativeResize );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>" );
		// TODO: finish picLocks
		//if (p!=null) ooxml.append(p.getOOXML());
		//ooxml.append("</>");
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CNvPicPr( this );
	}
}

