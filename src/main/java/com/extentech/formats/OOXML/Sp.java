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
 * sp (Shape)
 * <p/>
 * OOXML/DrawingML element representing a single shape
 * A shape can either be a preset or a custom geometry,
 * defined using the DrawingML framework. In addition to a geometry each shape can have both visual and non
 * visual properties attached. Text and corresponding styling information can also be attached to a shape.
 * <p/>
 * parent = twoCellAnchor, oneCellAnchor, absoluteAnchor, grpSp (group shape)
 * children:  nvSpPr, spPr, style, txBody
 */
public class Sp implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Sp.class );
	private static final long serialVersionUID = 7454285931503575078L;
	private NvSpPr nvsp;
	private SpPr sppr;
	private Style sty;
	private TxBody txb;
	private HashMap<String, String> attrs = null;

	public Sp( NvSpPr nvsp, SpPr sppr, Style sty, TxBody txb, HashMap<String, String> attrs )
	{
		this.nvsp = nvsp;
		this.sppr = sppr;
		this.sty = sty;
		this.txb = txb;
		this.attrs = attrs;
	}

	public Sp( Sp shp )
	{
		nvsp = shp.nvsp;
		sppr = shp.sppr;
		sty = shp.sty;
		txb = shp.txb;
		attrs = shp.attrs;
	}

	/**
	 * return the OOXML specific for this object
	 *
	 * @param xpp
	 * @param lastTag
	 * @return
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		NvSpPr nvsp = null;
		SpPr sppr = null;
		Style sty = null;
		TxBody txb = null;
		HashMap<String, String> attrs = new HashMap<>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "sp" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "nvSpPr" ) )
					{        // non-visual shape props
						lastTag.push( tnm );
						nvsp = NvSpPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "spPr" ) )
					{        // shape properties
						lastTag.push( tnm );
						sppr = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
						//sppr.setNS("xdr");
					}
					else if( tnm.equals( "style" ) )
					{        // shape style
						lastTag.push( tnm );
						sty = (Style) Style.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "txBody" ) )
					{        // shape text body
						lastTag.push( tnm );
						txb = (TxBody) TxBody.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "sp" ) )
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
			log.error( "sp.parseOOXML: " + e.toString() );
		}
		Sp shp = new Sp( nvsp, sppr, sty, txb, attrs );
		return shp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:sp" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		ooxml.append( nvsp.getOOXML() );
		ooxml.append( sppr.getOOXML() );
		if( sty != null )
		{
			ooxml.append( sty.getOOXML() );
		}
		if( txb != null )
		{
			ooxml.append( txb.getOOXML() );
		}
		ooxml.append( "</xdr:sp>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Sp( this );
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
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( sppr != null )
		{
			return sppr.getEmbed();
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
		if( sppr != null )
		{
			return sppr.getLink();
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
		if( sppr != null )
		{
			sppr.setEmbed( embed );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( sppr != null )
		{
			sppr.setLink( link );
		}
	}

	/**
	 * return the name of this shape, if any
	 *
	 * @return
	 */
	public String getName()
	{
		if( nvsp != null )
		{
			return nvsp.getName();
		}
		return null;
	}

	/**
	 * set the name of this shape, if any
	 */
	public void setName( String name )
	{
		if( nvsp != null )
		{
			nvsp.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( nvsp != null )
		{
			return nvsp.getDescr();
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
		if( nvsp != null )
		{
			nvsp.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( nvsp != null )
		{
			nvsp.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( nvsp != null )
		{
			return nvsp.getId();
		}
		return -1;
	}

}

/**
 * This element specifies all non-visual properties for a shape. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a shape.
 * This allows for additional information that does not affect the appearance of the shape to be stored
 * <p/>
 * parent:  sp
 * children: cNvPr  REQ cNvSpPr REQ
 */
class NvSpPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( NvSpPr.class );
	private static final long serialVersionUID = 9121235009516398367L;
	private CNvPr cnv = null;
	private CNvSpPr cnvsp = null;

	public NvSpPr( CNvPr cnv, CNvSpPr cnvsp )
	{
		this.cnv = cnv;
		this.cnvsp = cnvsp;
	}

	public NvSpPr( NvSpPr nvsp )
	{
		cnv = nvsp.cnv;
		cnvsp = nvsp.cnvsp;
	}

	public static NvSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		CNvPr cnv = null;
		CNvSpPr cnvsp = null;
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
						cnv = (CNvPr) CNvPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "cNvSpPr" ) )
					{
						lastTag.push( tnm );
						cnvsp = CNvSpPr.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "nvSpPr" ) )
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
			log.error( "txBody.parseOOXML: " + e.toString() );
		}
		NvSpPr nvp = new NvSpPr( cnv, cnvsp );
		return nvp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:nvSpPr>" );
		ooxml.append( cnv.getOOXML() );
		ooxml.append( cnvsp.getOOXML() );
		ooxml.append( "</xdr:nvSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NvSpPr( this );
	}

	/**
	 * get name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( cnv != null )
		{
			return cnv.getName();
		}
		return null;
	}

	/**
	 * set name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		if( cnv != null )
		{
			cnv.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( cnv != null )
		{
			return cnv.getDescr();
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
		if( cnv != null )
		{
			cnv.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( cnv != null )
		{
			cnv.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( cnv != null )
		{
			return cnv.getId();
		}
		return -1;
	}
}

/**
 * cNvSpPr (Non-visual Drawing Properties for a shape
 * <p/>
 * This OOXML/DrawingML element specifies the non-visual drawing properties for a shape.
 * These properties are to be used by the generating application to determine how the shape should be dealt with
 * <p/>
 * attributes:  txBox (optional boolean)
 * parent:    nvSpPr
 * children:  spLocks (shapeLocks)
 */
class CNvSpPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CNvSpPr.class );
	private static final long serialVersionUID = 7895953516797436713L;
	private String txBox = null;
	private SpLocks sp = null;

	public CNvSpPr( String t, SpLocks sp )
	{
		txBox = t;
		this.sp = sp;
	}

	public CNvSpPr( CNvSpPr c )
	{
		txBox = c.txBox;
		sp = c.sp;
	}

	public static CNvSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String txBox = null;
		SpLocks sp = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cNvSpPr" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							if( xpp.getAttributeName( i ).equals( "txBox" ) )
							{
								txBox = xpp.getAttributeValue( i );
							}
						}
					}
					else if( tnm.equals( "spLocks" ) )
					{
						lastTag.push( tnm );
						sp = SpLocks.parseOOXML( xpp, lastTag );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cNvSpPr" ) )
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
			log.error( "cNvSpPr.parseOOXML: " + e.toString() );
		}
		CNvSpPr cnv = new CNvSpPr( txBox, sp );
		return cnv;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:cNvSpPr" );
		if( txBox != null )
		{
			ooxml.append( " txBox=\"" + txBox + "\"" );
		}
		ooxml.append( ">" );
		if( sp != null )
		{
			ooxml.append( sp.getOOXML() );
		}
		ooxml.append( "</xdr:cNvSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CNvSpPr( this );
	}
}

/**
 * spLocks (Shape Locks)
 * <p/>
 * OOXML/DrawingML element specifies all locking properties for a shape.
 * These properties inform the generating application
 * about specific properties that have been previously locked and thus should not be changed.
 * <p/>
 * parents:    cNvSpPr
 */
class SpLocks implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( SpLocks.class );
	private static final long serialVersionUID = -3805557220039550941L;
	private HashMap<String, String> attrs = null;

	public SpLocks( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public SpLocks( SpLocks sp )
	{
		attrs = sp.attrs;
	}

	public static SpLocks parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "spLocks" ) )
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
					if( endTag.equals( "spLocks" ) )
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
			log.error( "spLocks.parseOOXML: " + e.toString() );
		}
		SpLocks sp = new SpLocks( attrs );
		return sp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:spLocks" );
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
		return new SpLocks( this );
	}
}


