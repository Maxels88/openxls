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
 * One of:  cxnSp, graphicFrame, grpSp, pic, sp
 */
public class ObjectChoice implements OOXMLElement
{

	private static final long serialVersionUID = 3548474869557092714L;
	private CxnSp cxnSp = null;
	private GraphicFrame graphicFrame = null;
	private GrpSp grpSp = null;
	private Pic pic = null;
	private Sp sp = null;

	public ObjectChoice()
	{

	}

	public ObjectChoice( CxnSp c, GraphicFrame g, GrpSp grp, Pic p, Sp s )
	{
		this.cxnSp = c;
		this.graphicFrame = g;
		this.grpSp = grp;
		this.pic = p;
		this.sp = s;
	}

	public ObjectChoice( ObjectChoice oc )
	{
		this.cxnSp = oc.cxnSp;
		this.graphicFrame = oc.graphicFrame;
		this.grpSp = oc.grpSp;
		this.pic = oc.pic;
		this.sp = oc.sp;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		CxnSp c = null;
		GraphicFrame g = null;
		GrpSp grp = null;
		Pic p = null;
		Sp s = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cxnSp" ) )
					{    // connection shape
						lastTag.push( tnm );
						c = (CxnSp) CxnSp.parseOOXML( xpp, lastTag, bk );
						break;
					}
					else if( tnm.equals( "graphicFrame" ) )
					{    // graphic data usually chart
						lastTag.push( tnm );
						g = (GraphicFrame) GraphicFrame.parseOOXML( xpp, lastTag );
						break;
					}
					else if( tnm.equals( "grpSp" ) )
					{    // group shape - combines one or more of sp/pic/graphicFrame/cxnSp
						lastTag.push( tnm );
						grp = (GrpSp) GrpSp.parseOOXML( xpp, lastTag, bk );
						break;
					}
					else if( tnm.equals( "sp" ) )
					{        // shape
						lastTag.push( tnm );
						s = (Sp) Sp.parseOOXML( xpp, lastTag, bk );
						break;
					}
					else if( tnm.equals( "pic" ) )
					{        // picture/image
						lastTag.push( tnm );
						p = (Pic) Pic.parseOOXML( xpp, lastTag, bk );
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "ObjectChoice.parseOOXML: " + e.toString() );
		}
		ObjectChoice o = new ObjectChoice( c, g, grp, p, s );
		return o;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( cxnSp != null )
		{
			ooxml.append( cxnSp.getOOXML() );
		}
		if( graphicFrame != null )
		{
			ooxml.append( graphicFrame.getOOXML() );
		}
		if( grpSp != null )
		{
			ooxml.append( grpSp.getOOXML() );
		}
		if( pic != null )
		{
			ooxml.append( pic.getOOXML() );
		}
		if( sp != null )
		{
			ooxml.append( sp.getOOXML() );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ObjectChoice( this );
	}

	/**
	 * return if this Object Choice refers to an image rather than a chart or shape
	 *
	 * @return
	 */
	public boolean hasImage()
	{
		// o will be a pic element or a group shape containing a pic element, it's blipFill.blip child references the rId of the embedded file  
		if( this.getEmbed() != null )
		{
			return true;
		}
		return false;
	}

	/**
	 * return if this Object Choice refers to a shape, as opposed a chart or an image
	 *
	 * @return
	 */
	public boolean hasShape()
	{
		return ((cxnSp != null) || (sp != null) || ((grpSp != null) && grpSp.hasShape()));
	}

	/**
	 * return if this Object Choice element refers to a chart as opposed to a shape or image
	 *
	 * @return
	 */
	public boolean hasChart()
	{
		return this.getChartRId() != null;
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( cxnSp != null )
		{
			return cxnSp.getName();
		}
		else if( sp != null )
		{
			return sp.getName();
		}
		else if( pic != null )
		{
			return pic.getName();
		}
		else if( graphicFrame != null )
		{
			return graphicFrame.getName();
		}
		else if( grpSp != null )
		{
			return grpSp.getName();
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
		if( cxnSp != null )
		{
			cxnSp.setName( name );
		}
		else if( sp != null )
		{
			sp.setName( name );
		}
		else if( pic != null )
		{
			pic.setName( name );
		}
		else if( graphicFrame != null )
		{
			graphicFrame.setName( name );
		}
		else if( grpSp != null )
		{
			grpSp.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( cxnSp != null )
		{
			return cxnSp.getDescr();
		}
		else if( sp != null )
		{
			return sp.getDescr();
		}
		else if( pic != null )
		{
			return pic.getDescr();
		}
		else if( graphicFrame != null )
		{
			return graphicFrame.getDescr();
		}
		else if( grpSp != null )
		{
			return grpSp.getDescr();
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
		if( cxnSp != null )
		{
			cxnSp.setDescr( descr );
		}
		else if( sp != null )
		{
			sp.setDescr( descr );
		}
		else if( pic != null )
		{
			pic.setDescr( descr );
		}
		else if( graphicFrame != null )
		{
			graphicFrame.setDescr( descr );
		}
		else if( grpSp != null )
		{
			grpSp.setDescr( descr );
		}
	}

	/**
	 * get macro attribute (valid for cnxSp, sp and graphicFrame)
	 *
	 * @return
	 */
	public String getMacro()
	{
		if( cxnSp != null )
		{
			return cxnSp.getMacro();
		}
		else if( graphicFrame != null )
		{
			return graphicFrame.getMacro();
		}
		else if( sp != null )
		{
			return sp.getMacro();
		}
		return null;
	}

	/**
	 * set Macro attribute (valid for cnxSp, sp and graphicFrame)
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setMacro( String macro )
	{
		if( cxnSp != null )
		{
			cxnSp.setMacro( macro );
		}
		else if( graphicFrame != null )
		{
			graphicFrame.setMacro( macro );
		}
		else if( sp != null )
		{
			sp.setMacro( macro );
		}
		else if( grpSp != null )
		{
			grpSp.setMacro( macro );
		}
	}

	/**
	 * get the URI associated with this graphic Data
	 */
	public String getURI()
	{
		if( graphicFrame != null )
		{
			return graphicFrame.getURI();
		}
		return null;
	}

	/**
	 * set the URI associated with this graphic data
	 *
	 * @param uri
	 */
	public void setURI( String uri )
	{
		if( graphicFrame != null )
		{
			graphicFrame.setURI( uri );
		}
		else if( grpSp != null )
		{
			grpSp.setURI( uri );
		}
	}

	/**
	 * return the rid for the embedded object (picture or picture shape) (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( sp != null )
		{
			return sp.getEmbed();    // embedded blip/pict
		}
		else if( pic != null )
		{
			return pic.getEmbed();    // embedded image
		}
		else if( grpSp != null )
		{
			return grpSp.getEmbed();    // group shape embedded image
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
		if( sp != null )
		{
			return sp.getLink();
		}
		else if( pic != null )
		{
			return pic.getLink();
		}
		else if( grpSp != null )
		{
			return grpSp.getLink();
		}
		return null;
	}

	/**
	 * set the rid for this object (picture or picture shape) (resides within the file)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( sp != null )
		{
			sp.setEmbed( embed );
		}
		else if( pic != null )
		{
			pic.setEmbed( embed );
		}
		else if( grpSp != null )
		{
			grpSp.setEmbed( embed );
		}
	}

	/**
	 * set the rid for this chart (resides within the file)
	 *
	 * @param rId
	 */
	public void setChartRId( String rId )
	{
		if( graphicFrame != null )
		{
			graphicFrame.setChartRId( rId );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param link
	 */
	public void setLink( String link )
	{
		if( sp != null )
		{
			sp.setLink( link );
		}
		else if( pic != null )
		{
			pic.setLink( link );
		}
		else if( grpSp != null )
		{
			grpSp.setLink( link );
		}
	}

	/**
	 * return the rid of the chart element, if exists
	 *
	 * @return
	 */
	public String getChartRId()
	{
		if( graphicFrame != null )
		{
			return graphicFrame.getChartRId();
		}
		else if( grpSp != null )
		{
			return grpSp.getChartRId();
		}
		return null;
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( sp != null )
		{
			sp.setId( id );    // embedded blip/pict
		}
		else if( pic != null )
		{
			pic.setId( id );    // embedded image
		}
		else if( graphicFrame != null )
		{
			graphicFrame.setId( id );    // chart
		}
		else if( cxnSp != null )
		{
			cxnSp.setId( id );
		}
		else if( grpSp != null )
		{
			grpSp.setId( id );    // embedded image
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( sp != null )
		{
			return sp.getId();    // embedded blip/pict
		}
		else if( pic != null )
		{
			return pic.getId();    // embedded image
		}
		else if( graphicFrame != null )
		{
			return graphicFrame.getId();    // chart
		}
		else if( cxnSp != null )
		{
			return cxnSp.getId();
		}
		else if( grpSp != null )
		{
			return grpSp.getId();    // embedded image
		}
		return -1;
	}

	/**
	 * utility to return the shape properties element for this picture
	 * should be depreciated when OOXML is completely distinct from BIFF8
	 *
	 * @return
	 */
	public SpPr getSppr()
	{
		if( pic != null )
		{
			return pic.getSppr();
		}
		else if( grpSp != null )
		{
			return grpSp.getSppr();
		}
		return null;
	}

	/**
	 * return the actual object associated with this ObjectChoice
	 *
	 * @return
	 */
	public Object getObject()
	{
		if( cxnSp != null )
		{
			return cxnSp;
		}
		else if( graphicFrame != null )
		{
			return graphicFrame;
		}
		else if( pic != null )
		{
			return pic;
		}
		else if( sp != null )
		{
			return sp;
		}
		else if( grpSp != null )
		{
			return grpSp;
		}
		return null;
	}

	/**
	 * set the object associated with this ObjectChoice
	 *
	 * @param o
	 */
	public void setObject( Object o )
	{
		if( o instanceof GraphicFrame )
		{
			graphicFrame = (GraphicFrame) o;
		}
		else if( o instanceof Pic )
		{
			pic = (Pic) o;
		}
		else if( o instanceof Sp )
		{
			sp = (Sp) o;
		}
		else if( o instanceof CxnSp )
		{
			cxnSp = (CxnSp) o;
		}
		else if( o instanceof GrpSp )
		{
			grpSp = (GrpSp) o;
		}
	}
}