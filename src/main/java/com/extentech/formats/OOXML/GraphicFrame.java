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

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * graphicFrame (Graphic Frame)
 * This element describes a single graphical object frame for a spreadsheet which contains a graphical object
 * <p/>
 * parent: oneCellAnchor, twoCellAnchor, absoluteAnchor, grpSp
 * children: graphic, nvGraphicFramePr, xfrm (all required and in sequence)
 */
//TODO: finish cNvGraphicFramePr.graphicFrameLocks element
public class GraphicFrame implements OOXMLElement
{

	private static final long serialVersionUID = 2494490998000511917L;
	private HashMap<String, String> attrs = new HashMap<String, String>();
	private Graphic graphic = new Graphic();

	private NvGraphicFramePr graphicFramePr = new NvGraphicFramePr();
	private Xfrm xfrm = new Xfrm();

	public GraphicFrame()
	{
		attrs.put( "macro", "" );
	}

	public GraphicFrame( HashMap<String, String> attrs, Graphic g, NvGraphicFramePr gfp, Xfrm x )
	{
		this.attrs = attrs;
		this.graphic = g;
		this.graphicFramePr = gfp;
		this.xfrm = x;
	}

	public GraphicFrame( GraphicFrame gf )
	{
		this.attrs = gf.attrs;
		this.graphic = gf.graphic;
		this.graphicFramePr = gf.graphicFramePr;
		this.xfrm = gf.xfrm;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		Graphic g = null;
		NvGraphicFramePr gfp = null;
		Xfrm x = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "graphicFrame" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "xfrm" ) )
					{
						lastTag.push( tnm );
						x = (Xfrm) Xfrm.parseOOXML( xpp, lastTag );
						x.setNS( "xdr" );
					}
					else if( tnm.equals( "graphic" ) )
					{
						lastTag.push( tnm );
						g = (Graphic) Graphic.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "nvGraphicFramePr" ) )
					{
						lastTag.push( tnm );
						gfp = (NvGraphicFramePr) NvGraphicFramePr.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "graphicFrame" ) )
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
			Logger.logErr( "graphicFrame.parseOOXML: " + e.toString() );
		}
		GraphicFrame gf = new GraphicFrame( attrs, g, gfp, x );
		return gf;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:graphicFrame" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		// all are required so no null checks - must ensure x.ns is set
		ooxml.append( graphicFramePr.getOOXML() );
		ooxml.append( xfrm.getOOXML() );
		ooxml.append( graphic.getOOXML() );
		ooxml.append( "</xdr:graphicFrame>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GraphicFrame( this );
	}

	/**
	 * get graphicFrame Macro attribute
	 */
	public String getMacro()
	{
		if( attrs.get( "macro" ) != null )
		{
			return (String) attrs.get( "macro" );
		}
		return null;
	}

	/**
	 * set graphicFrame Macro attribute
	 *
	 * @param macro
	 */
	public void setMacro( String macro )
	{
		attrs.put( "macro", macro );
	}

	/**
	 * get the URI associated with this graphic Data
	 */
	public String getURI()
	{
		if( graphic != null )
		{
			return graphic.getURI();
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
		if( graphic != null )
		{
			graphic.setURI( uri );
		}
	}

	/**
	 * return the rid of the chart element, if exists
	 *
	 * @return
	 */
	public String getChartRId()
	{
		if( graphic != null )
		{
			return graphic.getChartRId();
		}
		return null;
	}

	/**
	 * set the rid of the chart element
	 */
	public void setChartRId( String rid )
	{
		if( graphic != null )
		{
			graphic.setChartRId( rid );
		}
	}

	/**
	 * get the cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( graphicFramePr != null )
		{
			return graphicFramePr.getName();
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
		if( graphicFramePr != null )
		{
			graphicFramePr.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( graphicFramePr != null )
		{
			return graphicFramePr.getDescr();
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
		if( graphicFramePr != null )
		{
			graphicFramePr.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( graphicFramePr != null )
		{
			graphicFramePr.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( graphicFramePr != null )
		{
			return graphicFramePr.getId();
		}
		return -1;
	}
}

/**
 * nvGraphicFramePr (Non-Visual Properties for a Graphic Frame)
 * This element specifies all non-visual properties for a graphic frame. This element is a container for the non-visual
 * identification properties, shape properties and application properties that are to be associated with a graphic
 * frame. This allows for additional information that does not affect the appearance of the graphic frame to be
 * stored.
 * <p/>
 * parent: graphicFrame
 * children: cNvPr REQ cNvGraphicFramePr REQ
 */
class NvGraphicFramePr implements OOXMLElement
{
	private static final long serialVersionUID = -47476384268955296L;
	private CNvPr cp = new CNvPr();
	private CNvGraphicFramePr nvpr = new CNvGraphicFramePr();

	public NvGraphicFramePr()
	{
	}

	public NvGraphicFramePr( CNvPr cp, CNvGraphicFramePr nvpr )
	{
		this.cp = cp;
		this.nvpr = nvpr;
	}

	public NvGraphicFramePr( NvGraphicFramePr nvg )
	{
		this.cp = nvg.cp;
		this.nvpr = nvg.nvpr;
	}

	public static NvGraphicFramePr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		CNvPr cp = null;
		CNvGraphicFramePr nvpr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cNvGraphicFramePr" ) )
					{
						lastTag.push( tnm );
						nvpr = (CNvGraphicFramePr) CNvGraphicFramePr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "cNvPr" ) )
					{
						lastTag.push( tnm );
						cp = (CNvPr) CNvPr.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "nvGraphicFramePr" ) )
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
			Logger.logErr( "nvGraphicFramePr.parseOOXML: " + e.toString() );
		}
		NvGraphicFramePr gfp = new NvGraphicFramePr( cp, nvpr );
		return gfp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:nvGraphicFramePr>" );
		if( cp != null )
		{
			ooxml.append( cp.getOOXML() );
		}
		if( nvpr != null )
		{
			ooxml.append( nvpr.getOOXML() );
		}
		ooxml.append( "</xdr:nvGraphicFramePr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NvGraphicFramePr( this );
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( cp != null )
		{
			return cp.getName();
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
		if( cp != null )
		{
			cp.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( cp != null )
		{
			return cp.getDescr();
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
		if( cp != null )
		{
			cp.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( cp != null )
		{
			cp.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( cp != null )
		{
			return cp.getId();
		}
		return -1;
	}

}

/**
 * cNvGraphicFramePr (Non-Visual Graphic Frame Drawing Properties)
 * <p/>
 * This element specifies the non-visual properties for a single graphical object frame within a spreadsheet. These
 * are the set of properties of a frame which do not affect its display within a spreadsheet
 * <p/>
 * parent: nvGraphicFramePr
 * children: graphicFrameLocks
 */
// TODO: finish graphicFrameLocks
class CNvGraphicFramePr implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 769474804434194488L;

	// private graphicFrameLocks gf= null;
	public CNvGraphicFramePr()
	{
	}

	public CNvGraphicFramePr( CNvGraphicFramePr g )
	{
	}

	public static CNvGraphicFramePr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
/*		            String tnm = xpp.getName();
		            if (tnm.equals("graphicFrameLocks")) {
	        			 lastTag.push(tnm);
	        			 //gf= (graphicFrameLocks) graphicFrameLocks.parseOOXML(xpp, lastTag).clone();		            	
		            }
*/
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cNvGraphicFramePr" ) )
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
			Logger.logErr( "cNvGraphicFramePr.parseOOXML: " + e.toString() );
		}
		CNvGraphicFramePr cpr = new CNvGraphicFramePr();
		return cpr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:cNvGraphicFramePr>" );
		// TODO: Finish child graphicFrameLocks
		//if (gf!=null) ooxml.append(gf.getOOXML());
		ooxml.append( "<a:graphicFrameLocks/>" );
		ooxml.append( "</xdr:cNvGraphicFramePr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CNvGraphicFramePr( this );
	}

}

