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

import java.util.Stack;

/**
 * graphic (Graphic Object)
 * This element specifies the existence of a single graphic object. Document authors should refer to this element
 * when they wish to persist a graphical object of some kind. The specification for this graphical object will be
 * provided entirely by the document author and referenced within the graphicData child element
 * <p/>
 * parent:  anchor, graphicFrame, inline
 * children: graphicData
 */
public class Graphic implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Graphic.class );
	private static final long serialVersionUID = -7027946026352255398L;
	private GraphicData graphicData = new GraphicData();

	public Graphic()
	{
	}

	public Graphic( GraphicData g )
	{
		this.graphicData = g;
	}

	public Graphic( Graphic gr )
	{
		this.graphicData = gr.graphicData;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		GraphicData g = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "graphicData" ) )
					{
						lastTag.push( tnm );
						g = GraphicData.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "graphic" ) )
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
			log.error( "graphic.parseOOXML: " + e.toString() );
		}
		Graphic gr = new Graphic( g );
		return gr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:graphic>" );
		if( graphicData != null )
		{
			ooxml.append( graphicData.getOOXML() );
		}
		ooxml.append( "</a:graphic>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Graphic( this );
	}

	/**
	 * get the URI associated with this graphic Data
	 */
	public String getURI()
	{
		if( graphicData != null )
		{
			return graphicData.getURI();
		}
		return null;
	}

	public void setURI( String uri )
	{
		if( graphicData != null )
		{
			graphicData.setURI( uri );
		}
	}

	/**
	 * return the rid of the chart element, if exists
	 *
	 * @return
	 */
	public String getChartRId()
	{
		if( graphicData != null )
		{
			return graphicData.getChartRId();
		}
		return null;
	}

	/**
	 * set the rid of the chart element
	 */
	public void setChartRId( String rid )
	{
		if( graphicData != null )
		{
			graphicData.setChartRId( rid );
		}
	}
}

/**
 * graphicData (Graphic Object Data)
 * This element specifies the reference to a graphic object within the document. This graphic object is provided
 * entirely by the document authors who choose to persist this data within the document.
 * <p/>
 * parent: graphic
 * children: chart ... anything else?
 */
class GraphicData implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( GraphicData.class );
	private static final long serialVersionUID = 7395991759307532325L;
	private String uri = OOXMLConstants.chartns; //xmlns:r=\"" + OOXMLConstants.relns + "\"";	// default
	private String rid = null;

	public GraphicData()
	{
	}

	public GraphicData( String uri, String rid )
	{
		this.uri = uri;
		this.rid = rid;
	}

	public GraphicData( GraphicData gd )
	{
		this.uri = gd.uri;
		this.rid = gd.rid;
	}

	public static GraphicData parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		String uri = null;
		String rid = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "graphicData" ) )
					{        // get attributes
						if( xpp.getAttributeCount() > 0 )
						{
							uri = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "chart" ) )
					{        // one of many possible children
						if( xpp.getAttributeCount() > 0 )
						{
							rid = xpp.getAttributeValue( 0 );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "graphicData" ) )
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
			log.error( "graphicData.parseOOXML: " + e.toString() );
		}
		GraphicData g = new GraphicData( uri, rid );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:graphicData" );
		if( uri != null )
		{
			ooxml.append( " uri=\"" + uri + "\"" );
		}
		if( rid != null )
		{ // we'll assume it's a chart for nwo
			ooxml.append(
					"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"  r:id=\"" + rid + "\"/></a:graphicData>" );
		}
		else
		{
			ooxml.append( "/>" );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GraphicData( this );
	}

	/**
	 * return the URI attribute associated with this graphic Data
	 *
	 * @return
	 */
	public String getURI()
	{
		return uri;
	}

	/**
	 * set the URI attribute for this graphic data
	 *
	 * @param uri
	 */
	public void setURI( String uri )
	{
		this.uri = uri;
	}

	/**
	 * return the rid of the chart element, if exists
	 *
	 * @return
	 */
	public String getChartRId()
	{
		return rid;
	}

	/**
	 * set the rid of the chart element
	 */
	public void setChartRId( String rid )
	{
		this.rid = rid;
	}
}

