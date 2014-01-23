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

import java.util.Stack;

/**
 * OOXML Legend element
 * <p/>
 * parent:  chart
 * children:	layout, legendEntry, legendPos, overlay, spPr, txPr
 */
public class Legend implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Legend.class );
	private static final long serialVersionUID = 419453456635220517L;
	private String legendpos;    // actually an element but only contains 1 attribute: val
	private LegendEntry le;
	private Layout layout;
	private SpPr shapeProps;                    // defines the shape properties for the legend for this chart
	private TxPr txpr;                        // defines text properties
	private String overlay;

	public Legend( String pos, String overlay, Layout l, LegendEntry le, SpPr sp, TxPr txpr )
	{
		this.legendpos = pos;
		this.le = le;
		this.overlay = overlay;
		this.layout = l;
		this.shapeProps = sp;
		this.txpr = txpr;
	}

	public Legend( Legend l )
	{
		this.legendpos = l.legendpos;
		this.le = l.le;
		this.overlay = l.overlay;
		this.layout = l.layout;
		this.shapeProps = l.shapeProps;
		this.txpr = l.txpr;
	}

	/**
	 * parse OOXML legend element
	 *
	 * @param xpp
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String legendpos = null;
		LegendEntry le = null;
		String overlay = null;
		Layout layout = null;
		SpPr shapeProps = null;                    // defines the shape properties for the legend for this chart
		TxPr txpr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "legendPos" ) )
					{
						legendpos = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "layout" ) )
					{
						lastTag.push( tnm );
						layout = (Layout) Layout.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "legendEntry" ) )
					{
						lastTag.push( tnm );
						le = LegendEntry.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						shapeProps = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
						//shapeProps.setNS("c");
					}
					else if( tnm.equals( "txPr" ) )
					{
						lastTag.push( tnm );
						txpr = (TxPr) TxPr.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "overlay" ) )
					{
						overlay = xpp.getAttributeValue( 0 );
					}

				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "legend" ) )
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
			log.error( "OOXMLAdapter.parseLegendElement: " + e.toString() );
		}
		Legend l = new Legend( legendpos, overlay, layout, le, shapeProps, txpr );
		return l;
	}

	/**
	 * fill 2003-style legend from OOXML Legend
	 *
	 * @param l_2003
	 */
	public void fill2003Legend( com.extentech.formats.XLS.charts.Legend l_2003 )
	{
		// 0= bottom, 1= corner, 2= top, 3= right, 4= left, 7= not docked
		String[] pos = { "b", "tr", "t", "r", "l" };
		for( int i = 0; i < pos.length; i++ )
		{
			if( pos[i].equals( legendpos ) )
			{
				l_2003.setLegendPosition( (short) i );
				break;
			}
		}
		if( this.hasBox() )
		{
			l_2003.addBox();
		}

	}

	/**
	 * generate the ooxml necessary to display chart legend
	 *
	 * @return
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:legend>" );
		// sequence
		if( legendpos != null )
		{
			ooxml.append( "<c:legendPos val=\"" + legendpos + "\"/>" );
		}
		if( le != null )
		{
			ooxml.append( le.getOOXML() );
		}
		if( layout != null )
		{
			ooxml.append( layout.getOOXML() );
		}
		if( overlay != null )
		{
			ooxml.append( "<c:overlay val=\"" + overlay + "\"/>" );
		}
		if( shapeProps != null )
		{
			ooxml.append( shapeProps.getOOXML() );
		}
		if( txpr != null )
		{
			ooxml.append( txpr.getOOXML() );
		}
		ooxml.append( "</c:legend>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Legend( this );
	}

	/**
	 * returns true if this legend should have a box around it
	 *
	 * @return
	 */
	public boolean hasBox()
	{
		if( this.shapeProps != null )
		{
			return this.shapeProps.hasLine();
		}
		return false;
	}

	/**
	 * create an OOXML legend from a 2003-vers legend
	 *
	 * @param l
	 */
	public static Legend createLegend( com.extentech.formats.XLS.charts.Legend l )
	{

		Legend ooxmllegend = null;
		try
		{
			SpPr sp = null;
			sp = new SpPr( "c" );
			sp.setLine( 3175, "000000" );
			l.getFnt();
			ooxmllegend = new Legend( l.getLegendPositionString(), "1", null, null, sp, null );
		}
		catch( Exception e )
		{
			log.warn( "Error creating 2007+ version Legend: " + e.toString(), e );
		}
		return ooxmllegend;
	}
}

/**
 * legend Entry element
 * <p/>
 * parent: legend
 * attributes;  val= position of legend
 * children:  idx, choice of: delete or txPr
 */
class LegendEntry implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( LegendEntry.class );
	private static final long serialVersionUID = 1859347855337611982L;
	private TxPr tx;
	private int idx = -1;
	private boolean delete;

	public LegendEntry( int idx, boolean d, TxPr tx )
	{
		this.idx = idx;
		this.delete = d;
		this.tx = tx;
	}

	public LegendEntry( LegendEntry le )
	{
		this.idx = le.idx;
		this.delete = le.delete;
		this.tx = le.tx;
	}

	public static LegendEntry parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		TxPr tx = null;
		int idx = -1;
		boolean delete = true;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "idx" ) )
					{
						idx = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "delete" ) )
					{
						delete = Boolean.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "txPr" ) )
					{
						lastTag.push( tnm );
						tx = (TxPr) TxPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "legendEntry" ) )
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
			log.error( "legendEntry.parseOOXML: " + e.toString() );
		}
		LegendEntry le = new LegendEntry( idx, delete, tx );
		return le;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<c:legendEntry>" );
		tooxml.append( "<c:idx val=\"" + idx + "\"/>" );
		if( !delete )
		{
			tooxml.append( "<c:delete=\"" + delete + "\">" );
		}
		if( tx != null )
		{
			tooxml.append( tx.getOOXML() );
		}
		tooxml.append( "</c:legendEntry>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new LegendEntry( this );
	}
}
