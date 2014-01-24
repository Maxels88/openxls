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

/**
 * bodyPr (Body Properties)
 * <p/>
 * This element defines the body properties for the text body within a shape
 * <p/>
 * parents:  many, including txBody and txPr
 * attributes:  many
 * children: flatTx, noAutoFit, normAutoFit, prstTxWarp, scene3d, sp3d, spAutoFit
 */
// TODO: Handle CHILDREN ***********************************
public class BodyPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( BodyPr.class );
	private static final long serialVersionUID = 3693893834015788452L;
	private HashMap<String, String> attrs = new HashMap<>();
	private PrstTxWarp txwarp = null;
	private boolean spAutoFit = false;

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<>();
		PrstTxWarp txwarp = null;
		boolean spAutoFit = false;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "bodyPr" ) )
					{        // body text properties
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
						// TODO: handle flatTx element
						// TODO: handle scene3d element
						// TODO: handle sp3d element
					}
					else if( tnm.equals( "spAutoFit" ) )
					{    // TODO: should be a choice of autofit options
						spAutoFit = true;    // no attributes or children
					}
					else if( tnm.equals( "prstTxWarp" ) )
					{
						lastTag.push( tnm );
						txwarp = PrstTxWarp.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "bodyPr" ) )
					{
						lastTag.pop();    // pop this tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "bodyPr.parseOOXML: " + e.toString() );
		}
		BodyPr bpr = new BodyPr( attrs, txwarp, spAutoFit );
		return bpr;
	}

	public BodyPr()
	{
	}

	public BodyPr( HashMap<String, String> attrs, PrstTxWarp txwarp, boolean spAutoFit )
	{
		this.attrs = attrs;
		this.txwarp = txwarp;
		this.spAutoFit = spAutoFit;
	}

	public BodyPr( BodyPr tpr )
	{
		attrs = tpr.attrs;
		txwarp = tpr.txwarp;
		spAutoFit = tpr.spAutoFit;
	}

	/**
	 * defines the body properties for the text body within a shape
	 *
	 * @param hrot
	 * @param vert Determines if the text within the given text body should be displayed vertically. If this attribute is omitted, then a value of horz, or no vertical text is implied.
	 *             vert Determines if all of the text is vertical orientation	(each line is 90 degrees rotated clockwise, so it goes from top to bottom; each next line is to the left from the previous one).
	 *             vert270 Determines if all of the text is vertical orientation (each line is 270 degrees rotated clockwise, so it goes from bottom to top; each next line is to the right from the previous one).
	 *             wordArtVert Determines if all of the text is vertical ("one letter on top of another").
	 *             wordArtVertRtl  Specifies that vertical WordArt should be shown from right to left rather than left to right.
	 *             eaVert  A special version of vertical text, where some fonts are displayed as if rotated by 90 degrees while some fonts (mostly East Asian) are displayed vertical.
	 */
	public BodyPr( int hrot, String vert )
	{
		attrs = new HashMap<>();
		attrs.put( "rot", String.valueOf( hrot ) );
		if( vert != null )
		{
			attrs.put( "vert", vert );
		}
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:bodyPr" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( txwarp != null )
		{
			ooxml.append( txwarp.getOOXML() );
		}
		if( spAutoFit )
		{
			ooxml.append( "<a:spAutoFit/>" );        // TODO: Should be a choice of autofit options
		}
		// scene3d
		// text3d choice
		ooxml.append( "</a:bodyPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new BodyPr( this );
	}

}

/**
 * prstTxWarp (Preset Text Warp)
 * <p/>
 * This element specifies when a preset geometric shape should be used to transform a piece of text. This
 * operation is known formally as a text warp. The generating application should be able to render all preset
 * geometries enumerated in the ST_TextShapeType list.
 * <p/>
 * parent:  bodyPr
 * children: avLst
 */
class PrstTxWarp implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( PrstTxWarp.class );
	private static final long serialVersionUID = -2627323317407321668L;
	private String prst = null;
	private AvLst av = null;

	public PrstTxWarp()
	{
	}

	public PrstTxWarp( String prst, AvLst av )
	{
		this.prst = prst;
		this.av = av;
	}

	public PrstTxWarp( PrstTxWarp p )
	{
		prst = p.prst;
		av = p.av;
	}

	public static PrstTxWarp parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String prst = null;
		AvLst av = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "prstTxWarp" ) )
					{        // prst is only attribute and is required
						prst = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "avLst" ) )
					{
						lastTag.push( tnm );
						av = (AvLst) AvLst.parseOOXML( xpp, lastTag );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "prstTxWarp" ) )
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
			log.error( "prstTxWarp.parseOOXML: " + e.toString() );
		}
		PrstTxWarp p = new PrstTxWarp( prst, av );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:prstTxWarp prst=\"" + prst + "\">" );
		if( av != null )
		{
			ooxml.append( av.getOOXML() );
		}
		ooxml.append( "</a:prstTxWarp>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PrstTxWarp( this );
	}
}

