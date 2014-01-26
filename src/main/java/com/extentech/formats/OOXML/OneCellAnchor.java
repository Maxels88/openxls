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
 * oneCellAnchor (One Cell Anchor Shape Size)
 * <p/>
 * This element specifies a one cell anchor placeholder for a group, a shape, or a drawing element. It moves with
 * the cell and its extents is in EMU units.
 * <p/>
 * parent: wsDr
 * children: from, ext, OBJECTCHOICES (sp, grpSp, graphicFrame, cxnSp, pic), clientData
 */
//TODO: finish grpSp Group Shape
// TODO: finish clientData element
public class OneCellAnchor implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( OneCellAnchor.class );
	private static final long serialVersionUID = -8498556079325357165L;
	public static final short EMU = 1270;
	private From from;
	private Ext ext;
	private ObjectChoice objectChoice;

	public OneCellAnchor( From f, Ext e, ObjectChoice o )
	{
		from = f;
		ext = e;
		objectChoice = o;
	}

	public OneCellAnchor( OneCellAnchor oca )
	{
		from = oca.from;
		ext = oca.ext;
		objectChoice = oca.objectChoice;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		From f = null;
		Ext e = null;
		ObjectChoice o = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "from" ) )
					{
						lastTag.push( tnm );
						f = From.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "ext" ) )
					{
						lastTag.push( tnm );
						e = (Ext) Ext.parseOOXML( xpp, lastTag ).cloneElement();
						//e.setNS("xdr");
					}
					else if( tnm.equals( "cxnSp" ) ||    // connection shape
							tnm.equals( "graphicFrame" ) ||
							tnm.equals( "grpSp" ) ||    // group shape
							tnm.equals( "pic" ) ||    // picture
							tnm.equals( "sp" ) )
					{        // shape
						lastTag.push( tnm );
						o = (ObjectChoice) ObjectChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "oneCellAnchor" ) )
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
			log.error( "oneCellAnchor.parseOOXML: " + ex.toString() );
		}
		OneCellAnchor oca = new OneCellAnchor( f, e, o );
		return oca;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:oneCellAnchor>" );
		if( from != null )
		{
			ooxml.append( from.getOOXML() );
		}
		ooxml.append( ext.getOOXML() );
		ooxml.append( objectChoice.getOOXML() );
		ooxml.append( "<xdr:clientData/>" );
		ooxml.append( "</xdr:oneCellAnchor>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new OneCellAnchor( this );
	}

	// access methods ******

	/**
	 * return the bounds of this object
	 * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST
	 *
	 * @return bounds short[4]
	 */
	public short[] getBounds()
	{
		short[] bounds = new short[8];
		System.arraycopy( from.getBounds(), 0, bounds, 0, 4 ); // from bounds
		return bounds;
	}

	/**
	 * set the bounds of this object
	 * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST
	 *
	 * @param bounds short[4]
	 */
	public void setBounds( int[] bounds )
	{
		int[] b = new int[4];
		System.arraycopy( bounds, 0, b, 0, 4 );
		if( from == null )
		{
			from = new From( b );
		}
		else
		{
			from.setBounds( b );
		}
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( objectChoice != null )
		{
			return objectChoice.getName();
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
		if( objectChoice != null )
		{
			objectChoice.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( objectChoice != null )
		{
			return objectChoice.getDescr();
		}
		return null;
	}

	/**
	 * set cNvPr descr attribute
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setDescr( String descr )
	{
		if( objectChoice != null )
		{
			objectChoice.setDescr( descr );
		}
	}

	/**
	 * get macro attribute
	 *
	 * @return
	 */
	public String getMacro()
	{
		if( objectChoice != null )
		{
			return objectChoice.getMacro();
		}
		return null;
	}

	/**
	 * set Macro attribute
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setMacro( String macro )
	{
		if( objectChoice != null )
		{
			objectChoice.setMacro( macro );
		}
	}

	/**
	 * get the URI associated with this graphic Data
	 */
	public String getURI()
	{
		if( objectChoice != null )
		{
			return objectChoice.getURI();
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
		if( objectChoice != null )
		{
			objectChoice.setURI( uri );
		}
	}

	/**
	 * return the id for the embedded picture, shape or chart (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( objectChoice != null )
		{
			return objectChoice.getEmbed();
		}
		return null;
	}

	/**
	 * return the id for the linked object (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( objectChoice != null )
		{
			return objectChoice.getLink();
		}
		return null;
	}

	/**
	 * set the embed or rId attribute for the embedded picture, shape or chart (i.e. resides within the file)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( objectChoice != null )
		{
			objectChoice.setEmbed( embed );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( objectChoice != null )
		{
			objectChoice.setLink( link );
		}
	}

	/**
	 * return if this oneCellAnchor element refers to an image rather than a chart or shape
	 *
	 * @return
	 */
	public boolean hasImage()
	{
		if( objectChoice != null )    // o will be a pic element, it's blipFill.blip child references the rId of the embedded file
		{
			return ((objectChoice.getObject() instanceof Pic) && (objectChoice.getEmbed() != null));
		}
		return false;
	}

	/**
	 * utility to return the shape properties element (picture element only)
	 * should be depreciated when OOXML is moved into ImageHandle
	 *
	 * @return
	 */
	public SpPr getSppr()
	{
		if( objectChoice != null )
		{
			return objectChoice.getSppr();
		}
		return null;
	}

	/**
	 * return if this oneCellAnchor element refers to a chart as opposed to a shape or image
	 *
	 * @return
	 */
	public boolean hasChart()
	{
		if( objectChoice != null )
		{
			return ((objectChoice.getObject() instanceof GraphicFrame) && (objectChoice.getChartRId() != null));
		}
		return false;
	}

	/**
	 * return if this oneCellAnchor element refers to a shape, as opposed a chart or an image
	 *
	 * @return
	 */
	public boolean hasShape()
	{
		if( objectChoice != null )
		{
			return ((objectChoice.getObject() instanceof CxnSp) || (objectChoice.getObject() instanceof Sp));
		}
		return false;
	}

	/**
	 * set this oneCellAnchor as a chart element
	 *
	 * @param rid
	 * @param name
	 * @param bounds
	 */
	public void setAsChart( int rid, String name, int[] bounds )
	{
		objectChoice = new ObjectChoice();
		objectChoice.setObject( new GraphicFrame() );
		objectChoice.setName( name );
		objectChoice.setEmbed( "rId" + Integer.valueOf( rid ).toString() );
		objectChoice.setId( rid );
		setBounds( bounds );
		// id???
	}

	/**
	 * set this oneCellAnchor as an image
	 *
	 * @param rid
	 * @param name
	 * @param id
	 */
	public static void setAsImage( String rid, String name, String id )
	{
		ObjectChoice o = new ObjectChoice();
		o.setObject( new Pic() );
		o.setName( name );
		o.setEmbed( rid );
	}
}

