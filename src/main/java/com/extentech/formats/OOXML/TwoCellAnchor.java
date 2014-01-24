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

import com.extentech.ExtenXLS.ColHandle;
import com.extentech.ExtenXLS.RowHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * twoCellAnchor (Two Cell Anchor Shape Size)
 * <p/>
 * This element specifies a two cell anchor placeholder for a group, a shape, or a drawing element.
 * It moves with cells and its extents are in EMU units.
 * <p/>
 * This is the root element for charts, images and shapes.
 * <p/>
 * parent: wsDr
 * children: from, to, OBJECTCHOICES (sp, grpSp, graphicFrame, cxnSp, pic), clientData
 */
//TODO: finish grpSp Group Shape
// TODO: finish clientData element
public class TwoCellAnchor implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( TwoCellAnchor.class );
	private static final long serialVersionUID = 4180396678197959710L;
	// EMU = pixel * 914400 / Resolution (96?)
	public static final short EMU = 1270;
	private String editAs = null;
	private String embedName = null;
	private From from = null;
	private To to = null;
	private ObjectChoice o = null;

	public TwoCellAnchor( String editAs )
	{
		this.editAs = editAs;
	}

	public TwoCellAnchor( String editAs, From f, To t, ObjectChoice o )
	{
		this.editAs = editAs;
		from = f;
		to = t;
		this.o = o;
	}

	public TwoCellAnchor( TwoCellAnchor tce )
	{
		editAs = tce.editAs;
		from = tce.from;
		to = tce.to;
		o = tce.o;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String editAs = null;
		From f = null;
		To t = null;
		ObjectChoice o = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "twoCellAnchor" ) )
					{        // get attributes
						if( xpp.getAttributeCount() > 0 )
						{
							editAs = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "from" ) )
					{
						lastTag.push( tnm );
						f = From.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "to" ) )
					{
						lastTag.push( tnm );
						t = To.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "cxnSp" ) ||
							tnm.equals( "graphicFrame" ) ||
							tnm.equals( "grpSp" ) ||
							tnm.equals( "pic" ) ||
							tnm.equals( "sp" ) )
					{
						o = (ObjectChoice) ObjectChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "twoCellAnchor" ) )
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
			log.error( "twoCellAnchor.parseOOXML: " + e.toString() );
		}
		TwoCellAnchor tca = new TwoCellAnchor( editAs, f, t, o );
		return tca;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:twoCellAnchor" );
		if( editAs != null )
		{
			ooxml.append( " editAs=\"" + editAs + "\"" );
		}
		ooxml.append( ">" );
		if( from != null )
		{
			ooxml.append( from.getOOXML() );
		}
		if( to != null )
		{
			ooxml.append( to.getOOXML() );
		}
		ooxml.append( o.getOOXML() );
		ooxml.append( "<xdr:clientData/>" );
		ooxml.append( "</xdr:twoCellAnchor>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new TwoCellAnchor( this );
	}

	// access methods ******

	/**
	 * return the (to, from) bounds of this object
	 * by concatenating the bounds for the to and the from
	 * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
	 *
	 * @return bounds short[8]  from [4] and to [4]
	 */
	public int[] getBounds()
	{
		int[] bounds = new int[8];
		System.arraycopy( from.getBounds(), 0, bounds, 0, 4 ); // from bounds
		System.arraycopy( to.getBounds(), 0, bounds, 4, 4 );  // to bounds
		return bounds;
	}

	/**
	 * set the (to, from) bounds of this object
	 * bounds represent:  COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
	 * NOTE: COL
	 *
	 * @param bounds int[8]  from [4] and to [4]
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
		System.arraycopy( bounds, 4, b, 0, 4 );
		if( to == null )
		{
			to = new To( b );
		}
		else
		{
			to.setBounds( b );
		}
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( o != null )
		{
			return o.getName();
		}
		return null;
	}

	public String toString()
	{
		return getName();
	}

	/**
	 * return if this twoCellAnchor element refers to an image rather than a chart or shape
	 *
	 * @return
	 */
	public boolean hasImage()
	{
		return o.hasImage();
	}

	/**
	 * return if this twoCellAnchor element refers to a chart as opposed to a shape or image
	 *
	 * @return
	 */
	public boolean hasChart()
	{
		if( o != null )
		{
			return o.hasChart();
		}
		return false;
	}

	/**
	 * return if this twoCellAnchor element refers to a shape, as opposed a chart or an image
	 *
	 * @return
	 */
	public boolean hasShape()
	{
		if( o != null )
		{
			return o.hasShape();
		}
		return false;
	}

	/**
	 * set cNvPr name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		if( o != null )
		{
			o.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( o != null )
		{
			return o.getDescr();
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
		if( o != null )
		{
			o.setDescr( descr );
		}
	}

	/**
	 * get macro attribute
	 *
	 * @return
	 */
	public String getMacro()
	{
		if( o != null )
		{
			return o.getMacro();
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
		if( o != null )
		{
			o.setMacro( macro );
		}
	}

	/**
	 * get the URI associated with this graphic Data
	 */
	public String getURI()
	{
		if( o != null )
		{
			return o.getURI();
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
		if( o != null )
		{
			o.setURI( uri );
		}
	}

	/**
	 * return the rid for the chart defined by this twocellanchor
	 *
	 * @return
	 */
	public String getChartRId()
	{
		if( o != null )
		{
			return o.getChartRId();
		}
		return null;
	}

	public void setChartRId( String rId )
	{
		if( o != null )
		{
			o.setChartRId( rId );
		}
	}

	/**
	 * return the id for the embedded picture or shape (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( o != null )
		{
			return o.getEmbed();
		}
		return null;
	}

	/**
	 * set the embed for the embedded picture or shape (i.e. resides within the file)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( o != null )
		{
			o.setEmbed( embed );
		}
	}

	/**
	 * return the Embedded Object's filename as saved on disk
	 *
	 * @return
	 */
	public String getEmbedFilename()
	{
		return embedName;
	}

	/**
	 * set the Embedded Object's filename as saved on disk
	 *
	 * @param embed
	 */
	public void setEmbedFilename( String embedFile )
	{
		embedName = embedFile;
	}

	/**
	 * return the id for the linked object (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( o != null )
		{
			return o.getLink();
		}
		return null;
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( o != null )
		{
			o.setLink( link );
		}
	}

	/**
	 * return the editAs editMovement attribute
	 *
	 * @return
	 */
	public String getEditAs()
	{
		return editAs;
	}

	/**
	 * set the editAs editMovement attribute
	 *
	 * @return
	 */
	public void setEditAs( String editAs )
	{
		this.editAs = editAs;
	}

	/**
	 * utility to return the shape properties element (picture element only)
	 * should be depreciated when OOXML is completely distinct from BIFF8
	 *
	 * @return
	 */
	public SpPr getSppr()
	{
		if( o != null )
		{
			return o.getSppr();
		}
		return null;
	}

	/**
	 * set this twoCellAnchor as a chart element
	 * used for
	 *
	 * @param rid
	 * @param name
	 * @param bounds
	 */
	public void setAsChart( int rid, String name, int[] bounds )
	{
		o = new ObjectChoice();
		o.setObject( new GraphicFrame() );
		o.setName( name );
		o.setChartRId( "rId" + Integer.valueOf( rid ).toString() );
		o.setId( rid );
		setBounds( bounds );
		// id???
	}

	/**
	 * set this twoCellAnchor as an image
	 *
	 * @param rid
	 * @param name
	 * @param id
	 */
	public void setAsImage( int rid, String name, String descr, int spid, SpPr sp )
	{
		o = new ObjectChoice();
		o.setObject( new Pic() );
		o.setName( name );
		o.setDescr( descr );
		o.setEmbed( "rId" + Integer.valueOf( rid ).toString() );
		o.setId( spid );
		if( sp != null )
		{
			((Pic) o.getObject()).setSppr( sp );
		}
	}

	/**
	 * given bounds[8] in BIFF 8 coordinates and convert to OOXML units
	 * mostly this means adjusting ROWOFFSET 0 & 1 (bounds[5] & bounds[7]
	 * and COLOFFSET 0 & 1 (bounds[1] & bounds[5] to EMU units +
	 * adjust according to emperical calc garnered from observation
	 *
	 * @param sheet
	 * @param bbounds short[] bounds in BIFF8 units
	 * @return new bounds int[] bounds in OOXML units
	 */
	public static int[] convertBoundsFromBIFF8( com.extentech.formats.XLS.Boundsheet sheet, short[] bbounds )
	{
		// note on bounds:
		// -- offsets (bounds 1,3,5 & 7) are %'s of column width or row height, respectively
		// below calculations are garnered from OOXML info + by comparing to Excel's values ... may not be 100% but appears to work for both 2003 and 2007 versions
		// NOTE: if change offset algorithm here must modify algorithm in twoCellAnchor from and to classes
		int[] bounds = new int[8];
		bounds[0] = bbounds[0];
		double cw = ColHandle.getWidth( sheet, bbounds[0] );    ///5.0;
// below is more correct (???)		bounds[1]= (int)(EMU*cw*(bbounds[1]/1024.0));
		bounds[1] = (int) (cw * 256 * (bbounds[1] / 1024.0));
		double rh = RowHandle.getHeight( sheet, bbounds[2] ); ///2.0;
		bounds[2] = bbounds[2];
		bounds[3] = (int) (EMU * rh * (bbounds[3] / 256.0));
		cw = ColHandle.getWidth( sheet, bbounds[4] );    ///5.0;
		bounds[4] = bbounds[4];
//		bounds[5]= (int)(EMU*cw*(bbounds[5]/1024.0));
		bounds[5] = (int) (cw * 256 * (bbounds[5] / 1024.0));
		rh = RowHandle.getHeight( sheet, bbounds[6] );    ///2.0;
		bounds[6] = bbounds[6];
		bounds[7] = (int) (EMU * rh * (bbounds[7] / 256.0));
		return bounds;
	}

	/**
	 * convert bounds[8] from OOXML to BIFF8 Units
	 * basically must adjust COLOFFSETs and ROWOFFSETs to BIFF8 units
	 *
	 * @param sheet
	 * @param bounds int[]
	 * @return bounds short[] (BIFF8 uses short[] + different units)
	 */
	public static short[] convertBoundsToBIFF8( com.extentech.formats.XLS.Boundsheet sheet, int[] bounds )
	{
		short[] bbounds = new short[8];
		bbounds[0] = (short) bounds[0];
		double cw = ColHandle.getWidth( sheet, bounds[0] );    ///5.0;
// below is more correct (???)		bbounds[1]= (short)((bounds[1]*1024)/(TwoCellAnchor.EMU*cw));
		bbounds[1] = (short) ((bounds[1] * 1024.0) / (cw * 256));
		bbounds[2] = (short) bounds[2];
		double rh = RowHandle.getHeight( sheet, bounds[2] );    ///2.0;
		bbounds[3] = (short) ((bounds[3] * 256.0) / (TwoCellAnchor.EMU * rh));
		bbounds[4] = (short) bounds[4];
		cw = ColHandle.getWidth( sheet, bounds[4] );    ///5.0;
//    	bbounds[5]= (short)((bounds[5]*1024)/(TwoCellAnchor.EMU*cw));
		bbounds[5] = (short) ((bounds[5] * 1024.0) / (cw * 256));
		bbounds[6] = (short) bounds[6];
		rh = RowHandle.getHeight( sheet, bounds[6] );    ///2.0;
		bbounds[7] = (short) ((bounds[7] * 256.0) / (TwoCellAnchor.EMU * rh));
		return bbounds;
	}
}

/**
 * from (Starting Anchor Point)
 * This element specifies the first anchor point for the drawing element. This will be used to anchor the top and left
 * sides of the shape within the spreadsheet. That is when the cell that is specified in the from element is adjusted,
 * the shape will also be adjusted.
 * <p/>
 * NOTE: Coordinates are in OOXML units; they are converted to BIFF8 units using twoCellAnchor.convertBoundsToBIFF8
 * parent: oneCellAnchor, twoCellAnchor
 * children: col, colOff, row, rowOff
 */
class From implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( From.class );
	private static final long serialVersionUID = -4776435343244555855L;
	private int[] bounds;

	public From( int[] bounds )
	{
		this.bounds = new int[4];
		System.arraycopy( bounds, 0, this.bounds, 0, 4 );
	}

	public From( From f )
	{
		bounds = f.bounds;
	}

	public static From parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		int[] bounds = new int[4];
		int boundsidx = 0;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "col" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "colOff" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "row" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "rowOff" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "from" ) )
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
			log.error( "from.parseOOXML: " + e.toString() );
		}
		From f = new From( bounds );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:from>" );
		ooxml.append( "<xdr:col>" + bounds[0] + "</xdr:col>" );
		ooxml.append( "<xdr:colOff>" + bounds[1] + "</xdr:colOff>" );
		ooxml.append( "<xdr:row>" + bounds[2] + "</xdr:row>" );
		ooxml.append( "<xdr:rowOff>" + bounds[3] + "</xdr:rowOff>" );
		ooxml.append( "</xdr:from>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new From( this );
	}

	public int[] getBounds()
	{
		if( bounds == null )
		{
			bounds = new int[4];
		}
		return bounds;
	}

	public void setBounds( int[] bounds )
	{
		System.arraycopy( bounds, 0, this.bounds, 0, 4 );
	}
}

/**
 * to (Ending Anchor Point)
 * This element specifies the second anchor point for the drawing element. This will be used to anchor the bottom
 * and right sides of the shape within the spreadsheet. That is when the cell that is specified in the to element is
 * adjusted, the shape will also be adjusted.to (Ending Anchor Point)
 * This element specifies the second anchor point for the drawing element. This will be used to anchor the bottom
 * and right sides of the shape within the spreadsheet. That is when the cell that is specified in the to element is
 * adjusted, the shape will also be adjusted.
 * <p/>
 * NOTE: Coordinates are in OOXML units; they are converted to BIFF8 units using twoCellAnchor.convertBoundsToBIFF8
 * parent: twoCellAnchor
 * children: col, colOff, row, rowOff
 */
class To implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( To.class );
	private static final long serialVersionUID = 1500243445505400113L;
	private int[] bounds;

	public To( int[] bounds )
	{
		this.bounds = new int[4];
		System.arraycopy( bounds, 0, this.bounds, 0, 4 );
	}

	public To( To f )
	{
		bounds = f.bounds;
	}

	public static To parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		int[] bounds = new int[4];
		int boundsidx = 0;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "col" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "colOff" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "row" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
					else if( tnm.equals( "rowOff" ) )
					{
						bounds[boundsidx++] = Integer.parseInt( xpp.nextText() );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "to" ) )
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
			log.error( "to.parseOOXML: " + e.toString() );
		}
		To f = new To( bounds );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:to>" );
		ooxml.append( "<xdr:col>" + bounds[0] + "</xdr:col>" );
		ooxml.append( "<xdr:colOff>" + bounds[1] + "</xdr:colOff>" );
		ooxml.append( "<xdr:row>" + bounds[2] + "</xdr:row>" );
		ooxml.append( "<xdr:rowOff>" + bounds[3] + "</xdr:rowOff>" );
		ooxml.append( "</xdr:to>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new To( this );
	}

	// Access Methods
	public int[] getBounds()
	{
		if( bounds == null )
		{
			bounds = new int[4];
		}
		return bounds;
	}

	public void setBounds( int[] bounds )
	{
		this.bounds = new int[4];
		System.arraycopy( bounds, 0, this.bounds, 0, 4 );
	}
}

