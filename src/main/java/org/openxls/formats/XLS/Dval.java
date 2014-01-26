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
package org.openxls.formats.XLS;

import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.List;

/**
 * <b>Dval: Data Validity Settings (01B2h)</b><br>
 * <p/>
 * This record is the list header of the Data Validity Table in the current sheet.
 * <p/>
 * Offset          Name            Size                 Contents
 * -------------------------------------------------------
 * 0                 wDviFlags     2                       Option flags:
 * 2                 xLeft             4                       Horizontal position of the prompt box, if it has fixed position, in pixel
 * 6                 yTop              4                       Vertical position of the prompt box, if it has fixed position, in pixel
 * 10              inObj               4                       Object identifier of the drop down arrow object for a list box , if a list box is visible at
 * the current cursor position, FFFFFFFFH otherwise
 * 14                idvMac            4                      Number of following DV records
 * <p/>
 * <p/>
 * wDviFlags
 * Bit         Mask              Name            Contents
 * ------------------------------------------------------
 * 0           0001H          fWnClosed            0 = Prompt box not visible 1 = Prompt box currently visible
 * 1           0002H          fWnPinned            0 = Prompt box has fixed position 1 = Prompt box appears at cell
 * 2           0004H          fCached              1 = Cell validity data cached in following DV records
 */

public class Dval extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Dval.class );
	private static final long serialVersionUID = 3954586766300169606L;
	// primary fields
	private short grbit;
	private int xLeft;
	private int yTop;
	private int inObj;
	private int idvMac;

	// grbitlookups
	private static final short BITMASK_F_WN_CLOSED = 0x0001;
	private static final short BITMASK_F_WN_PINNED = 0x0002;
	private static final short BITMASK_F_CACHED = 0x0004;

	// 200906060 KSC: Mod bytes to have no x, y or dv recs following:    private static final byte[] PROTOTYPE_BYTES = {04, 00, 124, 00, 00, 00, 00, 00, 00, 00, -1, -1, -1, -1, 01, 00, 00, 00};
//    private static final byte[] PROTOTYPE_BYTES = {04, 00, 00, 00, 00, 00, 00, 00, 00, 00, -1, -1, -1, -1, 00, 00, 00, 00};
	private byte[] PROTOTYPE_BYTES = { 04, 00, 00, 00, 00, 00, 00, 00, 00, 00, -1, -1, -1, -1, 00, 00, 00, 00 };

	private ArrayList dvRecs;

	/**
	 * Standard init method, nothing new
	 *
	 * @see XLSRecord#init()
	 */
	@Override
	public void init()
	{
		super.init();
		int offset = 0;
		dvRecs = new ArrayList();
		grbit = ByteTools.readShort( getByteAt( offset++ ), getByteAt( offset++ ) );
		xLeft = ByteTools.readInt( getByteAt( offset++ ), getByteAt( offset++ ), getByteAt( offset++ ), getByteAt(
				offset++ ) );
		yTop = ByteTools.readInt( getByteAt( offset++ ), getByteAt( offset++ ), getByteAt( offset++ ), getByteAt( offset++ ) );
		inObj = ByteTools.readInt( getByteAt( offset++ ), getByteAt( offset++ ), getByteAt( offset++ ), getByteAt(
				offset++ ) );
		idvMac = ByteTools.readInt( getByteAt( offset++ ), getByteAt( offset++ ), getByteAt( offset++ ), getByteAt(
				offset++ ) );
	}

	/**
	 * Create a dval record & populate with prototype bytes
	 *
	 * @return
	 */
	protected static XLSRecord getPrototype()
	{
		Dval dval = new Dval();
		dval.setOpcode( DVAL );
		dval.setData( dval.PROTOTYPE_BYTES );
		dval.init();
		return dval;
	}

	/**
	 * Apply the grbit to the record in the streamer
	 */
	public void setGrbit()
	{
		byte[] data = getData();
		byte[] b = ByteTools.shortToLEBytes( grbit );
		System.arraycopy( b, 0, data, 0, 2 );
		setData( data );
	}

	/**
	 * Is Cell validity data cached in following DV records?
	 *
	 * @return
	 */
	public boolean isValidityCached()
	{
		return ((grbit & BITMASK_F_CACHED) == BITMASK_F_CACHED);
	}

	/**
	 * Set cell validity data cached in following DV records
	 */
	public void setValidityCached( boolean cached )
	{
		if( cached )
		{
			grbit = (short) (grbit | BITMASK_F_CACHED);
		}
		else
		{
			grbit = (short) (grbit ^ BITMASK_F_CACHED);
		}
		setGrbit();
	}

	/**
	 * Get where the prompt box is located.
	 * true = Prompt box is at cell; false= Prompt box in fixed position
	 *
	 * @return
	 */
	public boolean isPromptBoxAtCell()
	{
		return ((grbit & BITMASK_F_WN_PINNED) == BITMASK_F_WN_PINNED);
	}

	/**
	 * Set where the prompt box is located.
	 * true = Prompt box is at cell; false= Prompt box in fixed position
	 *
	 * @return
	 */
	public void setPromptBoxAtCell( boolean location )
	{
		if( location )
		{
			grbit = (short) (grbit | BITMASK_F_WN_PINNED);
		}
		else
		{
			grbit = (short) (grbit ^ BITMASK_F_WN_PINNED);
		}
		setGrbit();
	}

	/**
	 * Get visibility of prompt box
	 * true = Prompt box currently visible; false = Prompt box not visible
	 *
	 * @return
	 */
	public boolean isPromptBoxVisible()
	{
		return ((grbit & BITMASK_F_WN_CLOSED) == BITMASK_F_WN_CLOSED);
	}

	/**
	 * Set visibility of prompt box
	 * true = Prompt box currently visible; false = Prompt box not visible
	 *
	 * @return
	 */
	public void setPromptBoxVisible( boolean location )
	{
		if( location )
		{
			grbit = (short) (grbit | BITMASK_F_WN_CLOSED);
		}
		else
		{
			grbit = (short) (grbit ^ BITMASK_F_WN_CLOSED);
		}
		setGrbit();
	}

	/**
	 * Get the count of dv records following this Dval
	 *
	 * @return
	 */
	public int getFollowingDvCount()
	{
		return idvMac;
	}

	/**
	 * Set the count of following Dv records
	 *
	 * @param cnt = count
	 */
	public void setFollowingDvCount( int cnt )
	{
		idvMac = cnt;
		byte[] data = getData();
		byte[] b = ByteTools.cLongToLEBytes( idvMac );
		System.arraycopy( b, 0, data, 14, 4 );
		setData( data );
	}

	/**
	 * Object identifier of the drop down arrow object for a list box ,
	 * if a list box is visible at the current cursor position, FFFFFFFFH otherwise
	 *
	 * @return
	 */
	public int getObjectIdentifier()
	{
		return inObj;
	}

	/**
	 * Object identifier of the drop down arrow object for a list box ,
	 * if a list box is visible at the current cursor position, FFFFFFFFH otherwise
	 *
	 * @param cnt = identifier
	 */
	public void setObjectIdentifier( int cnt )
	{
		inObj = cnt;
		byte[] data = getData();
		byte[] b = ByteTools.cLongToLEBytes( inObj );
		System.arraycopy( b, 0, data, 10, 4 );
		setData( data );
	}

	/**
	 * Horizontal position of the prompt box, if it has fixed position, in pixel
	 *
	 * @return
	 */
	public int getHorizontalPosition()
	{
		return xLeft;
	}

	/**
	 * Horizontal position of the prompt box, if it has fixed position, in pixel
	 *
	 * @param cnt = position
	 */
	public void setHorizontalPosition( int cnt )
	{
		xLeft = cnt;
		byte[] data = getData();
		byte[] b = ByteTools.cLongToLEBytes( xLeft );
		System.arraycopy( b, 0, data, 2, 4 );
		setData( data );
	}

	/**
	 * Vertical position of the prompt box, if it has fixed position, in pixel
	 *
	 * @return
	 */
	public int getVerticalPosition()
	{
		return yTop;
	}

	/**
	 * Vertical position of the prompt box, if it has fixed position, in pixel
	 *
	 * @param cnt = position
	 */
	public void setVerticalPosition( int cnt )
	{
		yTop = cnt;
		byte[] data = getData();
		byte[] b = ByteTools.cLongToLEBytes( yTop );
		System.arraycopy( b, 0, data, 2, 4 );
		setData( data );
	}

	/**
	 * Add a new Dv record to this Dval;
	 *
	 * @param dv
	 */
	public void addDvRec( Dv dv )
	{
		dvRecs.add( dv );
	}

	/**
	 * Add a new (ie not on parse) dv record,
	 * updates the parent record with the count
	 *
	 * @param location = the cell/range that the dv attaches to.  Sheet name not required
	 *                 as Dval is a sheet not book level record
	 */
	public Dv createDvRec( String location )
	{
		Dv d = (Dv) Dv.getPrototype( getWorkBook() );
		d.setSheet( getSheet() );
		d.setRange( location );
		addDvRec( d );
		setFollowingDvCount( dvRecs.size() );
		return d;
	}

	/**
	 * Return all dvs for this Dval
	 *
	 * @return
	 */
	public List getDvs()
	{
		return dvRecs;
	}

	/**
	 * Remove a dv rec from this dval.
	 *
	 * @param dv
	 */
	public void removeDvRec( Dv dv )
	{
		dvRecs.remove( dv );
	}

	/**
	 * @param cellAddress
	 * @return
	 */
	public Dv getDv( String cellAddress )
	{
		if( cellAddress.indexOf( "!" ) != -1 )
		{
			cellAddress = cellAddress.substring( cellAddress.indexOf( "!" ) );
		}
		for( Object dvRec : dvRecs )
		{
			Dv d = (Dv) dvRec;
			if( d.isInRange( cellAddress ) )
			{
				return d;
			}
		}
		return null;
	}

/**
 * OOXML Element:
 * dataValidations (Data Validations)
 * This collection expresses all data validation information for cells in a sheet which have data validation features
 * applied.
 * Data validation is used to specify constaints on the type of data that can be entered into a cell. Additional UI can
 * be provided to help the user select valid values (e.g., a dropdown control on the cell or hover text when the cell
 * is active), and to help the user understand why a particular entry was considered invalid (e.g., alerts and
 * messages).
 * Various data types can be selected, and logical operators (e.g., greater than, less than, equal to, etc) can be
 * used. Additionally, instead of specifying an explicit set of values that are valid, a cell or range reference may be
 * used.
 * An input message can be specified to help the user know what kind of value is expected, and a warning message
 * (and warning type) can be specified to alert the user when they've entered invalid data.
 *
 * parent:  worksheet
 * children:  datraValidation
 * attributes:  count (uint), xWindow (uint)-per-sheet x, yWindow (uint)- per-sheet y, disablePrompts (bool)
 *
 * TODO:  shouldnt this be in the OOXML package? why generate/treat this differently?
 */

	/**
	 * create one or more Data Validation records based on OOXML input
	 */
	public static Dval parseOOXML( XmlPullParser xpp, Boundsheet bs )
	{
		Dval dval = bs.insertDvalRec();    // creates or retrieves Dval rec
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "dataValidations" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( n.equals( "count" ) )
							{
							}
							else if( n.equals( "disablePrompts" ) )
							{
								dval.setPromptBoxVisible( false );
							}
							else if( n.equals( "xWindow" ) )
							{
								dval.setHorizontalPosition( Integer.valueOf( v ) );
							}
							else if( n.equals( "yWindow" ) )
							{
								dval.setVerticalPosition( Integer.valueOf( v ) );
							}
						}
					}
					else if( tnm.equals( "dataValidation" ) )
					{    // one or more
						Dv.parseOOXML( xpp, bs );    // creates and adds a new Dv record to the Dval recs list
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "dataValidations" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "OOXMLELEMENT.parseOOXML: " + e.toString() );
		}
		return dval;
	}

	/**
	 * generate the proper OOXML to define this Dval
	 *
	 * @return
	 */
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( dvRecs.size() > 0 )
		{
			ooxml.append( "<dataValidations count=\"" + dvRecs.size() + "\"" );
			if( !isPromptBoxVisible() )
			{
				ooxml.append( " disablePrompts=\"1\"" );
			}
			if( getHorizontalPosition() != 0 )
			{
				ooxml.append( " xWindow=\"" + getHorizontalPosition() + "\"" );
			}
			if( getVerticalPosition() != 0 )
			{
				ooxml.append( " yWindow=\"" + getVerticalPosition() + "\"" );
			}
			ooxml.append( ">" );
			for( Object dvRec : dvRecs )
			{
				ooxml.append( ((Dv) dvRec).getOOXML() );
			}
			ooxml.append( "</dataValidations>" );
		}
		return ooxml.toString();
	}
}
