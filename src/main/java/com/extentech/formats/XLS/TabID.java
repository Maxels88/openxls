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
package com.extentech.formats.XLS;

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <b>RRTabID:  Revision Tab ID Record</b><br>
 * <p/>
 * The RRTabId record specifies an array of unique sheet identifiers,
 * each of which is associated with a sheet in the workbook.
 * The order of the sheet identifiers in the array matches the order of
 * the BoundSheet8 records as they appear in the Globals substream.
 */
public final class TabID extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( TabID.class );
	private static final long serialVersionUID = 722748113519841817L;
	CompatibleVector tabIDs = new CompatibleVector();

	/**
	 * Default init
	 */
	@Override
	public void init()
	{
		super.init();
		for( int i = 0; i < (this.getLength() - 4); )
		{
			short s = ByteTools.readShort( this.getByteAt( i ), this.getByteAt( i + 1 ) );
			Short sh = s;
			tabIDs.add( sh );
			i += 2;
		}
	}

	/**
	 * Looks sequentally at the tabIDs and
	 * makes a new one larger than the previous largest...
	 */
	void removeRecord()
	{
		short largest = 0;
		for( Object tabID : tabIDs )
		{
			Short sh = (Short) tabID;
			short newshort = sh;
			if( newshort > largest )
			{
				largest = newshort;
			}
		}
		tabIDs.remove( Short.valueOf( largest ) );
		this.updateRecord();
	}

	/**
	 * Looks sequentally at the tabIDs and
	 * makes a new one larger than the previous largest...
	 */
	void addNewRecord()
	{
		short largest = 0;
		for( Object tabID : tabIDs )
		{
			Short sh = (Short) tabID;
			short newshort = sh;
			if( newshort > largest )
			{
				largest = newshort;
			}
		}
		largest += 0x1;
		Short sh = largest;
		tabIDs.add( sh );
		this.updateRecord();
	}

	/**
	 * This DOES NOT do what was expected.  Sheet order is soley based off of physical Boundsheet ordering
	 * in the output file.  I'm keeping this code in here in case we start supporting revisions.
	 */
	private boolean changeOrder( int sheet, int newpos )
	{
		int sz = tabIDs.size();
		if( ((sheet < 0) || (newpos < 0)) || ((sheet >= sz) || (newpos >= sz)) )
		{
			log.warn( "changing Sheet order failed: invalid Sheet Index: " + sheet + ":" + newpos );
			return false;
		}
		Object b = tabIDs.get( sheet );
		tabIDs.remove( b );
		tabIDs.insertElementAt( b, newpos );
		this.updateRecord();
		return true;
	}

	/**
	 * Updates the underlying byte array with the ordered tabId's
	 * Call after any modification to this record
	 */
	public void updateRecord()
	{
		short newlen = (short) (tabIDs.size() * 2);
		byte[] newbody = new byte[newlen];
		int counter = 0;
		for( Object tabID : tabIDs )
		{
			Short sh = (Short) tabID;
			byte[] b = ByteTools.shortToLEBytes( sh );
			newbody[counter] = b[0];
			newbody[counter + 1] = b[1];
			counter += 2;
		}
		this.setData( newbody );
	}

	/**
	 * @return Returns the tabIDs.
	 */
	public CompatibleVector getTabIDs()
	{
		return tabIDs;
	}
}