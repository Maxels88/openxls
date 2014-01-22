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

import java.io.Serializable;

/**
 * <b>SELECTION 0x1Dh: Describes the currently selected area of a Sheet.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       pnn         1       Number of the pane described
 * 5       rwAct       2       Row of the active Cell
 * 7       colAct      2       Col of the active Cell
 * 9       irefAct     2       Ref number of the active Cell
 * 11      cref        2       Number of refs in the selection
 * 13      rgref       var     Array of refs
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 * @see ROW
 * @see Cell
 * @see XLSRecord
 */
public class Selection extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2949920585425685061L;
	short pnn = 0;
	short rwAct = 0;
	short colAct = 0;
	short irefAct = 0;
	short cref = 0;
	rgref[] refs;

	public void init()
	{
		super.init();
		rwAct = ByteTools.readShort( this.getByteAt( 1 ), this.getByteAt( 2 ) );
		colAct = ByteTools.readShort( this.getByteAt( 3 ), this.getByteAt( 4 ) );
		irefAct = ByteTools.readShort( this.getByteAt( 5 ), this.getByteAt( 6 ) );
		cref = ByteTools.readShort( this.getByteAt( 7 ), this.getByteAt( 8 ) );

		// cref is count of ref structs -- each one is 6 bytes
		refs = new rgref[cref];
		int ctr = 9;
		for( int i = 0; i < (int) cref; i++ )
		{
			byte[] b1 = new byte[6];
			for( int x = 0; x < 6; x++ )
			{
				b1[x] = this.getByteAt( ctr++ );
			}
			refs[i] = new rgref( b1 );
		}
		int checklen = (cref * 6) + 9;
		// Logger.logInfo("Done adding " + String.valueOf(cref) + " rgrefs to SELECTION.  Record should be: " + String.valueOf(checklen) + " bytes long.");
	}

}

/**
 * Ref Structures -- allows handling of multiple
 * selection information
 */
class rgref implements Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7261215340827889609L;
	byte[] info;

	public rgref( byte[] b )
	{
		info = b;
	}
}