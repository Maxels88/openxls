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
import com.extentech.toolkit.Logger;

/**
 * <b>USERSVIEWBEGIN: Custom View Settings (1AAh)</b><br>
 * <p/>
 * USERSVIEWBEGIN describes the settings for a custom view for the sheet
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       guid            16      GID for custom view
 * 20      iTabid          4       Tab index for the sheet (1-based)
 * 24      wScale          4       Window Zoom
 * 28      icv             4       Index to color val
 * 32      pnnSel          4       Pane number of active pane
 * 36      grbit           4       Option flags
 * 40      refTopLeft      8       Ref struct describing the visible area of top left pane
 * 48      operNum         16      array of 2 floats specifying vert/horiz pane split
 * 64      colRPane        2       first visible right pane col
 * 66      rwBPane         2       first visible bottom pane col
 * <p/>
 * </p></pre>
 */
public final class Usersviewbegin extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1877650235927064991L;
	// record fields
	private int tabid = -1;
	private int wScale = -1;
	private int icv = -1;
	private int pnnSel = -1;
	private int grbit = -1;
	//private float operNum1 = 0;
	//private float operNum2 = 0;
	//private short colRPane = 0;
	//private short rwBPane = 0;

	//66 Long!
	//byte[] RECBYTES =  {0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0, 0, 0, 0, 0,0, 0, 0, 0, 0, 0};

	/**
	 * Constructor for a Usersviewbegin to be made on the fly.
	 */
	Usersviewbegin()
	{
		// TODO: init Usersviewbegin
		this( new byte[]{
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0
		} );
	}

	Usersviewbegin( byte[] b )
	{
		setData( b );
		setOpcode( USERSVIEWBEGIN );
		setLength( (short) 6 );
		this.init();
	}

	// grbit fields
	private boolean fDspGutsSv = false; // true if outline symbols are displayed

	// TODO: implement this class
	@Override
	public void init()
	{
		super.init();
		short num1 = ByteTools.readShort( this.getByteAt( 16 ), this.getByteAt( 17 ) );
		short num2 = ByteTools.readShort( this.getByteAt( 18 ), this.getByteAt( 19 ) );
		tabid = ByteTools.readInt( num2, num1 );
		num1 = ByteTools.readShort( this.getByteAt( 20 ), this.getByteAt( 21 ) );
		num2 = ByteTools.readShort( this.getByteAt( 22 ), this.getByteAt( 23 ) );
		wScale = ByteTools.readInt( num1, num2 );
		num1 = ByteTools.readShort( this.getByteAt( 24 ), this.getByteAt( 25 ) );
		num2 = ByteTools.readShort( this.getByteAt( 26 ), this.getByteAt( 27 ) );
		icv = ByteTools.readInt( num1, num2 );
		num1 = ByteTools.readShort( this.getByteAt( 24 ), this.getByteAt( 25 ) );
		num2 = ByteTools.readShort( this.getByteAt( 26 ), this.getByteAt( 27 ) );
		pnnSel = ByteTools.readInt( num1, num2 );
		num1 = ByteTools.readShort( this.getByteAt( 28 ), this.getByteAt( 29 ) );
		num2 = ByteTools.readShort( this.getByteAt( 30 ), this.getByteAt( 31 ) );
		grbit = ByteTools.readInt( num1, num2 );
		this.decodeGrbit();
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "Usersviewbegin Tab Index: " + tabid );
		}
	}

	/**
	 * decodeGrbit does masking to determine the grbit settings
	 */
	private void decodeGrbit()
	{
		if( (grbit & 0x00000010) == 0x00000010 )
		{
			fDspGutsSv = true;
		}
		else
		{
			fDspGutsSv = false;
		}

	}

	/**
	 * updateGrbit looks at all the grbit variables and rebuilds the grbit
	 * field based off those.
	 */
	private void updateGrbit()
	{
		if( fDspGutsSv )
		{
			grbit = (0x00000010 | grbit);
		}
		else
		{
			grbit = (0xFFFFFFEF & grbit);
		}
		// update the record bytes
		byte[] b = this.getData();
		byte grbytes[] = ByteTools.cLongToLEBytes( grbit );
		System.arraycopy( grbytes, 0, b, 28, 4 );
		this.setData( b );
	}

	/**
	 * Checks to see if outlines are being displayed on the worksheet
	 *
	 * @return fDspGutsSv
	 */
	public boolean getDisplayOutlines()
	{
		return fDspGutsSv;
	}

	/**
	 * Sets whether outlines are displayed on the worksheet.  Should be set to true
	 * when using any of the 'grouping' methods.
	 *
	 * @param disp
	 */
	public void setDisplayOutlines( boolean disp )
	{
		fDspGutsSv = disp;
		updateGrbit();
	}
	/*
	  20      iTabid          4       Tab index for the sheet (1-based)
    */

}