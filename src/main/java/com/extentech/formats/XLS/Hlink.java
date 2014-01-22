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

import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.Serializable;
import java.io.UnsupportedEncodingException;

/**
 * <b>Hlink: Hyperlink (1b8h)</b><br>
 * <p/>
 * hyperlink record
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       rwFirst         2       First row of link
 * 6       rwLast          2       Last row of link
 * 8       colFirst        2       First col of link
 * 10      colLast         2       Last col of link
 * 12      rgbHlink        var     HLINK data
 * <p/>
 * <p/>
 * </p></pre>
 */
public final class Hlink extends XLSRecord
{

	private static final long serialVersionUID = -4259979643231173799L;
	int colFirst = -1, colLast = -1, rowFirst = -1, rowLast = -1;
	private HLinkStruct linkStruct = null;

	/**
	 * set last/first cols/rows
	 */
	public void setRowFirst( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 0, 2 );
		this.rowFirst = c;
	}

	public int getRowFirst()
	{
		return rowFirst;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setRowLast( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 2, 2 );
		this.rowLast = c;
	}

	public int getRowLast()
	{
		return rowLast;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setColFirst( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 4, 2 );
		this.colFirst = c;
	}

	public int getColFirst()
	{
		return colFirst;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setColLast( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 6, 2 );
		this.colLast = c;
	}

	public int getColLast()
	{
		return colLast;
	}

	/**
	 * get the URL for this Hlink
	 */
	public String getURL()
	{
		if( linkStruct == null )
		{
			return "";
		}
		return linkStruct.getUrl();
	}

	/**
	 * return the description part of the hyperlink
	 *
	 * @return
	 */
	public String getDescription()
	{
		if( linkStruct == null )
		{
			return "";
		}
		return linkStruct.getLinkText();
	}

	public static XLSRecord getPrototype()
	{
		Hlink retlab = new Hlink();
		byte[] rbytes = {
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				-48,
				-55,
				-22,
				121,
				-7,
				-70,
				-50,
				17,
				-116,
				-126,
				0,
				-86,
				0,
				75,
				-87,
				11,
				2,
				0,
				0,
				0,
				2,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				-32,
				-55,
				-22,
				121,
				-7,
				-70,
				-50,
				17,
				-116,
				-126,
				0,
				-86,
				0,
				75,
				-87,
				11,
				-116,
				0,
				0,
				0,
				0,
				0
		};

		retlab.setOpcode( HLINK );
		retlab.setLength( (short) 94 );
		retlab.setData( rbytes );
		retlab.init();
//        retlab.setURL(url, desc, textMark);		200606 KSC: do separately!
		return retlab;
	}

	/**
	 * set link URL with description and test mark
	 * note that either url or text mark must be present ...
	 *
	 * @param url
	 * @param textMark
	 * @param desc
	 */
	public void setURL( String url, String desc, String textMark )
	{
		try
		{
			if( url.equals( "" ) && textMark.equals( "" ) )
			{
				Logger.logWarn( "HLINK.setURL:  no url or text mark specified" );
				return;
			}
			linkStruct.setUrl( url, desc, textMark );
		}
		catch( Exception e )
		{
			Logger.logWarn( "setting URL " + url + " failed: " + e );
		}
		byte[] bt = linkStruct.getBytes();
		this.setData( bt );
	}

	public void setFileURL( String url, String desc, String textMark )
	{
		try
		{
			linkStruct.setFileURL( url, desc, textMark );
		}
		catch( Exception e )
		{
			Logger.logWarn( "setting URL " + url + " failed: " + e );
		}
		byte[] bt = linkStruct.getBytes();
		this.setData( bt );
	}

	/**
	 * returns whether a given col
	 * is referenced by this Hyperlink
	 */
	public boolean inrange( int x )
	{
		if( (x <= colLast) && (x >= colFirst) )
		{
			return true;
		}
		return false;
	}

	private CellRange range = null;

	@Override
	public void init()
	{
		super.init();
		int pos = 0;
		rowFirst = (int) ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
		rowLast = (int) ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
		colFirst = (int) ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
		colLast = (int) ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
		Unicodestring ustr = new Unicodestring();

		String nm = "";
		if( this.getSheet() != null )
		{
			if( !this.getSheet().getSheetName().equals( "" ) )
			{
				nm = this.getSheet().getSheetName() + "!";
			}
		}
		nm += com.extentech.ExtenXLS.ExcelTools.getAlphaVal( colFirst );
		nm += rowFirst + ":";
		nm += com.extentech.ExtenXLS.ExcelTools.getAlphaVal( colLast );
		nm += rowLast;
		try
		{
			range = new CellRange( nm, null );
		}
		catch( Exception e )
		{
			Logger.logWarn( "initializing Hlink record failed: " + e );
		}
		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "Hlink Cells: " + range.toString() );
		}

		try
		{
			linkStruct = new HLinkStruct( this.getBytesAt( 0, this.getLength() ) );
		}
		catch( Exception e )
		{
			Logger.logWarn( "Hyperlink parse failed for Cells " + range.toString() + ": " + e );
		}
		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "Hlink URL: " + linkStruct.getUrl() );
		}
	}

	/**
	 * @return
	 */
	public CellRange getRange()
	{
		return range;
	}

	/**
	 * @param range
	 */
	public void setRange( CellRange range )
	{
		this.range = range;
	}

	public void initCells( WorkBookHandle wbook )
	{
		int[] cellcoords = new int[4];
		cellcoords[0] = this.getRowFirst();
		cellcoords[2] = this.getRowLast();
		cellcoords[1] = this.getColFirst();
		cellcoords[3] = this.getColLast();

		try
		{
			CellRange cr = new CellRange( wbook.getWorkSheet( this.getSheet().getSheetName() ), cellcoords );
			cr.setWorkBook( wbook );
			BiffRec[] ch = cr.getCellRecs();
			for( int t = 0; t < ch.length; t++ )
			{
				ch[t].setHyperlink( this );
			}
		}
		catch( Exception e )
		{
			Logger.logWarn( "initializing Hyperlink Cells failed: " + e );
		}
	}
}

/*
    HLINK Struct
    ??      16
    int?    4
    int?    4
    cch?    4
    var     var
    int?    4
    ?       16
    cch     4
    urlstr  var
*/
class HLinkStruct implements XLSConstants, Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1915454683496117350L;
	boolean isLink = false;
	boolean isAbsoluteLink = false;
	boolean hasDescription = false;
	boolean hasTextMark = false;
	boolean hasTargetFrame = false;
	boolean isUNCPath = false;
	byte grbit[] = new byte[4];

	/**
	 * decodes option flag into vars
	 * Here is the breakdown on what these mean:
	 * <p/>
	 * standard (non-local) URL:	isLink= true, isAbsoluteLink= true, isUNCPath= false
	 * local file:					isLink= true, isUNCPath= false
	 * UNC path:					isLink= true, isAbsoluteLink= true, isUNCPath= true
	 * link in current workbook:	isLink= false, isAbsoluteLink= false, hasTextMark= true, isUNCPath= false
	 *
	 * @param grbytes
	 */
	void decodeGrbit( byte[] grbytes )
	{
		if( (byte) (grbytes[0] & 0x1) > 0x0 )
		{
			isLink = true;
		}
		if( (byte) (grbytes[0] & 0x2) > 0x0 )
		{
			isAbsoluteLink = true;
		}
		// 20060406 KSC: need both bits 2 & 4 to be set for hasDescription
//		if ((byte)(grbytes[0] & 0x14) > 0x0)hasDescription = true;
		if( (byte) (grbytes[0] & 0x14) >= 0x14 )
		{
			hasDescription = true;
		}
		if( (byte) (grbytes[0] & 0x8) > 0x0 )
		{
			hasTextMark = true;
		}
		if( (byte) (grbytes[0] & 0x80) > 0x0 )
		{
			hasTargetFrame = true;
		}
		if( (byte) (grbytes[0] & 0x100) > 0x0 )
		{
			isUNCPath = true;
		}
	}

	void setGrbit()
	{
		grbit[0] = 0x0;
		if( isLink )
		{
			grbit[0] = (byte) (0x1 | grbit[0]);
		}
		if( isAbsoluteLink )
		{
			grbit[0] = (byte) (0x2 | grbit[0]);
		}
		if( hasDescription )
		{
			grbit[0] = (byte) (0x14 | grbit[0]);
		}
		if( hasTextMark )
		{
			grbit[0] = (byte) (0x8 | grbit[0]);
		}
		if( hasTargetFrame )
		{
			grbit[0] = (byte) (0x80 | grbit[0]);
		}
		if( isUNCPath )
		{
			grbit[0] = (byte) (0x100 | grbit[0]);
		}
		mybytes[28] = grbit[0];
	}

	boolean DEBUG = false;

	int getUrlPos()
	{
		return urlcch;
	}

	byte[] getBytes()
	{
		return mybytes;
	}

	String url = "", linktext = "", textMark = "", targetFrame = "";
	int int1 = -1, urlcch = -1, int4 = -1;
	private byte[] mybytes = null;

	/**
	 * get the URL for this Hlink
	 */
	String getUrl()
	{
		if( textMark.equals( "" ) )
		{
			return url;
		}
		return url + "#" + textMark;
	}

	/**
	 * get the URL link text for this Hlink
	 */
	public String getLinkText()
	{
		return linktext;
	}

	// 20060406 KSC:  mods to setUrl:
	// 	Added ability to set description + modified byte input to work mo' betta ... 

	/**
	 * set the URL for this Hlink, description= URL <default>
	 * <p/>
	 * Assume link URL i.e. no file URL, UNC path, etc.
	 */
	void setUrl( String url )
	{
		setUrl( url, url, "" );
	}

	/**
	 * set standard link URL for this Hlink and optional description
	 * i.e. no file URL, UNC path, text marks ...
	 */
	void setUrl( String ur, String desc )
	{
		setUrl( url, desc, "" );
	}

	/**
	 * set proper settings for URL and write bytes
	 *
	 * @param ur
	 * @param desc
	 * @param textMark
	 */
	void setUrl( String ur, String desc, String textMark )
	{
		isLink = true;
		isAbsoluteLink = true;
		isUNCPath = false;
		hasDescription = desc.length() > 0;
		hasTextMark = textMark.length() > 0;
		setBytes( ur, desc, textMark );
	}

	/**
	 * Assume NOT TRUE FILE URL (avoids writing complex dir info + using FILE_GUID)
	 * so difference between link URL and file URL is the isAbsoluteLink var ...
	 *
	 * @param ur
	 * @param desc
	 * @param textMark
	 */
	void setFileURL( String ur, String desc, String textMark )
	{
		isLink = true;
		isAbsoluteLink = false;
		isUNCPath = false;
		hasDescription = desc.length() > 0;
		hasTextMark = textMark.length() > 0;
		setBytes( ur, desc, textMark );
	}

	/**
	 * fills the HLINK bytes based on the settings of
	 * isLink
	 * isAbsoluteLink
	 * hasDescription
	 * hasTextMark
	 * hasTargetFrame
	 * isUNCPath
	 * <p/>
	 * how these are set, before calling this method, determine
	 * how the HLINK record bytes are written
	 *
	 * @param ur    URL string
	 * @param desc    optional descrption
	 * @param tm    optional text mark text as in:  ...#textmarktext
	 */
	static final byte[] URL_GUID = { -32, -55, -22, 121, -7, -70, -50, 17, -116, -126, 0, -86, 0, 75, -87, 11 };
	static final byte[] FILE_GUID = { 3, 3, 0, 0, 0, 0, 0, 0, -64, 0, 0, 0, 0, 0, 0, 70 };

	void setBytes( String ur, String desc, String tm )
	{
		try
		{
			setGrbit();
			int pos = 32;    // start of description/text input

			byte[] blankbytes = new byte[2];    // trailing zero word
			byte[] newbytes = new byte[pos];
			System.arraycopy( mybytes, 0, newbytes, 0, pos ); // copy old pre-string data into new array

			if( hasDescription )
			{
				// copy char array length (cch) of description + description bytes
				byte[] descbytes = desc.getBytes( UNICODEENCODING );
				int newcch = descbytes.length;
				byte[] newcchbytes = ByteTools.cLongToLEBytes( (newcch / 2) + 1 );

				// copy cch of desc in...
				newbytes = ByteTools.append( newcchbytes, newbytes );

				//copy bytes of description in
				newbytes = ByteTools.append( descbytes, newbytes );
				// copy trailing dumb str bytes in
				newbytes = ByteTools.append( blankbytes, newbytes );
			}			
	
			/* TODO:  Implement target frame right here
			if (hasTargetFrame) {
				// copy targetFrame bytes + length
				// get cch
				byte[] tfbytes= tf.getBytes(UNICODEENCODING);
	            int newcch = tfbytes.length;
				byte[] newcchbytes = ByteTools.cLongToLEBytes(newcch/2 +1);
				// copy cch of tm in...
				newbytes = ByteTools.append(newcchbytes, newbytes);
				
				//copy bytes of tf in
				newbytes = ByteTools.append(tfbytes, newbytes);

				// copy trailing dumb str bytes in
				newbytes = ByteTools.append(blankbytes, newbytes);		        
			}
			*/

			/* URL Handling */
			// copy GUID in - ASSUME URL_GUID since aren't supporting relative FILE_GUIDs!
			newbytes = ByteTools.append( URL_GUID, newbytes );

			if( isLink )
			{
				// copy url bytes + length
				// get cch (which is different alg. from both description + textmark)
				byte[] urlbytes = ur.getBytes( UNICODEENCODING );
				int newcch = urlbytes.length;

				// copy cch of url in...
				byte[] newcchbytes = ByteTools.cLongToLEBytes( newcch + 2 );
				newbytes = ByteTools.append( newcchbytes, newbytes );

				// copy url bytes in
				newbytes = ByteTools.append( urlbytes, newbytes );

				// copy trailing dumb str bytes in
				newbytes = ByteTools.append( blankbytes, newbytes );
			}

			if( hasTextMark )
			{
				// copy textmark bytes + length
				// get cch
				byte[] tmbytes = tm.getBytes( UNICODEENCODING );
				int newcch = tmbytes.length;
				byte[] newcchbytes = ByteTools.cLongToLEBytes( (newcch / 2) + 1 );
				// copy cch of tm in...
				newbytes = ByteTools.append( newcchbytes, newbytes );

				//copy bytes of tm in
				newbytes = ByteTools.append( tmbytes, newbytes );

				// copy trailing dumb str bytes in
				newbytes = ByteTools.append( blankbytes, newbytes );
			}
			this.mybytes = newbytes;

			this.linktext = desc;
			this.textMark = tm;
			this.url = ur;
		}
		catch( UnsupportedEncodingException e )
		{
			Logger.logWarn( "Setting URL failed: " + ur + ": " + e );
		}

	}

	/**
	 * Inner class with no documentation
	 */
	HLinkStruct( byte[] barr )
	{
		mybytes = barr;
		int pos = 28;

		System.arraycopy( barr, 28, grbit, 0, 4 );
		decodeGrbit( grbit );
		pos += 4;

		/*
		 * This section gets the display string for the Hyperlink, if it exists.
		 */
		if( hasDescription )
		{
			int cch = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
			if( cch > 0 )
			{ // 20070814 KSC: shouldn't be 0 and also hasDescription ...
				try
				{
					byte[] descripbytes = new byte[(cch * 2) - 2];
					System.arraycopy( barr, pos, descripbytes, 0, (cch * 2) - 2 );
					linktext = new String( descripbytes, UNICODEENCODING );
					pos += cch * 2;
					if( DEBUG )
					{
						Logger.logInfo( "Hlink.hlstruct Display URL: " + linktext );
					}
				}
				catch( Exception e )
				{
					if( DEBUG )
					{
						Logger.logWarn( "decoding Display URL in Hlink: " + e );
					}
				}
			}
		}
		/*
		 * if it has a target frame, read in
		 */
		if( hasTargetFrame )
		{
			int cch = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
			if( cch > 0 )
			{
				try
				{
					byte[] tfbytes = new byte[(cch * 2) - 2];
					System.arraycopy( barr, pos, tfbytes, 0, (cch * 2) - 2 );
					targetFrame = new String( tfbytes, UNICODEENCODING );
					if( DEBUG )
					{
						Logger.logInfo( "Hlink.hlstruct targetFrame: " + targetFrame );
					}
					pos += (cch * 2);
				}
				catch( Exception e )
				{
					if( DEBUG )
					{
						Logger.logWarn( "Hlink Decode of targetFrame failed: " + e );
					}
				}
			}
		}
		/*
		 * URL section:  non-local URL or Link in current file   
		 */
		if( isLink )
		{
			byte[] GUID = new byte[16];
			System.arraycopy( barr, pos, GUID, 0, 16 );
			boolean bIsCurrentFileRef = java.util.Arrays.equals( GUID, FILE_GUID );
			pos += 16;    // skip GUID
			if( !bIsCurrentFileRef )
			{    // then it's a URL or non-relative file path
				urlcch = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
				if( urlcch > 0 )
				{
					try
					{
						byte[] urlbytes = new byte[(urlcch) - 2];
						System.arraycopy( barr, pos, urlbytes, 0, (urlcch) - 2 );
						url = new String( urlbytes, UNICODEENCODING );
						if( DEBUG )
						{
							Logger.logInfo( "Hlink.hlstruct URL: " + url );
						}
						pos += urlcch;
					}
					catch( Exception e )
					{
						if( DEBUG )
						{
							Logger.logWarn( "Hlink Decode of URL failed: " + e );
						}
					}
				}
			}
			else
			{    // (appears to be a) current file link (Actuality is different than documentation!)
				int dirUps = ByteTools.readShort( barr[pos++], barr[pos++] );
				urlcch = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
				if( urlcch > 0 )
				{
					try
					{
						byte[] urlbytes = new byte[urlcch - 1];
						System.arraycopy( barr, pos, urlbytes, 0, urlcch - 1 );
						url = new String( urlbytes, DEFAULTENCODING );
						if( DEBUG )
						{
							Logger.logInfo( "Hlink.hlstruct File URL: " + url );
						}
						pos += urlcch + 24;    // add char count + avoid the 24 "unknown" bytes
						int extraInfo = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
						if( extraInfo > 0 )
						{
							pos += extraInfo;
							int sz = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );

						}
					}
					catch( Exception e )
					{
						if( DEBUG )
						{
							Logger.logWarn( "Hlink Decode of File URL failed: " + e );
						}
					}
				}
			}
		}

		if( hasTextMark )
		{
			int cch = ByteTools.readInt( barr[pos++], barr[pos++], barr[pos++], barr[pos++] );
			if( cch > 0 )
			{
				try
				{
					byte[] tmbytes = new byte[(cch * 2) - 2];
					System.arraycopy( barr, pos, tmbytes, 0, (cch * 2) - 2 );
					textMark = new String( tmbytes, UNICODEENCODING );
					if( DEBUG )
					{
						Logger.logInfo( "Hlink.hlstruct textMark: " + textMark );
					}
					pos += (cch * 2);
				}
				catch( Exception e )
				{
					if( DEBUG )
					{
						Logger.logWarn( "Hlink Decode of textmark failed: " + e );
					}
				}
			}
		}
	}

}

