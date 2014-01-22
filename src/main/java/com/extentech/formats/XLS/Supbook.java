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
import com.extentech.toolkit.Logger;

import java.util.Arrays;

/**
 * <b>Supbook: Supporting Workbook (1AEh)</b><br>
 * <p/>
 * Supbook records store information about a supporting
 * external workbook
 * <p/>
 * <p><pre>
 * for ADD-IN SUPBOOK record (occurs before Global SUPBOOK record)
 * offset  name        size    contents
 * ---
 * 0					2		01 00 (0001H)
 * 2					2		01 58 (0x3A) = Add-In supbook record
 * <p/>
 * for global SUPBOOK record (SUPBOOK 0) = Internal 3d References
 * offset  name        size    contents
 * ---
 * 0		nSheets		2		number of sheets in the workbook
 * 2					2		01 04 = global or own SUPBOOK record
 * <p/>
 * <p/>
 * SUPBOOK records for External References:
 * offset  name        size    contents
 * ---
 * 4       ctab        2       number of tabs in the workbook
 * 6       StVirtPath  var     Encoded file name (unicode)
 * var     Rgst        var     An array of tab sheet names (unicode)
 * <p/>
 * SUPBOOK records for OLE Objects/DDE
 * offset  name        size    contents
 * ---
 * 0					2		0000
 * 2					var		Encoded document name
 * <p/>
 * <p/>
 * </p></pre>
 *
 * @see Boundsheet
 * @see Externsheet
 */

/*
      * Encoded File URLS:
      * 1st char determines type of encoding:
      * 0x1 = Encoded URL follows
      * 0x2 = Reference to a sheet in the own document (sheetname follows)
      * 01H An MS-DOS drive letter will follow, or @ and the server name of a UNC path
  02H Start path name on same drive as own document
  03H End of subdirectory name
  04H Start path name in parent directory of own document (may occur repeatedly)
  05H Unencoded URL. Followed by the length of the URL (1 byte), and the URL itself.
  06H Start path name in installation directory of Excel
  08H Macro template directory in installation directory of Excel
  examples:
  =[ext.xls]Sheet1!A1           <01H>[ext.xls]Sheet1
  ='sub\[ext.xls]Sheet1'!A1         <01H>sub<03H>[ext.xls]Sheet1
  ='\[ext.xls]Sheet1'!A1            <01H><02H>[ext.xls]Sheet1
  ='\sub\[ext.xls]Sheet1'!A1        <01H><02H>sub<03H>[ext.xls]Sheet1
  ='\sub\sub2\[ext.xls]Sheet1'!A1 <01H><02H>sub<03H>sub2<03H>[ext.xls]Sheet1
  ='D:\sub\[ext.xls]Sheet1'!A1  <01H><01H>Dsub<03H>[ext.xls]Sheet1
  ='..\sub\[ext.xls]Sheet1'!A1  <01H><04H>sub<03H>[ext.xls]Sheet1
  ='\\pc\sub\[ext.xls]Sheet1'!A1    <01H><01H>@pc<03H>sub<03H>[ext.xls]Sheet1
  ='http://www.example.org/[ext.xls]Sheet1'!A1   <01H><05H><26H>http://www.example.org/[ext.xls]Sheet1
  (the length of the URL (38 = 26H) follows the 05H byte)
      */

public final class Supbook extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5010774855281594364L;
	int cstab = -1;
	int nSheets = -1;
	CompatibleVector tabs = new CompatibleVector();
	private String filename;    // for EXTERNAL references

	@Override
	public void init()
	{
		super.init();
		this.getData();
		// KSC: Interpret SUPBOOK code
		if( isGlobalRecord() )
		{
			nSheets = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		}
		else if( !isAddInRecord() )
		{    // then it's an External Reference
			// 20080122 KSC: Ressurrect + get code to work
			try
			{
				cstab = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );    // number of tabs/sheetnames referenced
				int ln = ByteTools.readShort( this.getByteAt( 2 ),
				                              this.getByteAt( 3 ) );    // len of bytes of filename "encoded URL without sheetname"
				int compression = this.getByteAt( 4 );    // TODO: ??? 0= compressed 8 bit 1= uncompressed 16 bit ... asian, rtf ...
				int encoding = this.getByteAt( 5 );    // MUST = 0x1, if not it's an internal ref
				int pos = 6;

				if( compression == 0 )
				{ //non-unicode
		            /* this is whacked.  invalid code in several use cases, its trying to get the length of the string
                     * from bytes that are internal in the string itself;
		        	while (this.getByteAt(pos)<0x8) {	// get control chars ...
		        		if (this.getByteAt(pos)==0x5) { // unencoded, file length follows
		        			ln= this.getByteAt(++pos)+2;
		        		}
		        		pos++;
		        		ln--;
		        	}**/
					byte[] f = this.getBytesAt( pos, ln - 1 );
					filename = new String( f );
					pos += ln - 1;
				}
				else
				{    // unicode
					byte[] f = this.getBytesAt( pos, ln * 2 - 1 );
					filename = new String( f, "UTF-16LE" );
					pos += ln * 2 - 1;
				}
				if( DEBUGLEVEL > 5 )
				{
					Logger.logInfo( "Supbook File: " + filename );
				}
				// now get sheetnames
				for( int i = 0; i < cstab; i++ )
				{
					ln = ByteTools.readShort( this.getByteAt( pos ), this.getByteAt( pos + 1 ) );
					compression = this.getByteAt( pos + 2 );    // TODO: ??? 0= compressed 8 bit 1= uncompressed 16 bit ... asian, rtf ...
					pos += 3;
					String sheetname = "";
					if( compression == 0 )
					{
						byte[] f = this.getBytesAt( pos, ln );
						sheetname = new String( f );
						pos += ln;
					}
					else
					{
						byte[] f = this.getBytesAt( pos, ln * 2 );
						sheetname = new String( f, "UTF-16LE" );
						pos += ln * 2;
					}
					tabs.add( sheetname );
					if( DEBUGLEVEL > 5 )
					{
						Logger.logInfo( "Supbook Sheet Reference: " + sheetname );
					}
				}
			}
			catch( Exception e )
			{
				Logger.logErr( "Supbook.init External Record: " + e.toString() );
			}

		}
	}

	static byte[] protocXTI = { 0x0, 0x0, 0x1, 0x4 };
	// KSC: different byte sequence for Add-ins:
	static byte[] AddInProto = { 0x1, 0x0, 0x1, 0x3A };

	static int defaultsize = 4;

	/**
	 * Constructor
	 */
	protected static XLSRecord getPrototype( int numtabs )
	{
		Supbook x = new Supbook();
		x.setLength( (short) defaultsize );
		x.setOpcode( SUPBOOK );
		// x.setLabel("SUPBOOK");
		byte[] dta = new byte[defaultsize];
		System.arraycopy( protocXTI, 0, dta, 0, protocXTI.length );
		dta[0] = (byte) numtabs;
		x.setData( dta );
		x.originalsize = defaultsize;
		x.init();
		return x;
	}

	/**
	 * SupBook record for add-in is different
	 */
	protected static XLSRecord getAddInPrototype()
	{
		Supbook x = new Supbook();
		x.setLength( (short) defaultsize );
		x.setOpcode( SUPBOOK );
		byte[] dta = new byte[AddInProto.length];
		System.arraycopy( AddInProto, 0, dta, 0, AddInProto.length );
		x.setData( dta );
		x.originalsize = defaultsize;
		x.init();
		return x;
	}

	/**
	 * creates a new External Supbook Record for the externalWorkbook
	 *
	 * @param externalWorkbook String
	 * @return
	 */
	protected static XLSRecord getExternalPrototype( String externalWorkbook )
	{
		Supbook x = new Supbook();
		x.setLength( (short) defaultsize );
		x.setOpcode( SUPBOOK );
		x.originalsize = defaultsize;
		// length of externalWorkBook + 4 (cstab + len) + number of encoding options (2 +)
		int compression = 0;    //=non-unicode??
		int encoding = 1;    //=Encoded URL follows
		int n = 2;    //4;	// number of encoding chars
		byte[] f = externalWorkbook.getBytes();
		int ln = f.length + 1;
		byte[] dta = new byte[4 + n + ln - 1];
		System.arraycopy( ByteTools.shortToLEBytes( (short) 0 ), 0, dta, 0, 2 );   // cstab
		System.arraycopy( ByteTools.shortToLEBytes( (short) ln ), 0, dta, 2, 2 );  // ln
		int pos = 4;
		dta[pos + 1] = (byte) compression;
		dta[pos + 2] = (byte) encoding;
		//dta[pos++]= 0x5;	// means unencoded URL follows
		//dta[pos++]= (byte) (ln-1);
		pos += n;
		System.arraycopy( f, 0, dta, pos, ln - 1 );
		x.setData( dta );
		x.init();
		return x;
	}

	/**
	 * add an external sheet reference to this External Supbook record
	 *
	 * @param sheetName
	 */
	public short addExternalSheetReference( String externalSheet )
	{
		// see if sheet exists here already
		for( int i = 0; i < cstab; i++ )
		{
			if( ((String) tabs.get( i )).equalsIgnoreCase( externalSheet ) )
			{
				return (short) i;
			}
		}
		cstab++;    // increment # sheets
		System.arraycopy( ByteTools.shortToLEBytes( (short) cstab ), 0, this.getData(), 0, 2 );
		// Add new sheet reference to this SUPBOOK
		byte[] f = externalSheet.getBytes();    // get bytes of new sheet to add
		int ln = f.length/*+1*/;                    // and the length
		int encoding = 0;                        // default = no unicode // TODO: is this correct?
		int pos = this.getData().length;        // start at the end of the sheet refs
		byte[] newData = new byte[pos + ln + 3];
		System.arraycopy( this.getData(), 0, newData, 0, pos );
		System.arraycopy( ByteTools.shortToLEBytes( (short) ln ), 0, newData, pos, 2 );
		newData[pos + 2] = (byte) encoding;
		System.arraycopy( f, 0, newData, pos + 3, ln/*-1*/ );
		this.setData( newData );
		// add newest
		tabs.add( externalSheet );
		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "Supbook Sheet Reference: " + externalSheet );
		}
		return (short) (cstab - 1);
	}

	/**
	 * FOR EXTERNAL SUPBOOKS, returns the i-th sheetname
	 *
	 * @param i
	 * @return
	 */
	public String getExternalSheetName( int i )
	{
		if( i < 0 || i > tabs.size() - 1 )
		{
			return "#REF";
		}
		return (String) tabs.get( i );
	}

	/**
	 * is this an Add-in Supbook record
	 */
	public boolean isAddInRecord()
	{
		byte[] supBookCode = new byte[4];
		System.arraycopy( this.getData(), 0, supBookCode, 0, 4 );
		return (Arrays.equals( supBookCode, AddInProto ));

	}

	/**
	 * is this the global or default supbook record,
	 * delineating the total number of sheets in this workbook
	 *
	 * @return
	 */
	public boolean isGlobalRecord()
	{
		byte[] supBookCode = new byte[4];
		System.arraycopy( this.getData(), 2, supBookCode, 2, 2 );
		return (Arrays.equals( supBookCode, protocXTI ));
	}

	/**
	 * returns if this supbook record is the external supbook record,
	 * linking sheets in an external workbook
	 *
	 * @return
	 */
	public boolean isExternalRecord()
	{
		byte[] supBookCode = new byte[4];
		System.arraycopy( this.getData(), 2, supBookCode, 2, 2 );
		if( !Arrays.equals( supBookCode, protocXTI ) && !Arrays.equals( supBookCode, AddInProto ) )
		{
			return true;
		}
		return false;
	}

	/**
	 * return the External WorkBook represented by this External Supbook Record
	 * or null if it's not an external supbook
	 *
	 * @return
	 */
	public String getExternalWorkBook()
	{
		if( !isExternalRecord() )
		{
			return null;
		}
		return filename;
	}

	/**
	 * If the record is an addinrecord, update the number of sheets referred to.
	 *
	 * @see com.extentech.formats.XLS.XLSRecord#preStream()
	 */
	@Override
	public void preStream()
	{
		if( this.isGlobalRecord() )
		{
			WorkBook b = this.getWorkBook();
			if( b != null )
			{
				nSheets = b.getNumWorkSheets();
				byte[] bite = ByteTools.shortToLEBytes( (short) nSheets );
				byte[] rkdata = this.getData();
				rkdata[0] = bite[0];
				rkdata[1] = bite[1];
				this.setData( rkdata );
			}

		}
	}

	/**
	 * return String representation
	 */
	public String toString()
	{
		if( nSheets != -1 )
		{
			return "SUPBOOK: Number of Sheets: " + nSheets;
		}
		String ret = "";
		if( filename != null )
		{    // assume it's a valid external ref
			ret = "SUPBOOK: External Ref: " + filename + " Sheets:";
			for( int i = 0; i < cstab; i++ )
			{
				ret += " " + tabs.get( i );
			}
		}
		else // assume it's an add-in
		{
			return "SUPBOOK: ADD-IN";
		}
		return ret;
	}
}