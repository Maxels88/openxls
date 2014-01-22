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

import com.extentech.formats.XLS.formulas.IxtiListener;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;

import java.util.Iterator;

/**
 * <b>Externsheet: External Sheet Record (17h)</b><br>
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       cXTI        2       Number of XTI Structures
 * 6       rgXTI       var     Array of XTI Structures
 * <p/>
 * XTI
 * offset  name        size    contents
 * ---
 * 0       iSUPBOOK    2       0-based index to table of SUPBOOK records
 * 2       itabFirst   2       0-based index to first sheet tab in reference
 * 4       itabLast    2       0-based index to last sheet tab in reference
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Supbook
 */

public final class Externsheet extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4460757130836967839L;
	short cXTI = 0;
	// int DEBUGLEVEL = 10;
	CompatibleVector rgs = new CompatibleVector();

	@Override
	public void preStream()
	{
		this.update();
	}

	@Override
	public void setWorkBook( WorkBook bk )
	{
		super.setWorkBook( bk );
		Iterator it = rgs.iterator();
		while( it.hasNext() )
		{
			((rgxti) it.next()).setWorkBook( bk );
		}
	}

	public void addPtgListener( IxtiListener p ) throws WorkSheetNotFoundException
	{
		short ix = p.getIxti();
		// if(ix>0)ix--;
		if( this.rgs.size() > ix )
		{
			rgxti rg = (rgxti) this.rgs.get( ix );
			rg.addListener( p );
		}
		else
		{
			rgxti rg = new rgxti();
			rg.setWorkBook( this.wkbook );    // 20080306 KSC
			rg.setSheet1( ix );   // 20080306 KSC: use actual sheet# this.wkbook.getWorkSheetByNumber(ix));
			rg.setSheet2( ix );   // ""  this.wkbook.getWorkSheetByNumber(ix));
			this.rgs.add( rg );
			rg.addListener( p );
			this.update();
		}
	}

	public Boundsheet[] getBoundSheets( int cLoc )
	{

		if( rgs.size() == 0 )
		{
			return null;
		}

		if( cLoc > (rgs.size() - 1) )
		{
			cLoc = rgs.size() - 1;
		}

		rgxti rg = (rgxti) rgs.get( cLoc );

		short first = rg.sheet1num;
		short last = rg.sheet2num;
		if( first == (short) 0xFFFE )  // associated with a Name record
		{
			return null;
		}
		if( first == (short) 0xFFFF )  // 20080212 KSC - should be a ref to a deleted or unfound sheet
		{
			return null;    // error trap trying to get virtual sheet
		}

		int numshts = (++last) - first;
		if( numshts < 1 )
		{
			numshts = 1;
		}
		Boundsheet[] bs = new Boundsheet[numshts];
		int p = 0;
		for( int t = first; t < last; t++ )
		{
			try
			{
				bs[p++] = this.wkbook.getWorkSheetByNumber( t );
			}
			catch( WorkSheetNotFoundException e )
			{
				// don't error out on the external workbook sheet references
				if( (DEBUGLEVEL > 1) && (t != 65535) && !rg.bIsExternal ) // 20080306 KSC: add external check ...
				{
					Logger.logWarn( "Attempt to access Externsheet reference for sheet failed: " + e );
				}
			}
		}
		return bs;
	}

	/**
	 * returns array of referenced sheet names, including external references ...
	 *
	 * @param cLoc
	 * @return
	 */
	public String[] getBoundSheetNames( int cLoc )
	{
		if( rgs.size() == 0 )
		{
			return null;
		}
		if( cLoc > (rgs.size() - 1) )
		{
			cLoc = rgs.size() - 1;
		}

		rgxti rg = (rgxti) rgs.get( cLoc );

		short first = rg.sheet1num;
		short last = rg.sheet2num;

		if( first == (short) 0xFFFE )  // associated with a Name record (=Add-in)
		{
			return new String[]{ "AddIn" };
		}

		if( first == (short) 0xFFFF )  // is a ref to a deleted or un-found sheet
		{
			return new String[]{ "#REF!" };
		}

		int numshts = (++last) - first;
		if( numshts < 1 )
		{
			numshts = 1;
		}
		if( first < 0 )
		{
			first = 1;
		}

		String[] sheets = new String[numshts];
		if( first == last )
		{
			return new String[]{ "#REF!" };
		}
		int p = 0;
		for( int t = first; t < last; t++ )
		{
			try
			{
				sheets[p++] = rg.getSheetName( t );   // should successfully retrieve External Sheetnames
			}
			catch( WorkSheetNotFoundException we )
			{
				if( DEBUGLEVEL > 1 )
				{
					Logger.logWarn( "Attempt to access Externsheet reference for sheet failed: " + we );
				}
			}
		}
		return sheets;
	}

	/**
	 * returns true if the passed in sheet number is an
	 * external link (i.e. an external sheet reference)
	 *
	 * @param loc external sheet number
	 * @return
	 */
	public boolean getIsExternalLink( int loc )
	{
		if( rgs.size() == 0 )
		{
			return false;
		}

		rgxti rg = (rgxti) rgs.get( loc );
		return rg.getIsExternal();
	}

	/**
	 * get the number of refs in this Externsheet rec
	 */
	public int getcXTI()
	{
		return cXTI;
	}

	@Override
	public void init()
	{
		super.init();
		cXTI = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		int pos = 2;
		for( int t = 0; t < cXTI; t++ )
		{
			try
			{
				byte[] bts = this.getBytesAt( pos, 6 ); // System.arraycopy(rkdata,pos,bts,0,6);
				rgxti rg = new rgxti( bts );
				rgs.add( rg );
				if( wkbook != null )
				{
					rg.setWorkBook( wkbook );
				}
			}
			catch( Exception e )
			{
				if( DEBUGLEVEL > 10 )
				{
					Logger.logWarn( "init of Externsheet record failed: " + e );
				}
			}
			pos += 6;
		}
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( "Done Creating Externsheet" );
		}
	}

	/**
	 * update the underlying bytes by recreating from the properties
	 */
	void update()
	{
		// init a new byte array
		int blen = rgs.size() * 6;
		blen += 2;
		byte[] newbytes = new byte[blen];
		this.cXTI = (short) (rgs.size() - 1);
		// get the number of structs
		byte[] cx = ByteTools.shortToLEBytes( (short) rgs.size() );
		System.arraycopy( cx, 0, newbytes, 0, 2 );

		// iterate the structs and get the bytes
		int pos = 2;
		Iterator it = rgs.iterator();
		while( it.hasNext() )
		{
			byte[] btx = ((rgxti) it.next()).getBytes();
			System.arraycopy( btx, 0, newbytes, pos, 6 );
			pos += 6;
		}

		// set the data
		this.setData( newbytes );
	}

	/**
	 * Remove a sheet from this reference table,
	 * <p/>
	 * Also updates the ixti listeners, which appear to be records that contain an internal
	 * reference to a sheet number.
	 */
	void removeSheet( int sheetnum ) throws WorkSheetNotFoundException
	{
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( "Removing Sheet from Externsheet" );
		}
		// iterate the cXTI and check if this sheet is contained
		Iterator it = rgs.iterator();
		while( it.hasNext() )
		{
			rgxti rg = (rgxti) it.next();
			int first = rg.sheet1num;
			int last = rg.sheet2num;
			if( (sheetnum == first) && (sheetnum == last) )
			{
				// reference should not be removed, rather set to -1;
				rg.setSheet1( 0xFFFF );
				rg.setSheet2( 0xFFFF );
				this.cXTI--;
			}
			else if( (sheetnum >= first) && (sheetnum <= last) )
			{ // it's contained in the ref
				rg.setSheet2( rg.getSheet2() - 1 );
			}
			else if( sheetnum <= first )
			{
				rg.setSheet1( rg.getSheet1() - 1 );
				rg.setSheet2( rg.getSheet2() - 1 );
			}
		}
		this.update();
	}

    /* 20100506 KSC: this is not necessary any longer since
     * rg entries are not removed upon deletes or moved in any operations
     * therefore the ixti will stay constant 
     * -- for deleted sheets, the sheet reference will be set to the 0xFFFE deleted
     * sheet reference        
    public void notifyIxtiListeners() {
        Iterator it = rgs.iterator();
        while(it.hasNext()) {
            ((rgxti)it.next()).notifyListeners();
        }
    }
    */

	/**
	 * add a sheet to this reference table
	 * <p/>
	 * why doesn't this return the new referenced int?  NR
	 */
	void addSheet( int sheetnum ) throws WorkSheetNotFoundException
	{
		this.addSheet( sheetnum, sheetnum );
	}

	private byte getAddInIndex( Supbook[] sb )
	{
		int i = 0;
		while( i < sb.length )
		{
			if( sb[i].isAddInRecord() )
			{
				return (byte) i;
			}
			i++;
		}
		return (byte) -1;
	}

	private byte getGlobalSupBookIndex( Supbook[] sb )
	{
		int i = 0;
		while( i < sb.length )
		{
			if( sb[i].isGlobalRecord() )
			{
				return (byte) i;
			}
			i++;
		}
		return (byte) -1;
	}

	/**
	 * add a sheet range this reference table
	 */
	void addSheet( int firstSheet, int lastSheet ) throws WorkSheetNotFoundException
	{
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( "Adding new Sheet to Externsheet" );
		}

		// KSC: Added logic to set correct supbook index for added XTI
		byte[] bts = new byte[6];
		if( this.wkbook != null )
		{   // should never happen!
			Supbook[] sb = this.wkbook.getSupBooks();    // must have SUPBOOK records when have an EXTERNSHEET!
			if( firstSheet == 0xFFFE )  // then link to ADD-IN SUPBOOK
			{
				bts[0] = getAddInIndex( sb );
			}
			else
			{// link to global SUPBOOK record
				bts[0] = getGlobalSupBookIndex( sb );
			}
		}
		rgxti newcXTI = new rgxti( bts );
		newcXTI.setWorkBook( this.wkbook );      // 20080306 KSC
		if( firstSheet == 0xFFFE )  // it's a virtual sheet range for Add-ins
		{
			newcXTI.setIsAddIn( true );
		}

		if( !newcXTI.getIsAddIn() )
		{
			newcXTI.setSheet1( firstSheet );
			newcXTI.setSheet2( lastSheet );
		}
		rgs.add( newcXTI );

		this.cXTI++;
		this.update();
	}

	public short addExternalSheetRef( String externalWorkbook, String externalSheetName )
	{
		// get the external supbook record for this external workbook, creates if not present
		Supbook sb = this.wkbook.getExternalSupbook( externalWorkbook, true );
		short sheetRef = sb.addExternalSheetReference( externalSheetName );
		short sbRef = (short) this.wkbook.getSupbookIndex( sb );

         /* see if external ref  exists already */
		Iterator it = rgs.iterator();
		int i = 0;
		while( it.hasNext() )
		{
			rgxti rg = (rgxti) it.next();
			if( (rg.sheet1num == sheetRef) && (rg.sheet2num == sheetRef) && (rg.sbs == sbRef) )
			{
				return (short) i;
			}
			i++;
		}

		byte[] bts = new byte[6];
		System.arraycopy( ByteTools.shortToLEBytes( sbRef ), 0, bts, 0, 2 );   // input SUPBOOK ref #
		System.arraycopy( ByteTools.shortToLEBytes( sheetRef ), 0, bts, 2, 2 );    // input Sheet ref #
		System.arraycopy( ByteTools.shortToLEBytes( sheetRef ), 0, bts, 4, 2 );    // input Sheet ref #
		rgxti newcXTI = new rgxti( bts );
		newcXTI.setIsExternalRef( true );        // flag don't look up worksheets in this workbook
		newcXTI.setWorkBook( this.wkbook );      // 20080306 KSC
		rgs.add( newcXTI );

		this.cXTI++;
		this.update();
		return this.cXTI;
	}

	/**
	 * Insert location checks if a specific boundsheet range already has a reference.
	 * <p/>
	 * If the range already exists within the externsheet the index to the
	 * range is returned.  Else, it adds the range to the externsheet and returns the index.
	 *
	 * @param firstBound
	 * @param lastBound
	 * @return
	 */
	public int insertLocation( int firstBound, int lastBound ) throws WorkSheetNotFoundException
	{
		Iterator it = rgs.iterator();
		int i = 0;
		while( it.hasNext() )
		{
			rgxti rg = (rgxti) it.next();
			int first = rg.sheet1num;   // 20080306 KSC: getSheet1Num();
			int last = rg.sheet2num;    // getSheet2Num();

			if( (first == firstBound) && (last == lastBound) && rg.sb.isGlobalRecord() )
			{
				return i;
			}
			i++;
		}
		this.addSheet( firstBound, lastBound );
		return rgs.size() - 1;
	}

	/**
	 * Constructor
	 */
	protected static XLSRecord getPrototype()
	{
		Externsheet x = new Externsheet();
		x.setLength( (short) 8 );
		x.setOpcode( EXTERNSHEET );
		byte[] dta = new byte[8];
		dta[0] = (byte) 0x1; // put a reference to sheet1 in as initial value
		x.setData( dta );
		x.originalsize = 8;
		x.init();
		return x;
	}

	/**
	 * Add new Externsheet record and set sheet
	 *
	 * @param sheetNum1
	 * @param sheetNum2
	 * @param bk
	 * @return
	 */
	protected static XLSRecord getPrototype( int sheetNum1, int sheetNum2, WorkBook bk )
	{
		Externsheet x = (Externsheet) getPrototype();
		try
		{
			x.cXTI--;
			x.rgs.remove( 0 );
			x.setWorkBook( bk );      // must, for addSheet
//            x.addSheet(sheetNum1, sheetNum2);       
		}
		catch( Exception e )
		{
			Logger.logWarn( "ExternSheet.getPrototype error:" + e.toString() );
		}
		return x;
	}

	/**
	 * Gets the xti reference for a boundsheet name
	 *
	 * @return xti reference, -1 if not located.
	 */
	public int getXtiReference( String firstSheet, String secondSheet )
	{
		for( int i = 0; i < rgs.size(); i++ )
		{
			rgxti thisXti = (rgxti) rgs.elementAt( i );
			try
			{
				if( thisXti.getSheetName( thisXti.sheet1num ).equalsIgnoreCase( firstSheet ) && thisXti.getSheetName( thisXti.sheet2num )
				                                                                                       .equalsIgnoreCase( secondSheet ) )
				{
					return i;
				}
			}
			catch( WorkSheetNotFoundException we )
			{
				if( DEBUGLEVEL > 10 )
				{
					Logger.logWarn( "Externsheet.getXtiReference:  Attempt to find Externsheet reference for sheet failed: " + we );
				}
			}
		}
		return -1;
	}

	/**
	 * Certain records require a virtual reference, this is not a real reference to a sheet,
	 * rather an entry that is used by add in formulas, values are FE FF FE FF
	 * <p/>
	 * This method either finds the existing reference, or creates a new one and returns
	 * the pointer
	 *
	 * @return
	 */
	public int getVirtualReference()
	{
		for( int i = 0; i < rgs.size(); i++ )
		{
			rgxti thisXti = (rgxti) rgs.elementAt( i );
			if( (thisXti.sheet1num == (short) 0xFFFE) && (thisXti.sheet2num == (short) 0xFFFE) )
			{
				return i;
			}
		}
		byte[] bts = new byte[6];
		if( this.wkbook != null )
		{
			Supbook[] sb = this.wkbook.getSupBooks();
			bts[0] = getAddInIndex( sb );
		}
		rgxti newcXTI = new rgxti( bts );
		newcXTI.setWorkBook( this.wkbook );
		newcXTI.setSheet1( 0xFFFE/*null*/ );
		newcXTI.setSheet2( 0xFFFE/*null*/ );
		rgs.add( newcXTI );
		this.cXTI++;
		this.update();
		return rgs.size() - 1;
	}

	/**
	 * In some cases, we need to have an Xti reference to a non-existing sheet.  For instance, a chart
	 * or formula with a ptgRef3d that referrs to a missing sheet.  Externsheet handles this by having
	 * an internal record populated with -1's   This method searches for a non-existing record, and if
	 * that doesn't exist it creates on and passes the reference back.
	 *
	 * @return pointer to broken (-1) reference
	 */
	public int getBrokenXtiReference()
	{
		for( int i = 0; i < rgs.size(); i++ )
		{
			rgxti thisXti = (rgxti) rgs.elementAt( i );
			if( (thisXti.sheet1num == 0xFFFF) && (thisXti.sheet2num == 0xFFFF) )
			{
				return i;
			}
		}
		byte[] bts = new byte[6];
		if( this.wkbook != null )
		{
			Supbook[] sb = this.wkbook.getSupBooks();
			bts[0] = getGlobalSupBookIndex( sb );
		}
		rgxti newcXTI = new rgxti( bts );
		newcXTI.setWorkBook( this.wkbook );
		newcXTI.setSheet1( 0xFFFF );
		newcXTI.setSheet2( 0xFFFF );
		rgs.add( newcXTI );
		this.cXTI++;
		this.update();
		return rgs.size() - 1;
	}

	@Override
	public void close()
	{
		while( rgs.size() > 0 )
		{
			rgs.remove( 0 );
		}
	}

	/**
	 * Internal structure tracks Sheet references
	 */
	class rgxti implements java.io.Serializable
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = -1591367957030959727L;
		short sbs = 0;              // supbook #        
		short sheet1num, sheet2num; // store numbers as sometimes sheets are not in workbook ...
		Supbook sb = null;
		WorkBook wkbook = null;
		byte[] bts = null;
		private boolean bIsAddIn = false;   // flag if this is "iSupbook" Externsheet
		private boolean bIsExternal = false; // flag if this is an external ref.
		// contains references to virtual 0xFFFE sheet #

		CompatibleVector listeners = new CompatibleVector();

		public void addListener( IxtiListener p )
		{
			listeners.add( p );
		}

		public void notifyListeners()
		{
			Iterator it = listeners.iterator();
			int locp = rgs.indexOf( this );
			while( it.hasNext() )
			{
				IxtiListener p = (IxtiListener) it.next();
				if( locp == -1 )
				{
					((Ptg) p).getParentRec().remove( true ); // this reference has been deleted, rec not valid
				}
				else
				{
					p.setIxti( (short) locp );
				}
			}
		}

		/**
		 * default constructor
		 */
		rgxti()
		{
			// do something?
		}

		void setWorkBook( WorkBook bk )
		{
			this.wkbook = bk;
			sbs = ByteTools.readShort( bts[0], bts[1] );           // supbook #
			sb = wkbook.getSupBooks()[sbs];                      // get supbook referenced by sbs

			sheet1num = ByteTools.readShort( bts[2], bts[3] );
			sheet2num = ByteTools.readShort( bts[4], bts[5] );

			if( sheet1num == 0xFFFE )
			{
				bIsAddIn = true;
			}

			if( sb.isExternalRecord() )
			{
				bIsExternal = true;
			}

		}

		/**
		 * constructor used to init from bytes without a book
		 *
		 * @param initbytes
		 */
		rgxti( byte[] initbytes )
		{
			this.bts = initbytes;
		}

		/**
		 * @return Returns the bts.
		 */
		public byte[] getBytes()
		{
			if( bts == null )
			{
				bts = new byte[6];
			}
			byte[] shtbt = ByteTools.shortToLEBytes( sheet1num );
			System.arraycopy( shtbt, 0, bts, 2, 2 );
			shtbt = ByteTools.shortToLEBytes( sheet2num );
			System.arraycopy( shtbt, 0, bts, 4, 2 );
			return bts;
		}

		/**
		 * return the sheet name for the sheetNum in the associated supbook
		 */
		public String getSheetName( int sheetNum ) throws WorkSheetNotFoundException
		{
			if( bIsAddIn )    // no sheets assoc; return virtual sheet #
			{
				return "Virtual Sheet Range 0xFFFE - 0xFFFE";
			}
			if( bIsExternal )
			{
				return sb.getExternalSheetName( sheetNum );
			}
			if( sheetNum == 0xFFFF )
			{
				return "Deleted Sheet";
			}
			//otherwise, try and get sheetname
			return wkbook.getWorkSheetByNumber( sheetNum ).getSheetName();
            
            /*
            if(this.sheet1==null)   // if sheet ref is deleted, return 0xFFFF
                return "null";  //-1;
            return this.sheet1.getSheetName();
            */
		}

		public void setIsAddIn( boolean bIsAddin )
		{
			this.bIsAddIn = bIsAddin;
		}

		public boolean getIsAddIn()
		{
			return bIsAddIn;
		}

		/**
		 * set if this rgi references an external SUPBOOK record
		 * i.e. one that references an External Workbook
		 * (therefore sheet references are not found in the current workbook)
		 *
		 * @param bIsExternal
		 */
		public void setIsExternalRef( boolean bIsExternal )
		{
			this.bIsExternal = bIsExternal;
		}

		/**
		 * return if this rgi references an external SUPBOOK record
		 * i.e. one that references an External Workbook
		 * (therefore sheet references are not found in the current workbook)
		 */
		public boolean getIsExternal()
		{
			return this.bIsExternal;
		}

		/**
		 * set the sheet# referenced
		 *
		 * @param int sheet #
		 */
		public void setSheet1( int sh1 )
		{
			this.sheet1num = (short) sh1;
		}

		public int getSheet1()
		{
			return (int) this.sheet1num;
		}

		/**
		 * set the sheet# referenced
		 *
		 * @param int sheet #
		 */
		public void setSheet2( int sh2 )
		{
			this.sheet2num = (short) sh2;
		}

		public int getSheet2()
		{
			return (int) this.sheet2num;
		}

		public String toString()
		{

			try
			{
				return "rgxti range: " + this.getSheetName( sheet1num ) + "-" + this.getSheetName( sheet2num );
			}
			catch( WorkSheetNotFoundException we )
			{
			}
			return "rgxti range: sheets not initialized";
		}
	}
}