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
package com.extentech.formats.XLS.formulas;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Externsheet;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.Shrfmla;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;

/**
 * ptgArea3d is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 * implies external sheet ref (rather than ptgarea)
 * <p/>
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           ixti        2       index into the Externsheet record
 * 2           rwFirst     2       The First row of the reference
 * 4           rwLast      2       The Last row of the reference
 * 6           grbitColFirst   2       (see following table)
 * 8           grbitColLast    2       (see following table)
 *
 * Only the low-order 14 bits specify the Col, the other bits specify
 * relative vs absolute for both the col or the row.
 *
 * Bits        Mask        Name    Contents
 * -----------------------------------------------------
 * 15          8000h       fRwRel  =1 if row offset relative,
 * =0 if otherwise
 * 14          4000h       fColRel =1 if row offset relative,
 * =0 if otherwise
 * 13-0        3FFFh       col     Ordinal column offset or number
 * </pre>
 * <p/>
 * For 3D references, the tokens contain a negative EXTERNSHEET index, indicating a reference into the own workbook.
 * The absolute value is the one-based index of the EXTERNSHEET record that contains the name of the first sheet. The
 * tokens additionally contain absolute indexes of the first and last referenced sheet. These indexes are independent of the
 * EXTERNSHEET record list. If the referenced sheets do not exist anymore, these indexes contain the value FFFFH (3D
 * reference to a deleted sheet), and an EXTERNSHEET record with the special name <04H> (own document) is used.
 * Each external reference contains the positive one-based index to an EXTERNSHEET record containing the URL of the
 * external document and the name of the sheet used. The sheet index fields of the tokens are not used.
 * <p/>
 * is the above correct?? Documentation sez different!!!!!
 *
 * @see Ptg
 * @see Formula
 */
public class PtgArea3d extends PtgArea implements Ptg, IxtiListener
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1176168076050592292L;

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	@Override
	public boolean getIsReference()
	{
		return true;
	}

	boolean quoted = false;
	public short ixti;
	private boolean isExternalRef = false;    // true if this ptg area is a reference in another workbook
	private Ptg[] comps = null;

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 */
	@Override
	public String getString()
	{
		try
		{
			if( this.getIsWholeCol() || this.getIsWholeRow() )
			{    // handle non-standard ranges i.e. $B:$C or $1:$3
				String s = firstPtg.getLocation();
				String y = lastPtg.getLocation();

				String loc1[] = ExcelTools.stripSheetNameFromRange( s );    // sheet, addr
				String loc2[] = ExcelTools.stripSheetNameFromRange( y );    // sheet, addr
				if( this.getIsWholeCol() )
				{
					int i = loc1[1].length();
					if( Character.isDigit( loc1[1].charAt( i - 1 ) ) )
					{
						while( Character.isDigit( loc1[1].charAt( --i ) ) )
						{
							;
						}
					}
					loc1[1] = loc1[1].substring( 0, i );
					i = loc2[1].length();
					if( Character.isDigit( loc2[1].charAt( i - 1 ) ) )
					{
						while( Character.isDigit( loc2[1].charAt( --i ) ) )
						{
							;
						}
					}
					loc2[1] = loc2[1].substring( 0, i );
				}
				else if( this.getIsWholeRow() )
				{
					int i = 0;
					while( !Character.isDigit( loc1[1].charAt( i++ ) ) )
					{
						;
					}
					loc1[1] = "$" + loc1[1].substring( i - 1 );
					i = 0;
					while( !Character.isDigit( loc2[1].charAt( i++ ) ) )
					{
						;
					}
					loc2[1] = "$" + loc2[1].substring( i - 1 );
				}
				sheetname = qualifySheetname( sheetname );

				return sheetname + "!" + loc1[1] + ":" + loc2[1];
			}
			// otherwise, 
			return getLocation();
		}
		catch( Exception e )
		{
			;// is ok if new ptg...
			return null;
		}
	}

	/**
	 * link to the externsheet to be automatically updated upon removals
	 */
	@Override
	public void addListener()
	{
		try
		{
			getParentRec().getWorkBook().getExternSheet().addPtgListener( this );
		}
		catch( Exception e )
		{
			// no need to output here.  NullPointer occurs when a ref has an invalid ixti, such as when a sheet was removed  Worksheet exception could never really happen.
		}
	}

	public String toString()
	{
		String ret = getString();
		return ret;
	}

	/**
	 * create new PtgArea3d
	 */
	public PtgArea3d()
	{
		ptgId = 0x3b;
		record = new byte[11];
		record[0] = ptgId; // ""
		this.is3dRef = true;
	}

	public PtgArea3d( boolean useReferenceTracker )
	{
		this();
		this.useReferenceTracker = useReferenceTracker;
	}

	/**
	 * set the Ptg Id type to one of:
	 * VALUE, REFERENCE or Array
	 * 25H (tAreaR), 45H (tAreaV), 65H (tAreaA)
	 * <br>The Ptg type is important for certain
	 * functions which require a specific type of operand
	 */
	@Override
	public void setPtgType( short type )
	{
		switch( type )
		{
			case VALUE:
				ptgId = 0x5B;
				break;
			case REFERENCE:
				ptgId = 0x3B;
				break;
			case Ptg.ARRAY:
				ptgId = 0x7B;
				break;
		}
		record[0] = ptgId;
	}

	/**
	 * @return Returns the ixti.
	 */
	@Override
	public short getIxti()
	{    // only valid for 3d refs!!!!
		return ixti;
	}

	/**
	 * set the pointer into the Externsheet Rec.
	 * this is only valid for 3d refs
	 */
	@Override
	public void setIxti( short ixf )
	{
		if( ixti != ixf )
		{
			ixti = ixf;
			// this seems to be only one byte...
			if( record != null )
			{
				record[1] = (byte) ixf;
				populateVals();    // add listener is done here
			}
		}
	}

	/**
	 * return true if this PtgArea3d is an external reference
	 * i.e. defined in another, external workbook
	 *
	 * @return
	 */
	public boolean isExternalRef()
	{
		return isExternalRef;
	}

	/**
	 * return the first sheet referenced
	 *
	 * @return
	 */
	public Boundsheet getFirstSheet()
	{
		if( parent_rec != null )
		{
			WorkBook wb = parent_rec.getWorkBook();
			if( sheetname != null )
			{
				try
				{
					return wb.getWorkSheetByName( sheetname );
				}
				catch( WorkSheetNotFoundException e )
				{
					;    // fall thru -- see sheet copy operations -- appears correct
				}
			}
			if( (wb != null) && (wb.getExternSheet() != null) )
			{
				Boundsheet[] bsa = wb.getExternSheet().getBoundSheets( this.ixti );
				if( bsa != null )
				{
					return bsa[0];
				}
			}
		}
		return null;
	}

	/**
	 * return the last sheet referenced
	 *
	 * @return
	 */
	public Boundsheet getLastSheet()
	{
		if( parent_rec != null )
		{
			WorkBook wb = parent_rec.getWorkBook();
			if( (wb != null) && (wb.getExternSheet() != null) )
			{
				Boundsheet[] bsa = wb.getExternSheet().getBoundSheets( this.ixti );
				if( bsa != null )
				{
					if( bsa.length > 1 )
					{
						return bsa[bsa.length - 1];
					}
					return bsa[0];
				}

			}
		}
		return null;
	}

	@Override
	public void setParentRec( XLSRecord rec )
	{
		super.setParentRec( rec );
		if( firstPtg != null )
		{
			firstPtg.setParentRec( parent_rec );
		}
		if( lastPtg != null )
		{
			lastPtg.setParentRec( parent_rec );
		}
	}

	/**
	 * return the first sheet referenced
	 *
	 * @return
	 */
	public Boundsheet getSheet()
	{
		return getFirstSheet();
	}

	/**
	 * get the sheet name from the 1st 3d reference
	 */
	@Override
	public String getSheetName()
	{
		if( sheetname == null )
		{
			if( parent_rec != null )
			{
				WorkBook wb = parent_rec.getWorkBook();
				if( (wb != null) && (wb.getExternSheet() != null) )
				{
					String[] sheets = wb.getExternSheet().getBoundSheetNames( this.ixti );
					if( (sheets != null) && (sheets.length > 0) )
					{
						sheetname = sheets[0];
						sheetname = qualifySheetname( sheetname );
					}
				}

				if( (sheetname == null) && (parent_rec != null) && (parent_rec.getSheet() != null) )
				{    // try this:
					sheetname = parent_rec.getSheet().getSheetName();
					sheetname = qualifySheetname( sheetname );
				}
			}
		}
		return sheetname;
	}

	/**
	 * return the name of the last sheet referenced if it's an external ref
	 *
	 * @return
	 */
	public String getLastSheetName()
	{
		String sheetname = this.sheetname;    // 20100217 KSC: default to 1st sheet
		if( parent_rec != null )
		{
			WorkBook wb = parent_rec.getWorkBook();
			if( (wb != null) && (wb.getExternSheet() != null) )
			{
				String[] sheets = wb.getExternSheet().getBoundSheetNames( this.ixti );
				if( (sheets != null) && (sheets.length > 0) )
				{
					sheetname = sheets[sheets.length - 1];
				}
			}
		}
		return sheetname;
	}

	/**
	 * get the worksheet that this ref is on
	 */
	public Boundsheet[] getSheets( WorkBook b )
	{
		Boundsheet[] bsa = b.getExternSheet().getBoundSheets( this.ixti );
		if( bsa[0] == null ) // 20080303 KSC: Catch Unresolved External refs
		{
			Logger.logErr( "PtgArea3d.getSheet: Unresolved External Worksheet" );
		}
		return bsa;
	}

	/**
	 * constructor, takes the array of the ptgRef, including
	 * the identifier so we do not need to figure it out again later...
	 * also takes the parent rec -- needed to init the sub-ptgs
	 *
	 * @param b
	 * @param parent
	 */
	public void init( byte[] b, XLSRecord parent )
	{
		ixti = ByteTools.readShort( b[1], b[2] );
		record = b;
		this.setParentRec( parent );
		populateVals();
	}

	/**
	 * Throw this data into two ptgref's
	 */
	@Override
	public void populateVals()
	{
		byte[] temp1 = new byte[7];    // PtgRef3d is 7 bytes
		byte[] temp2 = new byte[7];
		// Encoded Cell Range Address:
		// 0-2= first row
		// 2-4= last row
		// 4-6= first col
		// 6-8= last col
		// Encoded Cell Address:
		// 0-2=	row index
		// 2-4= col index + relative flags
		try
		{
			temp1[0] = 0x3a;
			temp1[1] = record[1];    // ixti
			temp1[2] = record[2];    // ""
			temp1[3] = record[3];    // first row
			temp1[4] = record[4];    // ""
			temp1[5] = record[7];    // first col
			temp1[6] = record[8];    // ""

			temp2[0] = 0x3a;
			temp2[1] = record[1];        // ixti
			temp2[2] = record[2];        // ""
			temp2[3] = record[5];        // last row
			temp2[4] = record[6];        // ""
			temp2[5] = record[9];        // last col
			temp2[6] = record[10];        //	""
		}
		catch( Exception e )
		{
			//should never happen!
			return;
		}
		// pass in parent_rec so can properly set formulaRow/formulaCol
		firstPtg = new PtgRef3d( false );

		// the following method registers the Ptg with the ReferenceTracker
		firstPtg.setParentRec( parent_rec );
		firstPtg.setSheetName( this.getSheetName() );
		firstPtg.init( temp1 );

		lastPtg = new PtgRef3d( false );
		lastPtg.setParentRec( parent_rec );
		lastPtg.setSheetName( this.getLastSheetName() );
		lastPtg.init( temp2 );
		// flag if it's an external reference

		isExternalRef = (((PtgRef3d) firstPtg).isExternalLink() || ((PtgRef3d) lastPtg).isExternalLink());

		setWholeRowCol();

		// take 1st Ptg as sample for relative state
		this.fColRel = firstPtg.isColRel();
		this.fRwRel = firstPtg.isRowRel();
		//init sets formula row to 1st row for a shared formula; adjust here
		if( (parent_rec != null) && (parent_rec instanceof Shrfmla) )
		{
			lastPtg.formulaRow = ((Shrfmla) parent_rec).getLastRow();
			lastPtg.formulaCol = ((Shrfmla) parent_rec).getLastCol();
		}
		((PtgArea) this).hashcode = super.getHashCode();
	}

	/**
	 * Set the location of this PtgRef.  This takes a location
	 * such as "a14:b15"
	 * <p/>
	 * NOTE: the reference stays on the same sheet!
	 */
	@Override
	public void setLocation( String address )
	{
		String[] s = ExcelTools.stripSheetNameFromRange( address );
		setLocation( s );
	}

	/**
	 * set Ptg Location to parsed location
	 *
	 * @param loc String[] sheet1, range, sheet2, exref1, exref2
	 */
	@Override
	public void setLocation( String[] s )
	{
		try
		{
			if( useReferenceTracker && (locax != null) )    // if in tracker already, remove
			{
				this.removeFromRefTracker();
			}
		}
		catch( Exception e )
		{
			;// will happen if this is not in tracker yet
		}
		String sheetname2 = null;
		String range = "";
		range = s[1];
		if( s[0] != null )
		{    // has a sheet in the address
			sheetname = s[0];
			sheetname2 = s[2];
			if( sheetname2 == null )
			{
				sheetname2 = sheetname;
			}
			// revised so can set ixti on error'd references
			WorkBook b = null;
			Externsheet xsht = null;
			if( parent_rec != null )
			{
				b = parent_rec.getWorkBook();
				if( b == null )
				{
					b = parent_rec.getSheet().getWorkBook();
				}
			}
			try
			{
				xsht = b.getExternSheet();
				int boundnum = b.getWorkSheetByName( sheetname ).getSheetNum();
				int boundnum2 = boundnum; // it could possibly be a 3d ref - check
				if( !sheetname.equals( sheetname2 ) && (sheetname2 != null) )
				{
					boundnum2 = b.getWorkSheetByName( sheetname2 ).getSheetNum();
				}
				this.setIxti( (short) xsht.insertLocation( boundnum, boundnum2 ) );
			}
			catch( WorkSheetNotFoundException e )
			{
				try
				{
					// try to link to external sheet, if possible
					int boundnum = xsht.getXtiReference( s[0], s[0] );
					if( boundnum == -1 )
					{    // can't resolve
						this.setIxti( (short) xsht.insertLocation( boundnum, boundnum ) );
					}
					else
					{
						this.setIxti( (short) boundnum );
						this.isExternalRef = true;
					}
				}
				catch( Exception ex )
				{
				}
			}
		}
		else if( parent_rec != null )
		{
			sheetname = sheetname2 = parent_rec.getSheet().getSheetName();    // use parent rec's sheet
		}
		int i = range.indexOf( ":" );
		if( i < 0 )
		{
			range = range + ":" + range;
			i = range.indexOf( ":" );
		}

		String firstcell = range.substring( 0, i );
		String lastcell = range.substring( i + 1 );
		if( sheetname != null )
		{
			firstcell = sheetname + "!" + firstcell;
		}
		if( sheetname2 != null )
		{
			lastcell = sheetname2 + "!" + lastcell;
		}
		if( s[3] != null )        // store OOXML External References
		{
			firstcell = s[3] + firstcell;
		}
		if( s[4] != null )
		{
			lastcell = s[4] + lastcell;
		}

		if( firstPtg == null )
		{
			firstPtg = new PtgRef3d( false );
			firstPtg.setParentRec( this.getParentRec() );
		}
		firstPtg.sheetname = sheetname;
		((PtgRef3d) firstPtg).setLocation( firstcell );
		((PtgRef3d) firstPtg).setIxti( this.ixti );

		if( lastPtg == null )
		{
			lastPtg = new PtgRef3d( false );
			lastPtg.setParentRec( this.getParentRec() );
		}
		lastPtg.sheetname = sheetname2;
		((PtgRef3d) lastPtg).setLocation( lastcell );
		((PtgRef3d) lastPtg).setIxti( this.ixti );

		this.setWholeRowCol();
		this.updateRecord();
		// TODO: must deal with non-symmetrical absolute i.e. if first and last ptgs don't match
		this.fRwRel = firstPtg.fRwRel;
		this.fColRel = firstPtg.fColRel;
		hashcode = getHashCode();
		if( useReferenceTracker )
		{
			if( !this.getIsWholeCol() && !this.getIsWholeRow() )
			{
				this.addToRefTracker();
			}
			else
			{
				useReferenceTracker = false;
			}
		}
	}

	/**
	 * returns the location of the ptg as an array of shorts.
	 * [0] = firstrow
	 * [1] = firstcol
	 * [2] = lastrow
	 * [3] = lastcol
	 */
	@Override
	public int[] getIntLocation()
	{
		int[] first = firstPtg.getIntLocation();
		int[] last = lastPtg.getIntLocation();
		int[] returning = new int[4];
		System.arraycopy( first, 0, returning, 0, 2 );
		System.arraycopy( last, 0, returning, 2, 2 );
		return returning;
	}

	/**
	 * returns whether this CellRange Contains a Cell
	 *
	 * @param the cell to test
	 * @return whether the cell is in the range
	 */
	@Override
	public boolean contains( CellHandle ch )
	{
		String chsheet = ch.getWorkSheetName();
		getSheetName();
		if( !chsheet.equalsIgnoreCase( sheetname ) )
		{
			return false;
		}
		String adr = ch.getCellAddress();
//      FIX broken COLROW
		int[] rc = ExcelTools.getRowColFromString( adr );
		return contains( rc );
	}

	/**
	 * Switches the two internal ptgref3ds to a new
	 * sheet.
	 */
	public void setReferencedSheet( Boundsheet b )
	{
		((PtgRef3d) firstPtg).setReferencedSheet( b );
		((PtgRef3d) lastPtg).setReferencedSheet( b );
		int boundnum = b.getSheetNum();
		Externsheet xsht = b.getWorkBook().getExternSheet( true );
		//TODO: add handling for multi-sheet reference.  Already handled in externsheet
		try
		{
			this.sheetname = null;    // 20100218 KSC: RESET
			int xloc = xsht.insertLocation( boundnum, boundnum );
			setIxti( (short) xloc );
		}
		catch( WorkSheetNotFoundException e )
		{
			Logger.logErr( "Unable to set referenced sheet in PtgRef3d " + e );
		}
	}

	/**
	 * return all of the Ptg values represented in this array
	 * <p/>
	 * will have to reference the workbook cells as well as
	 * any upstream formulas...
	 *
	 * @return
	 */
	public Object[] getAllVals()
	{

		return null;
	}

	/**
	 * Updates the record bytes so it can be pulled back out.
	 */
	@Override
	public void updateRecord()
	{
		comps = null;
		byte[] first = firstPtg.getRecord();
		byte[] last = lastPtg.getRecord();
		// KSC: this apparently is what excel wants:
		if( wholeRow )
		{
			first[5] = 0;
		}
		if( wholeCol )
		{
			first[3] = 0;
			first[4] = 0;
		}
		// the last record has an extra identifier on it.
		byte[] newrecord = new byte[PTG_AREA3D_LENGTH];
		newrecord[0] = 0x3B;
		System.arraycopy( first, 1, newrecord, 1, 2 );
		System.arraycopy( first, 3, newrecord, 3, 2 );
		System.arraycopy( last, 3, newrecord, 5, 2 );
		System.arraycopy( first, 5, newrecord, 7, 2 );
		System.arraycopy( last, 5, newrecord, 9, 2 );
		record = newrecord;
		if( parent_rec != null )
		{
			if( this.parent_rec instanceof Formula )
			{
				((Formula) this.parent_rec).updateRecord();
			}
			else if( this.parent_rec instanceof Name )
			{
				((Name) this.parent_rec).updatePtgs();
			}
		}
	}

	@Override
	public int getLength()
	{
		return PTG_AREA3D_LENGTH;
	}

	/**
	 * Returns all of the cells of this range as PtgRef's.
	 * This includes empty cells, values, formulas, etc.
	 * Note the setting of parent-rec requires finding the cell
	 * the PtgRef refer's to.  If that is null then the PtgRef
	 * will exist, just with a null value.  This could cause issues when
	 * programatically populating cells.
	 */
	@Override
	public Ptg[] getComponents()
	{
		if( comps != null )
		{
			return comps;
		}

		ArrayList<Ptg> components = new ArrayList<Ptg>();
		try
		{
			// loop through the cols
			String sht = "";
			if( this.toString().indexOf( "!" ) > -1 )
			{
				sht = this.toString();
				sht = sht.substring( 0, sht.indexOf( "!" ) ) + "!";
			}
			int startrow = 0, startcol = 0, endrow = 0, endcol = 0;
			if( !this.wholeCol && !this.wholeRow )
			{ // normal case
//				 TODO: check rc sanity here
				int[] startloc = firstPtg.getRealIntLocation();    // Get Actual Coordinates
				startcol = startloc[1];
				startrow = startloc[0];
				int[] endloc = lastPtg.getRealIntLocation();    // Get Actual Coordinates
				endcol = endloc[1];
				endrow = endloc[0];
			}
			else if( this.wholeRow )
			{        // like $1:$1
				startcol = 0;
				try
				{
					endcol = this.getSheet().getMaxCol();
				}
				catch( NullPointerException ne )
				{ // can happens when Name record is being init'd and sheet records are not set yet
					return null;
				}
				startrow = endrow = firstPtg.rw;
			}
			else if( this.wholeCol )
			{        // like $J:$J
				startrow = 0;    // Get Actual Coordinates
				startcol = endcol = firstPtg.col;
				try
				{
					endrow = this.getSheet().getMaxRow();
				}
				catch( NullPointerException ne )
				{ // can happens when Name record is being init'd and sheet records are not set yet
					return null;
				}
			}
			for(; startcol <= endcol; startcol++ )
			{
				// loop through the rows inside
				int rowholder = startrow;
				for(; rowholder <= endrow; rowholder++ )
				{
					String displaycol = ExcelTools.getAlphaVal( startcol );
					int displayrow = rowholder + 1;
					String loc = sht + displaycol + displayrow;

					// cache these suckers!
					Ptg pref = new PtgRef3d( false );
					pref.setParentRec( parent_rec );    // must set parentrec before setLocation
					((PtgRef3d) pref).setLocation( loc );
					components.add( pref );
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "calculating range value in PtgArea3d failed.", e );
		}
		PtgRef[] pref = new PtgRef[components.size()];
		components.toArray( pref );
		comps = pref;
		return comps;
	}

	/**
	 * sets the column to be relative (relative is true) or absolute (relative is false)
	 * <br>absolute references do not shift upon column inserts or deletes
	 * <br>NOTE: DOES NOT handle asymmetrical ranges i.e. 1st ref is absolute, 2nd is relative
	 *
	 * @param boolean relative
	 */
	@Override
	public void setColRel( boolean relative )
	{
		this.fColRel = relative;
		firstPtg.setColRel( relative );
		lastPtg.setColRel( relative );
		updateRecord();
	}

	/**
	 * sets the row to be relative (relative is true) or absolute (relative is false)
	 * <br>absolute references do not shift upon row inserts or deletes
	 * <br>NOTE: DOES NOT handle asymmetrical ranges i.e. 1st ref is absolute, 2nd is relative
	 *
	 * @param boolean relative
	 */
	@Override
	public void setRowRel( boolean relative )
	{
		if( this.fRwRel != relative )
		{
			this.fRwRel = relative;
			firstPtg.setRowRel( relative );
			lastPtg.setRowRel( relative );
			updateRecord();
		}
	}

	@Override
	public void close()
	{
		super.close();
		if( comps != null )
		{
			for( int i = 0; i < comps.length; i++ )
			{
				((GenericPtg) comps[i]).close();
				comps[i] = null;
			}
		}
	}

}