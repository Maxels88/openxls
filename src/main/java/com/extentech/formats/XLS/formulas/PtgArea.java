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
import com.extentech.formats.XLS.Array;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.Row;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.FastAddVector;
import com.extentech.toolkit.Logger;

import java.util.Vector;

/**
 * ptgArea is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 * <p/>
 * <pre>
 * Offset      Size    Contents
 * ----------------------------------------------------
 * 0			2 		Index to first row (065535) or offset of first row (method [B], -3276832767)
 * 2 			2 		Index to last row (065535) or offset of last row (method [B], -3276832767)
 * 4 			2 		Index to first column or offset of first column, with relative flags (see table above)
 * 6 			2 		Index to last column or offset of last column, with relative flags (see table above)
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
 *
 * @see Ptg
 * @see Formula
 */
public class PtgArea extends PtgRef implements Ptg
{

	public static final long serialVersionUID = 666555444333222l;

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

	protected PtgRef firstPtg;
	protected PtgRef lastPtg;

	/* constructor, takes the array of the ptgRef, including
	the identifier so we do not need to figure it out again later...
	*/
	@Override
	public void init( byte[] b )
	{
		locax = null; // cache reset
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	/*
	 Throw this data into two ptgref's
	*/
	@Override
	public void populateVals()
	{
		byte[] temp1 = new byte[5];
		byte[] temp2 = new byte[5];
		temp1[0] = 0x24;
		temp2[0] = 0x24;
		System.arraycopy( record, 1, temp1, 1, 2 );
		System.arraycopy( record, 5, temp1, 3, 2 );
		System.arraycopy( record, 3, temp2, 1, 2 );
		System.arraycopy( record, 7, temp2, 3, 2 );
		try
		{
			getSheetName(); // 20080212 KSC:
		}
		catch( WorkSheetNotFoundException we )
		{
			Logger.logErr( we );
		}
		firstPtg = new PtgRef( temp1, parent_rec, false );    // don't add to ref tracker as it's part of area
		firstPtg.sheetname = sheetname;

		lastPtg = new PtgRef( temp2, parent_rec, false );        // don't add to ref tracker as it's part of area
		lastPtg.sheetname = sheetname;
		setWholeRowCol();
		this.hashcode = getHashCode();
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
		Vector v = new Vector();
		try
		{
//       TODO: check rc sanity here
			int startcol = -1;
			int startrow = -1;
			int endrow = -1;
			int endcol = -1;
			int[] startloc = null;
			int[] endloc = null;

        /*if (this.wholeRow) {
			startcol= 0;
			endcol= this.getSheet().getMaxCol();
			startrow= endrow= firstPtg.rw;
        } if (this.wholeCol) {
		    startrow= 0;	// Get Actual Coordinates
			startcol= endcol= firstPtg.col;
			endrow= this.getSheet().getMaxRow();
        } */
			if( firstPtg != null )
			{
				startloc = firstPtg.getRealIntLocation();
				startcol = startloc[1];
				startrow = startloc[0];
			}
			else
			{
				startloc = ExcelTools.getRangeRowCol( locax );
				startcol = startloc[1];
				startrow = startloc[0];
			}

			if( lastPtg != null )
			{
				endloc = lastPtg.getRealIntLocation();
				endcol = endloc[1];
				endrow = endloc[0];
			}
			else
			{
				endloc = ExcelTools.getRangeRowCol( locax );
				endcol = endloc[3];
				endrow = endloc[2];
			}

			// usually don't need to set sheet on setlocation becuase uses parent_rec's sheet
			// cases of named range or if location sheet does not = parent_rec sheet, set sheet explicitly
			String sht = null;    // usual case, don't need to set sheet
			Boundsheet sh = parent_rec.getSheet();
			if( (sh == null) || ((this.sheetname != null) && !this.sheetname.equals( sh.getSheetName() )) )
			{
				if( (sh == null) || !GenericPtg.qualifySheetname( this.sheetname )
				                               .equals( GenericPtg.qualifySheetname( sh.getSheetName() ) ) )
				{
					sht = this.sheetname + "!";
				}
			}
			// loop through the cols
			for(; startcol <= endcol; startcol++ )
			{
				// loop through the rows inside
				int rowholder = startrow;
				for(; rowholder <= endrow; rowholder++ )
				{
					String displaycol = ExcelTools.getAlphaVal( startcol );
					int displayrow = rowholder + 1;
					PtgRef pref;
					if( sht == null )
					{
						pref = new PtgRef( displaycol + displayrow, parent_rec, false );
					}
					else
					{
						pref = new PtgRef( sht + displaycol + displayrow, parent_rec, false );
					}
					v.add( pref );
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "calculating formula range value failed.", e );
		}
		PtgRef[] pref = new PtgRef[v.size()];
		v.toArray( pref );
		return pref;
	}

	/**
	 * returns the row/col ints for the ref
	 * <p/>
	 * Format is FirstRow,FirstCol,LastRow,LastCol
	 *
	 * @return
	 */
	@Override
	public int[] getRowCol()
	{
		if( firstPtg == null )
		{
			return null;
		}
		if( (lastPtg == null) && (firstPtg != null) )
		{
			int[] rc1 = firstPtg.getRowCol();
			int[] ret = { rc1[0], rc1[1], rc1[0], rc1[1] };
			return ret;
		}
		int[] rc1 = firstPtg.getRowCol();
		int[] rc2 = lastPtg.getRowCol();
		int[] ret = { rc1[0], rc1[1], rc2[0], rc2[1] };
		return ret;
	}

	/**
	 * returns whether this CellRange Contains a Cell
	 *
	 * @param the cell to test
	 * @return whether the cell is in the range
	 */
	public boolean contains( CellHandle ch )
	{
		String chsheet = ch.getWorkSheetName();
		String mysheet = "";
		if( this.getParentRec() != null )
		{
			BiffRec b = this.getParentRec();
			if( b.getSheet() != null )
			{
				mysheet = b.getSheet().getSheetName();
			}
		}
		if( !chsheet.equalsIgnoreCase( mysheet ) )
		{
			return false;
		}
		String adr = ch.getCellAddress();
//      FIX broken COLROW
		int[] rc = ExcelTools.getRowColFromString( adr );
		return contains( rc );
	}

	/**
	 * check to see if the sheet and row/col are contained
	 * in this ref
	 *
	 * @param sheetname
	 * @param rc
	 * @return
	 */
	public boolean contains( String sn, int[] rc )
	{
		if( sheetname == null )
		{
			try
			{
				sheetname = this.getSheetName();
			}
			catch( Exception e )
			{
				;
			}
		}
		if( !sn.equalsIgnoreCase( sheetname ) )
		{
			return false;
		}
		return contains( rc );
	}

	/**
	 * returns whether this PtgArea Contains the specified row/col coordinate
	 * <p/>
	 * <p/>
	 * [0] = firstrow
	 * [1] = firstcol
	 * [2] = lastrow
	 * [3] = lastcol
	 *
	 * @param the rc coordinates to test
	 * @return whether the coordinates are in the range
	 */
	public boolean contains( int[] rc )
	{
		int[] thisRange = this.getIntLocation();
		// test the first rc
		if( rc[0] < thisRange[0] )
		{
			return false; // row above the first ref row?
		}
		if( rc[0] > thisRange[2] )
		{
			return false; // row after the last ref row?
		}

		if( rc[1] < thisRange[1] )
		{
			return false; // col before the first ref col?
		}
		if( rc[1] > thisRange[3] )
		{
			return false; // col after the last ref col?
		}

		return true;
	}

	// private byte[] PROTOTYPE_BYTES = {0x25, 0, 0, 0, 0, 0, 0, 0, 0};

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 */
	@Override
	public String getString()
	{
		return this.getLocation();
	}

	/*
	 * Creates a new PtgArea.  The parent rec is needed
	 * as getting a value goes to the boundsheet to determine
	 * values.  The parent rec *must* be on the same sheet
	 * as the PtgArea referenced!
	 */
	public PtgArea( String range, XLSRecord parent )
	{
		this( range, parent, true );
	}

	/**
	 * Creates a new PtgArea from 2 component ptgs, used by shared formula
	 * to create ptgareas.  ptg1 should be upperleft corner, ptg2 bottomright
	 */
	public PtgArea( PtgRef ptg1, PtgRef ptg2, XLSRecord parent )
	{
		this();
		firstPtg = ptg1;
		lastPtg = ptg2;
		parent_rec = parent;
		this.hashcode = getHashCode();
		this.updateRecord();
	}

	/*
	 * Creates a new PtgArea.  The parent rec is needed
	 * as getting a value goes to the boundsheet to determine
	 * values.  The parent rec *must* be on the same sheet
	 * as the PtgArea referenced!
	 *
	 * relativeRefs = true is excel default
	 */
	public PtgArea( String range, XLSRecord parent, boolean relativeRefs )
	{
		this();
		int[] loc = ExcelTools.getRangeRowCol( range );
		int[] temp = new int[2];
		temp[0] = loc[0];
		temp[1] = loc[1];
		String res = ExcelTools.formatLocation( temp, relativeRefs, relativeRefs );
		firstPtg = new PtgRef( res, parent, false );
		temp[0] = loc[2];
		temp[1] = loc[3];
		res = ExcelTools.formatLocation( temp, relativeRefs, relativeRefs );
		lastPtg = new PtgRef( res, parent, false );
		setWholeRowCol();
		parent_rec = parent;
		this.hashcode = getHashCode();
		this.updateRecord();
	}

	/*
	 * Creates a new PtgArea using an int array as [r,c,r1,c1].
	 * The parent rec is needed
	 * as getting a value goes to the boundsheet to determine
	 * values.  The parent rec *must* be on the same sheet
	 * as the PtgArea referenced!
	 *
	 * relativeRefs = true is excel default
	 */
	public PtgArea( int[] loc, XLSRecord parent, boolean relativeRefs )
	{
		int[] temp = new int[2];
		temp[0] = loc[0];
		temp[1] = loc[1];
		String res = ExcelTools.formatLocation( temp, relativeRefs, relativeRefs );
		firstPtg = new PtgRef( res, parent, false );
		temp[0] = loc[2];
		temp[1] = loc[3];
		res = ExcelTools.formatLocation( temp, relativeRefs, relativeRefs );
		lastPtg = new PtgRef( res, parent, false );
		setWholeRowCol();
		parent_rec = parent;
		this.hashcode = getHashCode();
		this.updateRecord();
	}

	/**
	 * set the wholeRow and/or wholeCol flag for this PtgArea
	 * for ranges such as:
	 * $B:$B and $5:%9
	 */
	public void setWholeRowCol()
	{
		if( (firstPtg.rw <= 1) && lastPtg.wholeCol ) // TODO: inconsistencies in 0-based or 1-based rows
		{
			this.wholeCol = true;
		}
		this.wholeRow = lastPtg.wholeRow;
		if( this.wholeCol )
		{
			useReferenceTracker = false;
		}
	}

	/*
	 * Default constructor
	 */
	public PtgArea()
	{
		record = new byte[9];
		ptgId = 0x25;
		record[0] = 0x25;
	}

	public PtgArea( boolean useReferenceTracker )
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
				ptgId = 0x45;
				break;
			case REFERENCE:
				ptgId = 0x25;
				break;
			case Ptg.ARRAY:
				ptgId = 0x65;
				break;
		}
		record[0] = ptgId;
	}

	public String toString()
	{
		String ret = getString();

		if( getParentRec() != null )
		{
			if( ret.indexOf( "!" ) < 0 )
			{
				try
				{ // Catch WorkSheetNotFoundException to handle Unresolved External refs
					getSheetName();
					if( sheetname != null )
					{
						ret = sheetname + "!" + ret;
					}
				}
				catch( WorkSheetNotFoundException we )
				{
					Logger.logErr( we );
				}
			}
		}
		return ret;
	}

	@Override
	public void setParentRec( XLSRecord rec )
	{
		super.setParentRec( rec );
		// 20080221 KSC: just set parent_rec super.setParentRec(rec);
		if( firstPtg != null )
		{
			firstPtg.setParentRec( parent_rec );
		}
		if( lastPtg != null )
		{
			lastPtg.setParentRec( parent_rec );
		}
	}

	/* Set the location of this PtgRef.  This takes a location
	   such as "a14:b15"
	*/
	@Override
	public void setLocation( String address )
	{
		String s[] = ExcelTools.stripSheetNameFromRange( address );
		setLocation( s );
		this.hashcode = getHashCode();
	}

	/**
	 * set Ptg to parsed location
	 *
	 * @param loc String[] sheet1, range, sheet2, exref1, exref2
	 */
	@Override
	public void setLocation( String[] loc )
	{
		locax = null; // cache reset
		if( firstPtg == null )
		{
			this.record = new byte[]{ 0x25, 0, 0, 0, 0, 0, 0, 0, 0 };

			if( this.getParentRec() != null )
			{
				this.populateVals();
			}
		}
		else if( this.useReferenceTracker )
		{
			this.removeFromRefTracker();
		}
		int i = loc[1].indexOf( ":" );
		// handle single cell addresses as:  A1:A1
		if( i == -1 )
		{
			loc[1] = loc[1] + ":" + loc[1];
			i = loc[1].indexOf( ":" );
		}
		String firstloc = loc[1].substring( 0, i );
		String lastloc = loc[1].substring( i + 1 );
		if( loc[0] != null )
		{
			firstloc = loc[0] + "!" + firstloc;
		}
		if( loc[2] != null )
		{
			lastloc = loc[2] + "!" + lastloc;
		}
		if( loc[3] != null )        // 20090325 KSC: store OOXML External References
		{
			firstloc = loc[3] + firstloc;
		}
		if( loc[4] != null )        // 20090325 KSC: store OOXML External References
		{
			lastloc = loc[4] + lastloc;
		}

		// TODO: do we need to remove refs from tracker?
		firstPtg.setParentRec( this.getParentRec() );
		lastPtg.setParentRec( this.getParentRec() );

		firstPtg.setUseReferenceTracker( false );
		lastPtg.setUseReferenceTracker( false );
		firstPtg.setLocation( firstloc );
		lastPtg.setLocation( lastloc );
		setWholeRowCol();
		this.hashcode = getHashCode();
		this.updateRecord();
		if( this.useReferenceTracker )
		{// check of boolean useReferenceTracker
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
	 * returns the location of the ptg as an array of ints.
	 * [0] = firstRow
	 * [1] = firstCol
	 * [2] = lastRow
	 * [3] = lastCol
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
	 * Set the location of this PtgArea.  This takes a location
	 * such as {1,2,3,4}
	 */
	@Override
	public void setLocation( int[] rowcol )
	{
		locax = null; // cache reset
		if( firstPtg == null )
		{
			//this.record = new byte[] {0x25, 0, 0, 0, 0, 0, 0, 0, 0}; -- don't as can be called from PtgArea3d			
			if( this.getParentRec() != null )
			{
				this.populateVals();
			}
		}
		else if( this.useReferenceTracker )
		{
			this.removeFromRefTracker();
		}

		// TODO: do we need to remove refs from tracker?
		firstPtg.setParentRec( this.getParentRec() );
		firstPtg.setSheetName( sheetname );
		lastPtg.setParentRec( this.getParentRec() );
		lastPtg.setSheetName( sheetname );

		firstPtg.setUseReferenceTracker( false );
		lastPtg.setUseReferenceTracker( false );
		firstPtg.setLocation( rowcol );
		int[] rc = new int[2];
		rc[0] = rowcol[2];
		rc[1] = rowcol[3];
		lastPtg.setLocation( rc );

		this.hashcode = getHashCode();
		this.updateRecord();
		if( this.useReferenceTracker )    // check of boolean useReferenceTracker
		{
			this.addToRefTracker();
		}
	}

	/*
		Returns the location of the Ptg as a string
	*/
	@Override
	public String getLocation()
	{
		String lc = getLocationHelper();
		locax = lc;
		return lc;
	}

	private String getLocationHelper()
	{
		//String loc= null;
		if( (firstPtg == null) || (lastPtg == null) )
		{
			this.populateVals();
			if( (firstPtg == null) || (lastPtg == null) ) // we tried!
			{
				throw new AssertionError( "PtgArea.getLocationHelper null ptgs" );
			}
		}
		String s = firstPtg.getLocation();
		String y = lastPtg.getLocation();

		String loc1[] = ExcelTools.stripSheetNameFromRange( s );    // sheet, addr
		String loc2[] = ExcelTools.stripSheetNameFromRange( y );    // sheet, addr

		String sh1 = loc1[0];
		String sh2 = loc2[0];
		String addr1 = loc1[1];
		String addr2 = loc2[1];

		if( !(this instanceof PtgArea3d) )
		{
			//if (addr1.equals(addr2))	// this is proper but makes so many assertions fail, revert for now
			//return addr2;
			return addr1 + ":" + addr2;
		}

		if( sh1 == null )
		{
			sh1 = sheetname;
		}
		if( sh1 == null )    // no sheetname avail
		{
			return addr1 + ":" + addr2;
		}

		// handle OOXML external references
		if( externalLink1 > 0 )
		{
			sh1 = "[" + externalLink1 + "]" + sh1;
		}
		if( (externalLink2 > 0) && (sh2 != null) )
		{
			sh2 = "[" + externalLink2 + "]" + sh2;
		}

		sh1 = qualifySheetname( sh1 );

		// have sheetname
		if( sh1.equals( sh2 ) )
		{ // range is in one sheet
			if( !sh1.equals( "" ) )
			{
				if( !addr1.equals( addr2 ) )
				{
					return sh1 + "!" + addr1 + ":" + addr2;
				}
				else
				{
					return sh1 + "!" + addr1;
				}
			}
			else if( sheetname != null )
			{    // both sheets in sub-ptgs are null
				sh1 = sheetname;
				// 20090325 KSC: handle OOXML external references
				if( externalLink1 > 0 )
				{
					sh1 = "[" + externalLink1 + "]" + sh1;
				}
				sh1 = qualifySheetname( sh1 );
				if( !addr1.equals( addr2 ) )  // 20081215 KSC:
				{
					return sh1 + "!" + addr1 + ":" + addr2;
				}
				return sh1 + "!" + addr1;
			}
		}
		else if( sh2 == null )
		{    // only 1 sheetnaame specified
			if( !addr1.equals( addr2 ) )  // 20081215 KSC:
			{
				return sh1 + "!" + addr1 + ":" + addr2;
			}
			return sh1 + "!" + addr1;
		}
		// otherwise, include both sheets in return string
		sh2 = qualifySheetname( sh2 );
		return sh1 + ":" + sh2 + "!" + addr1 + ":" + addr2;
	}

	/* Updates the record bytes so it can be pulled back out.
   */
	@Override
	public void updateRecord()
	{
		locax = null; // cache reset
		int[] pols = { firstPtg.getLocationPolicy(), lastPtg.getLocationPolicy() };
		byte[] first = firstPtg.getRecord();
		byte[] last = lastPtg.getRecord();
		// the last record has an extra identifier on it.
		byte[] newrecord = new byte[9];
		newrecord[0] = record[0];
		System.arraycopy( first, 1, newrecord, 1, 2 );
		System.arraycopy( last, 1, newrecord, 3, 2 );
		System.arraycopy( first, 3, newrecord, 5, 2 );
		System.arraycopy( last, 3, newrecord, 7, 2 );
		record = newrecord;
//        this.populateVals();
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
		firstPtg.setLocationPolicy( pols[0] );
		lastPtg.setLocationPolicy( pols[1] );
	}

	@Override
	public int getLength()
	{
		return PTG_AREA_LENGTH;
	}

	/*
		returns the sum of all fields within this range
		** this may need to be modified, as we might not always want sum's
		*20080730 KSC: Excel does *NOT* sum these values except in array formulas

		TODO:  Calculate cell values that are result of formula - take care of
		in cell?
		From Excel File Format Documentation:
			Value class tokens will be changed dependent on further conditions. In array type functions and name type
			  functions, or if the forced array class state is set, it is changed to array class. In all other cases (cell type formula
			  without forced array class), value class is retained.
	*/
	@Override
	public Object getValue()
	{
		// 20080214 KSC: underlying cells may have changed ...if(refCell==null)
		refCell = this.getRefCells();
		Object returnval = (double) 0;
		String retstr = null;
		String array = "";
		boolean isArray = (this.parent_rec instanceof Array);
		for( int t = 0; t < refCell.length; t++ )
		{
			BiffRec cel = refCell[t];
			if( cel == null )
			{ // 20090203 KSC
				continue;
			}

			try
			{
				Formula f = (Formula) cel.getFormulaRec();
				if( f != null )
				{
					Object oby = f.calculateFormula();
					String s = String.valueOf( oby );
					try
					{
						//Double d = new Double(s);
						returnval = new Double( s );
					}
					catch( NumberFormatException ex )
					{
						retstr = s;                    // 20090202 KSC: was +=
					}

				}
				else
				{
					returnval = cel.getInternalVal();    //DblVal();	// 20090202 KSC: was +=
				}
			}
			catch( FunctionNotSupportedException e )
			{
				; // keep going???
			}
			catch( Exception e )
			{
				returnval = cel.getInternalVal();    //DblVal();		// 20090202 KSC: was +=
			}
			if( !isArray )  // 20080730 KSC: if not an array, retrieve only 1st referenced cell value
			{
				break;
			}
			//
			if( retstr != null )
			{
				array = array + retstr + ",";
			}
			else
			{
				array = array + returnval + ",";
			}
			retstr = null;
		}
		if( isArray && (array != null) && (array.length() > 1) )
		{   // 20090817 KSC:  [BugTracker 2683]
			array = "{" + array.substring( 0, array.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( array );
			return pa;
		}
		if( retstr != null )
		{
			return retstr;
		}

		return returnval;    //new Double(returnval);
	}

	/*
		Returns all of the cells of this range as PtgRef's.
		This includes empty cells, values, formulas, etc.
		Note the setting of parent-rec requires finding the cell
		the PtgRef refer's to.  If that is null then the PtgRef
		will exist, just with a null value.  This could cause issues when
		programatically populating cells.
	*/
	public Ptg[] getColComponents( int colNum )
	{
		if( colNum < 0 )
		{
			return null;
		}
		String lu = this.toString();
		Object p = parent_rec.getWorkBook().getRefTracker().getVlookups().get( lu );

		if( p != null )
		{
			PtgArea par = (PtgArea) p;
			Ptg[] ret = (Ptg[]) par.getParentRec()
			                       .getWorkBook()
			                       .getRefTracker()
			                       .getLookupColCache()
			                       .get( lu + ":" + colNum );
			if( ret != null )
			{
				return ret;
			}
		}

		PtgRef[] v = null;
		try
		{
//			 TODO: check rc sanity here
			int[] startloc = firstPtg.getRealIntLocation();
			int startcol = colNum; // startloc[0];
			int startrow = startloc[0];
			int[] endloc = lastPtg.getRealIntLocation();
			int endrow = endloc[0];
			// error trap
			if( endrow < startrow )    // can happen if wholerow/wholecol, getMaxRow may be less than startRow
			{
				endrow = startrow;
			}
			int sz = endrow - startrow;
			sz++;
			v = new PtgRef[sz];
			// loop through the cols
			// loop through the rows inside
			int rowholder = startrow;
			int pos = 0;
			String sht = this.toString();
			if( sht.indexOf( "!" ) > -1 )
			{
				sht = sht.substring( 0, sht.indexOf( "!" ) );
			}
			for(; rowholder <= endrow; rowholder++ )
			{
				String displaycol = ExcelTools.getAlphaVal( startcol );
				int displayrow = rowholder + 1;
				String loc = sht + "!" + displaycol + displayrow;

				PtgRef pref = new PtgRef( loc, parent_rec, this.useReferenceTracker );

				v[pos++] = pref;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Getting column range in PtgArea failed.", e );
		}

		// cache
		parent_rec.getWorkBook().getRefTracker().getVlookups().put( this.toString(), this );
		parent_rec.getWorkBook().getRefTracker().getLookupColCache().put( lu + ":" + colNum, v );
		return v;
	}

	/**
	 * return the ptg components for a certain column within a ptgArea()
	 *
	 * @param rowNum
	 * @return all Ptg's within colNum
	 */
	public Ptg[] getRowComponents( int rowNum )
	{
		FastAddVector v = new FastAddVector();
		Ptg[] allComponents = this.getComponents();
		for( int i = 0; i < allComponents.length; i++ )
		{
			PtgRef p = (PtgRef) allComponents[i];
//			 TODO: check rc sanity here
			int[] x = p.getRealIntLocation();
			if( x[0] == rowNum )
			{
				v.add( p );
			}
		}
		PtgRef[] pref = new PtgRef[v.size()];
		v.toArray( pref );
		return pref;
	}

	/**
	 * @return
	 */
	public PtgRef getFirstPtg()
	{
		return firstPtg;
	}

	/**
	 * @return
	 */
	public PtgRef getLastPtg()
	{
		return lastPtg;
	}

	/**
	 * @param ref
	 */
	public void setFirstPtg( PtgRef ref )
	{
		locax = null; // cache reset
		firstPtg = ref;
	}

	/**
	 * @param ref
	 */
	public void setLastPtg( PtgRef ref )
	{
		locax = null; // cache reset
		lastPtg = ref;
	}

	/**
	 * @return Returns the refCell.
	 */
	@Override
	public BiffRec[] getRefCells()
	{
		double returnval = 0;
		try
		{
			Boundsheet bs = null; // this.parent_rec.getWorkBook().getWorkSheetByName(this.getSheetName());
			getSheetName();
			// handle misc sheets
			if( sheetname != null )
			{
				try
				{
					bs = this.parent_rec.getWorkBook().getWorkSheetByName( sheetname );
				}
				catch( Exception ex )
				{ // guard against NPEs
					bs = parent_rec.getSheet();
				}
			}
			else
			{
				bs = parent_rec.getSheet();
				sheetname = bs.getSheetName();    // 20080212 KSC
			}

//			 TODO: check rc sanity here
			int[] startloc = firstPtg.getIntLocation();
			int startcol = startloc[1];
			int startrow = startloc[0];
			int[] endloc = lastPtg.getRealIntLocation();
			int endcol = endloc[1];
			int endrow = endloc[0];
			// loop through the cols
			int numcols = endcol - startcol;
			if( numcols < 0 )
			{
				numcols = startcol - endcol;    // 20090521 KSC: may have range switched so that firstPtg>lastPtg (example in named ranges in tcr_formatted_2007.xlsm)
			}
			numcols++;
			int numrows = endrow - startrow;
			if( numrows < 0 )
			{
				numrows = startrow - endrow;    // 20090521 KSC: may have range switched so that firstPtg>lastPtg (example in named ranges in tcr_formatted_2007.xlsm)
			}
			numrows++;
			int totcell = numcols * numrows;
			if( totcell == 0 )
			{
				totcell++;
			}
			if( totcell < 0 )
			{
				Logger.logErr( "PtgArea.getRefCells.  Error in Ptg locations: " + firstPtg.toString() + ":" + lastPtg.toString() );
				totcell = 0;
			}
			refCell = new BiffRec[totcell];
			int rowctr = 0;
			// 20090521 KSC: try to handle both cases i.e. ranges such that first<last or first>last
			if( startcol < endcol )
			{
				endcol++;
			}
			else
			{
				endcol--;
			}
			if( startrow < endrow )
			{
				endrow++;
			}
			else
			{
				endrow--;
			}
			while( startcol != endcol )
			{
				int rowpos = startrow;
				while( rowpos != endrow )
				{
					Row r = bs.getRowByNumber( rowpos );
					if( r != null )
					{
						refCell[rowctr] = (BiffRec) r.getCell( (short) (startcol) );
					}
					rowctr++;
					if( rowpos < endrow )
					{
						rowpos++;
					}
					else
					{
						rowpos--;
					}
				}
				if( startcol < endcol )
				{
					startcol++;
				}
				else
				{
					startcol--;
				}
			}

            /*
            for (;startcol<=endcol;startcol++){ // loop the cols
                int rowpos = startrow;
                for (;rowpos<=endrow;rowpos++){ // loop through the rows
                    Row r = bs.getRowByNumber(rowpos);
                    if (r!=null)
                        refCell[rowctr] = (BiffRec)r.getCell((short)(startcol));
                    rowctr++;
                }
            }
            */
		}
		catch( Exception ex )
		{
			Logger.logErr( "PtgArea.getRefCells failed.", ex );
		}
		return refCell;
	}

	@Override
	protected long getHashCode()
	{
		return lastPtg.hashcode + ((firstPtg.hashcode) * ((long) MAXCOLS + ((long) MAXROWS * MAXCOLS)));
	}
}