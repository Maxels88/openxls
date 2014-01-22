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

import com.extentech.formats.XLS.Shrfmla;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;

/**
 * ptgArea is a reference to an area (rectangle) of cells.
 * Essentially it is a collection of two ptgRef's, so it will be
 * treated that way in the code...
 * <p/>
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     	2       The First row of the reference
 * 2           rwLast     		2       The Last row of the reference
 * 4           grbitColFirst   2       (see following table)
 * 6           grbitColLast    2       (see following table)
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
public class PtgAreaN extends PtgArea
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8433468704529379504L;

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

	PtgRefN firstPtgN;
	PtgRefN lastPtgN;
	private PtgArea parea = null;

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

/*
 *  0           rwFirst     2       The First row of the reference 
    2           rwLast     2       The Last row of the reference 
    4           grbitColFirst    2       (see following table)
    6           grbitColLast    2       (see following table)

    Only the low-order 14 bits specify the Col, the other bits specify
    relative vs absolute for both the col or the row.

    Bits        Mask        Name    Contents
    -----------------------------------------------------
    15          8000h       fRwRel  =1 if row offset relative, 
                                    =0 if otherwise
    14          4000h       fColRel =1 if row offset relative,
                                    =0 if otherwise
    13-0        3FFFh       col     Ordinal column offset or number
        
 */
		firstPtgN = new PtgRefN( false );
		firstPtgN.setParentRec( parent_rec );
		firstPtgN.init( temp1 );
		lastPtgN = new PtgRefN( false );
		lastPtgN.setParentRec( parent_rec );
		lastPtgN.init( temp2 );
		if( (parent_rec != null) && (parent_rec instanceof Shrfmla) )
		{
			// 20060301 KSC: init sets formula row to 1st row for a shared formula; adjust here
			lastPtgN.setFormulaRow( ((Shrfmla) parent_rec).getLastRow() );
			lastPtgN.setFormulaCol( ((Shrfmla) parent_rec).getLastCol() );
		}
//        if (this.useReferenceTracker)	
//        	this.addToRefTracker();
	}

	/**
	 * returns the row/col ints for the ref
	 *
	 * @return
	 */
	@Override
	public int[] getRowCol()
	{
		if( firstPtgN == null )
		{
			int[] rc1 = firstPtgN.getRowCol();
			int[] ret = { rc1[0], rc1[1], rc1[0], rc1[1] };
			return ret;
		}
		int[] rc1 = firstPtgN.getRowCol();
		int[] rc2 = lastPtgN.getRowCol();
		int[] ret = { rc1[0], rc1[1], rc2[0], rc2[1] };
		return ret;
	}

	/**
	 * returns the uncoverted, actual row col
	 *
	 * @return
	 */
	public int[] getRealRowCol()
	{
		return new int[]{ firstPtgN.rw, firstPtgN.col, lastPtgN.rw, lastPtgN.col };
	}

	public PtgArea convertToPtgArea( com.extentech.formats.XLS.XLSRecord r )
	{
		PtgRef p1 = firstPtgN.convertToPtgRef( r );
		PtgRef p2 = lastPtgN.convertToPtgRef( r );
		PtgArea par = new PtgArea( p1, p2, r );
		return par;
	}

	/**
	 * update record bytes
	 */
	// 20060223 KSC
	@Override
	public void updateRecord()
	{
		byte[] first = firstPtgN.getRecord();
		byte[] last = lastPtgN.getRecord();
		// the last record has an extra identifier on it.
		byte[] newrecord = new byte[9];
		newrecord[0] = record[0];
		System.arraycopy( first, 1, newrecord, 1, 2 );
		System.arraycopy( last, 1, newrecord, 3, 2 );
		System.arraycopy( first, 3, newrecord, 5, 2 );
		System.arraycopy( last, 3, newrecord, 7, 2 );
		record = newrecord;
	}

	/*
	Returns the location of the Ptg as a string
*/
	@Override
	public String getLocation()
	{
		if( (firstPtgN == null) || (lastPtgN == null) )
		{
			this.populateVals();
			if( (firstPtgN == null) || (lastPtgN == null) ) // we tried
			{
				throw new AssertionError( "PtgAreaN.getLocation null ptgs" );
			}
		}
		String s = firstPtgN.getLocation();
		String y = lastPtgN.getLocation();

		return s + ":" + y;
	}

	/**
	 * returns an array of the first and last addresses in the PtgAreaN
	 */
	// 20060223: KSC: customize from ptgArea
	@Override
	public int[] getIntLocation()
	{
		int[] returning = new int[4];
		try
		{
			int[] first = firstPtgN.getIntLocation();
			int[] last = lastPtgN.getIntLocation();
			System.arraycopy( first, 0, returning, 0, 2 );
			System.arraycopy( last, 0, returning, 2, 2 );
		}
		catch( Exception e )
		{
			;
		}
		return returning;
	}

	/**
	 * @return lastPtgN
	 */
	public PtgRefN getLastPtgN()
	{
		return lastPtgN;
	}

	/**
	 * @return firstPtgN
	 */
	public PtgRefN getFirstPtgN()
	{
		return firstPtgN;
	}

	/**
	 * custom RefTracker usage:  uses entire range covered by all shared formulas
	 */
	public PtgArea getArea()
	{
		Shrfmla sh = (Shrfmla) this.getParentRec();
		int[] i = new int[4];
		if( fRwRel )
		{
			i[0] = sh.getFirstRow() + firstPtgN.rw;
		}
		else
		{
			i[0] = firstPtgN.rw;
		}
		if( fColRel )
		{
			i[1] = sh.getFirstCol() + firstPtgN.col;
		}
		else
		{
			i[1] = firstPtgN.col;
		}
		if( fRwRel )
		{
			i[2] = sh.getLastRow() + lastPtgN.rw;
		}
		else
		{
			i[2] = lastPtgN.rw;
		}
		if( fColRel )
		{
			i[3] = sh.getLastCol() + lastPtgN.col;
		}
		else
		{
			i[3] = lastPtgN.col;
		}

		if( (i[1] >= MAXCOLS_BIFF8) && !this.parent_rec.getWorkBook().getIsExcel2007() )    // TODO: determine if this is an OK maxcol (Excel 2007)
		{
			i[1] -= MAXCOLS_BIFF8;
		}
		if( (i[3] >= MAXCOLS_BIFF8) && !this.parent_rec.getWorkBook().getIsExcel2007() )    // TODO: determine if this is an OK maxcol (Excel 2007)
		{
			i[3] -= MAXCOLS_BIFF8;
		}

		PtgArea parea = new PtgArea( i, (XLSRecord) sh, true );
		return parea;
	}

	/**
	 * add "true" area to reference tracker i.e. entire range referenced by all shared formula members
	 */
	@Override
	public void addToRefTracker()
	{
		int iParent = this.getParentRec().getOpcode();
		if( iParent == XLSConstants.SHRFMLA )
		{
			// KSC: TESTING - local ptgarea gets finalized and messes up ref. tracker on multiple usages without close
			//getArea();
			//parea.addToRefTracker();
			PtgArea parea = this.getArea();    // is finalized if local var --- but take out ptgarea finalize for now
			parea.addToRefTracker();
		}
	}

	/**
	 * remove "true" area from reference tracker i.e. entire range referenced by all shared formula members
	 */
	@Override
	public void removeFromRefTracker()
	{
		int iParent = this.getParentRec().getOpcode();
		if( iParent == XLSConstants.SHRFMLA )
		{
			PtgArea parea = this.getArea();
			parea.removeFromRefTracker();
		}
		//if (parea!=null) {
		//	parea.removeFromRefTracker();
		//parea.close();
		//}
		//parea= null;
	}

	@Override
	public void close()
	{
		removeFromRefTracker();
		if( parea != null )
		{
			parea.close();
		}
		parea = null;
		if( firstPtgN != null )
		{
			firstPtgN.close();
		}
		firstPtgN = null;
		if( lastPtgN != null )
		{
			lastPtgN.close();
		}
		lastPtgN = null;
	}
}	
