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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Dbcell;
import com.extentech.formats.XLS.Row;
import com.extentech.formats.XLS.Shrfmla;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.Logger;

/**
 * PtgRefN is a modified PtgRef that is for shared formulas.
 * Put here by M$ to make us miserable,
 * <p/>
 * it would have made much more sense to just use a PtgRef.
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rw          2       The row of the reference (so says the docs, but it is the row I think
 * 2           grbitCol    2       (see following table)
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
 * <p/>
 * This token contains the relative reference to a cell in the same sheet.
 * It stores relative components as signed offsets and is used in shared formulas, conditional formatting, and data validity.
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */
public class PtgRefN extends PtgRef
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2652944516984815274L;
	/* private int formulaRow;
	private int formulaCol;*/
	private int realRow;
	private int realCol;
	short column;
	private PtgArea parea = null;

	@Override
	public boolean getIsReference()
	{
		return true;
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		populateVals();
		hashcode = getHashCode();        // different from PtgRef calc
	}

	public PtgRefN( boolean useReference )
	{
		this.setUseReferenceTracker( useReference );
	}

	/**
	 * Returns the location of the Ptg as a string (ie c4)
	 * <p/>
	 * TODO: look into this possible bug
	 * <p/>
	 * There is a problem here as the location will always be relative and incorrect.
	 * this is deprecated and should be calling convertToPtgRef
	 */
	@Override
	public String getLocation()
	{
		//if (!populated){throw new FormulaNotFoundException("Cannot set location, no Formula Present");}
		realRow = rw;
		realCol = col;
		if( fRwRel )
		{  // the row is a relative location
			realRow += (short) formulaRow;
		}
		if( fColRel )
		{  // the column is a relative location
			realCol = (short) formulaCol;// - colNew;
			if( realCol >= MAXCOLS )
			{
				realCol -= MAXCOLS;
			}
		}
		String s = ExcelTools.getAlphaVal( realCol );
		String y = String.valueOf( realRow + 1 );

		return (fColRel ? "" : "$") + s + (fRwRel ? "" : "$") + y;
	}

	/* Set the location of this PtgRef.  This takes a location
	   such as {1,2}
	*/
	@Override
	public void setLocation( int[] rowcol )
	{
		if( useReferenceTracker )
		{
			this.removeFromRefTracker();
		}
		if( record != null )
		{    // 20090217 KSC: had some errors here, redid
			if( fRwRel )
			{
				formulaRow = rowcol[0];
			}
			else
			{
				rw = rowcol[0];
			}
			if( fColRel )
			{
				formulaCol = rowcol[1];
			}
			else
			{
				col = rowcol[1];
			}
			this.updateRecord();
			init( record );
		}
		else
		{
			Logger.logWarn( "PtgRefN.setLocation() failed: NO record data: " + rowcol.toString() );
		}
		hashcode = getHashCode();
		if( useReferenceTracker )
		{
			this.addToRefTracker();
		}

	}

	/**
	 * returns the row/col ints for the ref
	 * adjusted for the host cell
	 *
	 * @return
	 */
	@Override
	public int[] getRowCol()
	{
		realRow = rw;
		realCol = col;
		if( fRwRel )
		{  // the row is a relative location
			realRow += (short) formulaRow;
		}
		if( fColRel )
		{  // the column is a relative location
			realCol = (short) formulaCol;// - colNew;
			if( realCol >= MAXCOLS )
			{
				realCol -= MAXCOLS;
			}
		}
		int[] ret = { realRow, realCol };
		return ret;
	}

	/**
	 * this
	 *
	 * @return
	 */
	public int[] getRealRowCol()
	{
		return new int[]{ rw, col };
	}

	/* Set the location of this PtgRef.  This takes a location
	   such as "a14"

	   TODO: check why this is overridden / reversed 12/02 -jm
	*/
	@Override
	public void setLocation( String address )
	{
		if( record != null )
		{
			// 20080215 KSC: replace address stripping
			String s[] = ExcelTools.stripSheetNameFromRange( address );
			address = s[1];    //stripped of sheet name, if any ...

			int[] res = ExcelTools.getRowColFromString( address );
			// 20060301 KSC: Keep relativity
			if( fRwRel )
			{
				rw += formulaRow - res[0];
				if( rw < 0 ) // handle row shifting issues
				{
					rw = 0;
				}
				formulaRow = res[0];
			}
			else
			{
				rw = res[0];
			}
			if( fColRel )
			{
				col += formulaCol - res[1];
				formulaCol = res[1];
			}
			else
			{
				col = res[1];
			}

			updateRecord();
			init( record );
			// 20090325 KSC: trap OOXML external reference link, if any
			if( s[3] != null )
			{
				externalLink1 = Integer.valueOf( s[3].substring( 1, s[3].length() - 1 ) );
			}
			if( s[4] != null )
			{
				externalLink2 = Integer.valueOf( s[4].substring( 1, s[4].length() - 1 ) );
			}

		}
		else
		{
			Logger.logWarn( "PtgRefN.setLocation() failed: NO record data: " + address.toString() );
		}
	}

	/**
	 * Convert this PtgRefN to a PtgRef based on the offsets included in the PtgExp &
	 * if this uses relative or absolute offsets
	 *
	 * @param pxp
	 * @return
	 */
	public PtgRef convertToPtgRef( XLSRecord r/*PtgExp pxp*/ )
	{
		//XLSRecord r = (XLSRecord)pxp.getParentRec();
		int[] i = new int[2];
		if( fRwRel )
		{
			i[0] = r.getRowNumber() + rw;
		}
		else
		{
			i[0] = rw;
		}
		if( fColRel )
		{
			i[1] = r.getColNumber() + col;
		}
		else
		{
			i[1] = col;
		}

		if( (i[1] >= MAXCOLS_BIFF8) && !r.getWorkBook().getIsExcel2007() )    // TODO: determine if this is an OK maxcol (Excel 2007)
		{
			i[1] -= MAXCOLS_BIFF8;
		}

		PtgRef prf = new PtgRef( i, r, false );
//	  	String s = ExcelTools.formatLocation(i, fRwRel, fColRel);
//	  	PtgRef prf = new PtgRef(s, r /*pxp.getParentRec()*/, false); //false);	 	

		return prf;
	}

	/**
	 * /*
	 * (try to) return int[] array containing the row/column
	 * referenced by this PtgRefN.
	 *
	 * @returns int[] row/col absolute (non-offset) location
	 * @see com.extentech.formats.XLS.formulas.PtgRef#getIntLocation()
	 */
	@Override
	public int[] getIntLocation()
	{

		int rowNew = rw;
		int colNew = col;
		if( fRwRel )
		{  // the row is a relative location
			rowNew += formulaRow;
		}
		if( fColRel )
		{  // the column is a relative location
			colNew += formulaCol;
		}
		if( colNew >= MAXCOLS )
		{
			colNew -= MAXCOLS;    // 20070205 KSC: Added	20080102 KSC: added =
		}

		int[] returning = new int[2];
		returning[0] = rowNew;
		returning[1] = colNew;
		return returning;
	}

	/**
	 * set formula row
	 *
	 * @param r int new row
	 */
	// 20060301 KSC: access to formula row/col
	public void setFormulaRow( int r )
	{
		formulaRow = r;
	}

	/**
	 * set formula col
	 *
	 * @param c int new col
	 */
	// 20060301 KSC: access to formula row/col
	public void setFormulaCol( int c )
	{
		formulaCol = c;
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
			i[0] = sh.getFirstRow() + rw;
		}
		else
		{
			i[0] = rw;
		}
		if( fColRel )
		{
			i[1] = sh.getFirstCol() + col;
		}
		else
		{
			i[1] = col;
		}
		if( fRwRel )
		{
			i[2] = sh.getLastRow() + rw;
		}
		else
		{
			i[2] = rw;
		}
		if( fColRel )
		{
			i[3] = sh.getLastCol() + col;
		}
		else
		{
			i[3] = col;
		}

		if( (i[1] >= MAXCOLS_BIFF8) && !this.parent_rec.getWorkBook().getIsExcel2007() )    // TODO: determine if this is an OK maxcol (Excel 2007)
		{
			i[1] -= MAXCOLS_BIFF8;
		}
		if( (i[3] >= MAXCOLS_BIFF8) && !this.parent_rec.getWorkBook().getIsExcel2007() )    // TODO: determine if this is an OK maxcol (Excel 2007)
		{
			i[3] -= MAXCOLS_BIFF8;
		}

		PtgArea parea = new PtgArea( i, sh, true );
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
//				getArea();
//				parea.addToRefTracker();
			PtgArea parea = getArea(); // otherwise is finalized if local var --- but take out ptgarea finalize for now
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
			PtgArea parea = getArea(); // otherwise is finalized if local var --- but take out ptgarea finalize for now
			if( parea != null )
			{
				parea.removeFromRefTracker();
//			  		parea.close();
			}
			//	parea= null;
		}
	}

	@Override
	public void close()
	{
		//removeFromRefTracker();
		if( parea != null )
		{
			parea.close();
		}
		parea = null;
	}
}