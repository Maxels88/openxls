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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.formulas.FormulaParser;
import com.extentech.formats.XLS.formulas.GenericPtg;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgAreaN;
import com.extentech.formats.XLS.formulas.PtgExp;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.formats.XLS.formulas.PtgRefN;
import com.extentech.toolkit.ByteTools;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.SortedSet;
import java.util.Stack;
import java.util.TreeSet;

/**
 * SHRFMLA is a file optimization that stores many identical formulas once.
 * The SHRFMLA record appears immediately after the first FORMULA record in the
 * group. Each FORMULA record in the group will have fShrFmla set and will
 * contain only a PtgExp pointing to the cell containing the SHRFMLA record.
 * <p/>
 * Occasionally Excel writes a FORMULA record with fShrFmla set whose
 * expression is an instantiation of the relevant shared formula instead of the
 * usual PtgExp. We currently handle these by clearing fShrFmla. At some point
 * in the future we will attempt to detect which shared formula is applied and
 * make the cell a member of its group.
 * <p/>
 * ExtenXLS does not currently create shared formulas or add cells to existing
 * ones. Existing shared formulas will be preserved in output. Removal of
 * formulas from the group is supported, including the cell currently hosting
 * the SHRFMLA record. Shared formula member cells will be reference-tracked
 * and recalculated properly.
 * <pre>
 * OFFSET NAME        SIZE CONTENTS
 * 0      rwFirst     2    First Row
 * 2      rwLast      2    Last Row
 * 4      colFirst    1    First Column
 * 5      colLast     1    Last Column
 * 6      (reserved)  2    pass through, zero for new
 * 8      cce         2    Length of the parsed expression
 * 10     rgce        cce  Parsed Expression
 * </pre>
 */
public final class Shrfmla extends XLSRecord
{
	private static final long serialVersionUID = -6147947203791941819L;

	private int rwFirst;
	private int rwLast;
	private int colFirst;
	private int colLast;
	private Stack expression;
	private Ptg[] ptgcache;
	private Formula host;

	/**
	 * The set of Formula records referring to this shared formula.
	 */
	private SortedSet members = new TreeSet( new CellAddressComparator() );

	/**
	 * Whether this formula contains an indirect reference.
	 */
	boolean containsIndirectFunction = false;

	public int getFirstRow()
	{
		return (int) rwFirst;
	}

	public int getLastRow()
	{
		return (int) rwLast;
	}

	public int getFirstCol()
	{
		return (int) colFirst;
	}

	public int getLastCol()
	{
		return (int) colLast;
	}

	/**
	 * update location upon a shift (row insertion or deletion)-- ensure member formulas cache are cleared
	 *
	 * @param shiftamount
	 */
	public void updateLocation( int shiftamount, PtgRef pr )
	{
		// remove original reference
		if( ptgcache.length > 1 )
		{
			// for shrfmlas which contain multiple ptgs, ensure formulas get shifted only 1x!
			if( ptgcache[0] instanceof PtgRefN )
			{
				if( pr.hashcode != ((PtgRefN) ptgcache[0]).getArea().hashcode )
				{
					return;    // it's already been shifted
				}
			}
			else
			{
				if( pr.hashcode != ((PtgAreaN) ptgcache[0]).getArea().hashcode )
				{
					return;    // it's already been shifted
				}
			}
		}
		for( int i = 0; i < ptgcache.length; i++ )
		{
			if( ptgcache[i] instanceof PtgRefN )
			{
				((PtgRefN) ptgcache[i]).removeFromRefTracker();
			}
			else
			{
				((PtgAreaN) ptgcache[i]).removeFromRefTracker();
			}
		}
		Iterator<Formula> ii = members.iterator();
		while( ii.hasNext() )
		{
			Formula f = ii.next();
			f.clearCachedValue();
			// also update PtgExp
			PtgExp pointer = (PtgExp) f.getExpression().get( 0 );
			pointer.setRowFirst( pointer.getRwFirst() + shiftamount );
		}
		setFirstRow( rwFirst + shiftamount );
		setLastRow( rwLast + shiftamount );
		for( int i = 0; i < ptgcache.length; i++ )
		{
			if( ptgcache[i] instanceof PtgRefN )
			{
				((PtgRefN) ptgcache[i]).addToRefTracker();
			}
			else
			{
				((PtgAreaN) ptgcache[i]).addToRefTracker();
			}
		}
	}

	public void setFirstRow( int row )
	{
		rwFirst = row;
		rw = rwFirst;
	}

	public void setLastRow( int row )
	{
		rwLast = row;
	}

	@Override
	public void init()
	{
		super.init();
		rwFirst = ByteTools.readUnsignedShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		rwLast = ByteTools.readUnsignedShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		colFirst = this.getByteAt( 4 );
		colLast = this.getByteAt( 5 );
		short cce = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		byte[] rgce = this.getBytesAt( 10, cce );
		if( this.getSheet() == null )
		{
			this.setSheet( this.wkbook.getLastbound() );
		}
		rw = rwFirst;

		try
		{
			wkbook.lastFormula.initSharedFormula( this );
			this.setHostCell( wkbook.lastFormula );
		}
		catch( Exception e )
		{
			;
		}

		expression = ExpressionParser.parseExpression( rgce, this );
		if( containsIndirectFunction )
		{
			if( host != null )
			{
				host.registerIndirectFunction();
			}
		}
		// Cache Relative Ptgs for quickness of access
		ArrayList<Ptg> ptgs = new ArrayList();
		for( int idx = 0; idx < expression.size(); idx++ )
		{
			Ptg ptg = (Ptg) expression.get( idx );
			if( ptg instanceof PtgRefN )
			{
				ptgs.add( ptg );
			}
			else if( ptg instanceof PtgAreaN )
			{
				ptgs.add( ptg );
			}
		}
		ptgcache = new Ptg[ptgs.size()];
		ptgs.toArray( ptgcache );
	}

	@Override
	public void preStream()
	{
		super.preStream();

		byte[] data = getData();

		System.arraycopy( ByteTools.shortToLEBytes( (short) rwFirst ), 0, data, 0, 2 );

		System.arraycopy( ByteTools.shortToLEBytes( (short) rwLast ), 0, data, 2, 2 );

		data[4] = (byte) colFirst;
		data[5] = (byte) colLast;

		data[7] = (byte) members.size();

		setData( data );
	}

	boolean isInRange( String s )
	{
		return ExcelTools.isInRange( s, rwFirst, rwLast, colFirst, colLast );
	}

	/**
	 * Converts an expression stack that uses relative PTGs to
	 * a standard stack for calculation
	 */
	// NOTE: now these ptgs are not reference-tracked; see ExpressionParser and PtgRefN,PtgAreaN for reference-tracking these entities
	public static Stack convertStack( Stack in, Formula f )
	{
		Stack out = new Stack();
		for( int idx = 0; idx < in.size(); idx++ )
		{
			Ptg ptg = (Ptg) in.get( idx );
			// convert the Ptg if necessary, otherwise clone it
			if( ptg instanceof PtgRefN )
			{
				ptg = ((PtgRefN) ptg).convertToPtgRef( f );
			}
			else if( ptg instanceof PtgAreaN )
			{
				ptg = ((PtgAreaN) ptg).convertToPtgArea( f );
			}
			else
			{
				ptg = (Ptg) ptg.clone();
				ptg.setParentRec( f );
			}

			out.add( ptg );
		}
		return out;
	}

	public String toString()
	{
		return "Shared Formula [" + getCellRange() + "] " + FormulaParser.getExpressionString( expression );
	}

	/**
	 * Gets the area where references to this shared formula may occur.
	 *
	 * @return an A1-style range string
	 */
	public String getCellRange()
	{
		return ExcelTools.formatRange( new int[]{ colFirst, rwFirst, colLast, rwLast } );
	}

	/**
	 * Sets the formula record in which this record resides.
	 */
	public void setHostCell( Formula newHost )
	{
		if( host != null )
		{
			host.removeInternalRecord( this );
		}

		host = newHost;
		host.addInternalRecord( this );
		rw = host.getRowNumber();
		col = host.getColNumber();
	}

	/**
	 * Gets the formula record in which this record resides.
	 */
	public Formula getHostCell()
	{
		return host;
	}

	/**
	 * Gets a {@link PtgExp} pointer to this <code>ShrFmla</code>.
	 * The returned <code>PtgExp</code> points to this record at its current
	 * location. If the host cell or its address changes the pointer will
	 * become invalid.
	 */
	public PtgExp getPointer()
	{
		PtgExp pointer = new PtgExp();
		pointer.setParentRec( host );
		pointer.init( host.getRowNumber(), host.getColNumber() );
		return pointer;
	}

	/**
	 * Adds a member formula to this shared formula.
	 *
	 * @throws IndexOutOfBoundsException if there are already 255 members
	 */
	public void addMember( Formula member )
	{
		if( members.size() >= 255 )
		{
			throw new IndexOutOfBoundsException( "shared formula already has 255 members" );
		}

		members.add( member );
		if( members.size() == 1 )    // KSC: ADDED
		{
			setHostCell( member );
		}

		// Only do range/host manipulation if we're not parsing
		if( !getWorkBook().getFactory().iscompleted() )
		{
			return;
		}

		// If the newly added member is the first one it must become the host
//    	if (members.first() == member) setHostCell( member );  KSC: replaced with above

		int row = member.getRowNumber();
		int col = member.getColNumber();

		if( row < rwFirst )
		{
			rwFirst = row;
		}
		if( row > rwLast )
		{
			rwLast = row;
		}
		if( col < colFirst )
		{
			colFirst = col;
		}
		if( col > colLast )
		{
			colLast = col;
		}
	}

	public void removeMember( Formula member )
	{
		members.remove( member );

		// If we've just removed the last member, don't bother adjusting
		// because we're about to be deleted.
		if( members.size() == 0 )
		{
			return;
		}

    	/* Ideally we would shrink the range to the smallest possible value,
    	 * but it's not actually required. Finding the column bounds is somewhat
    	 * expensive as it requires iterating the member list. Therefore we
    	 * only update the row components of the range.
    	 */

		if( member.getRowNumber() == rwLast )
		{
			rwLast = (short) ((Formula) members.last()).getRowNumber();
		}

		// If we're removing the host cell, choose another one
		if( member == host )
		{
			setHostCell( (Formula) members.first() );
			rwFirst = host.getRowNumber();
		}
	}

	/**
	 * return all the formulas that use this Shrfmla
	 *
	 * @return
	 */
	public SortedSet getMembers()
	{
		return members;
	}

	public Stack getStack()
	{
		return expression;
	}

	/**
	 * convert expression using dimensions of specific member formula
	 *
	 * @param parent
	 * @return
	 */
	public Stack instantiate( Formula parent )
	{
		return convertStack( expression, parent );
	}

	/**
	 * Set if the formula contains Indirect()
	 *
	 * @param containsIndirectFunction The containsIndirectFunction to set.
	 */
	protected void setContainsIndirectFunction( boolean containsIndirectFunction )
	{
		this.containsIndirectFunction = containsIndirectFunction;
	}

	/**
	 * Adds an indirect function to the list of functions to be evaluated post load
	 *
	 * *
	 * NOT IMPLEMENTED YET
	 protected void registerIndirectFunction() {
	 this.getWorkBook().addIndirectFormula(this);
	 }*/

	/**
	 * determine which formula in set of shared formula members is affected by cell br
	 *
	 * @param br cell
	 * @return formula which references cell
	 */
	protected Formula getAffected( BiffRec br )
	{
		int[] rc = new int[2];
		rc[0] = br.getRowNumber();
		rc[1] = br.getColNumber();
		Iterator<Formula> ii = members.iterator();
		boolean isExcel2007 = this.getWorkBook().getIsExcel2007();
		while( ii.hasNext() )
		{
			Formula f = ii.next();
			int[] frc = new int[2];
			frc[0] = f.getRowNumber();
			frc[1] = f.getColNumber();
			for( int i = 0; i < ptgcache.length; i++ )
			{
				if( ptgcache[i] instanceof PtgRefN )
				{
					int[] refrc = ((PtgRefN) ptgcache[i]).getRealRowCol();
					if( ((refrc[0] + frc[0]) == rc[0]) && ((adjustCol( refrc[1] + frc[1], isExcel2007 )) == rc[1]) )
					{
						return f;
					}
				}
				else
				{    // it's a PtgAreaN
					int[] refrc = ((PtgAreaN) ptgcache[i]).getRealRowCol();
					refrc[0] += frc[0];
					refrc[2] += frc[1];
					refrc[1] = adjustCol( refrc[1] + frc[0], isExcel2007 );
					refrc[3] = adjustCol( refrc[3] + frc[0], isExcel2007 );
					if( (refrc[0] <= rc[0]) &&
							(refrc[1] <= rc[1]) &&
							(refrc[2] >= rc[0]) &&
							(refrc[3] >= rc[1]) )
					{
						return f;
					}
				}
			}
		}
		return null;
	}

	/**
	 * basic algorithm to adjust column dimensions when > MAXCOLS
	 *
	 * @param c
	 * @param isExcel2007
	 * @return
	 */
	private int adjustCol( int c, boolean isExcel2007 )
	{
		if( (c >= MAXCOLS_BIFF8) && !isExcel2007 )
		{
			c -= MAXCOLS_BIFF8;
		}
		return c;
	}

	@Override
	public void close()
	{
		if( members != null )
		{
			members.clear();
		}
		members = null;
		if( expression != null )
		{
			while( !expression.isEmpty() )
			{
				GenericPtg p = (GenericPtg) expression.pop();
				if( p instanceof PtgRef )
				{
					((PtgRef) p).close();
				}
				else
				{
					p.close();
				}
				p = null;
			}
		}
		ptgcache = null;
		host = null;
		super.close();
	}

	@Override
	protected void finalize()
	{
		this.close();
	}
}