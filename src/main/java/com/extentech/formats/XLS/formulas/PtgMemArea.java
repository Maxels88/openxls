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

import com.extentech.formats.XLS.ExpressionParser;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;
import java.util.Stack;

/**
 * PtgMemArea is an optimization of referenced areas.  Sweet!
 * *******************************************************************************
 * NOTE:  Below is from documentation but DOES NOT APPEAR to be what happens in actuality;
 * PtgMemArea token is followed by several ptg reference-types plus ptgunion(s), ends with a PtgParen.
 * The cce field is the length of all of these following tokens.
 * These following Ptgs are set and parsed in .setPostRecord
 * <p/>
 * <p/>
 * <p/>
 * Like most optimizations it really sucks.  It is also one of the few Ptg's that
 * has a variable length.
 * <p/>
 * Format of length section
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           (reserved)     4       Whatever it may be
 * 2           cce			   2	   length of the reference subexpression
 *
 * Format of reference Subexpression
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0			cref		2			The number of rectangles to follow
 * 2			rgref		var			An Array of rectangles
 *
 * Format of Rectangles
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     2       The First row of the reference
 * 2           rwLast     2       The Last row of the reference
 * 4           ColFirst    1       (see following table)
 * 6           ColLast    1       (see following table)
 * </pre>
 *
 * @see Ptg
 * @see Formula
 */
public class PtgMemArea extends GenericPtg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6869393084367355874L;
	int cce = 0;
	Stack subexpression = null;
	Ptg[] ptgs = null;

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	public int getnTokens()
	{
		return cce;
	} // KSC:

	/**
	 * sets the bytes to describe the subexpression (set of ptgs)
	 * and parses the subexpression
	 *
	 * @param b
	 */
	public void setSubExpression( byte[] b )
	{
		byte[] retbytes = new byte[7];
		System.arraycopy( record, 0, retbytes, 0, record.length );
		retbytes = ByteTools.append( b, retbytes );
		record = retbytes;
		this.populateVals();
		// TODO: 
	}

	// really no need to keep postexpression as can regenerate easily ...
	byte[] postExp = null;

	/**
	 * sets the PtgExtraMem structure which is appended to the end of the function array
	 *
	 * @param b
	 */
	public void setPostExpression( byte[] b, int expressionLen )
	{
		int len = b.length - expressionLen;
		postExp = new byte[len];
		System.arraycopy( b, expressionLen, postExp, 0, len );
/*		
 * 		parsing PtgExtraMem - not really necessary
 * 		
        if (b.length > expressionLen+2) { 		
	 		int z= expressionLen;
			int count= ByteTools.readShort(b[z++], b[z++]);        	
			for (int i= 0; i < count && z+7 < b.length; i++) {
				int[] rc= new int[4];
				rc[0]= ByteTools.readShort(b[z++], b[z++]);	// rw first
				rc[2]= ByteTools.readShort(b[z++], b[z++]);	// rw last
				rc[1]= ByteTools.readShort(b[z++], b[z++]);// col first
				rc[3]= ByteTools.readShort(b[z++], b[z++]); // col last
				ExcelTools.formatRangeRowCol(rc);
			}
		}*/
	}

	/**
	 * retrieves the PtgExtraMem structure that is located at the end of the function record
	 *
	 * @return
	 */
	public byte[] getPostRecord()
	{
		short cce = 0;
		// first count # refs (excluding ptgUnion, range, etc )
		for( int i = 0; i < subexpression.size(); i++ )
		{
			Ptg p = (Ptg) subexpression.get( i );
			if( p instanceof PtgRef )
			{
				cce++;
			}
		}

		// input cce + cce*Ref8U (8 bytes) describing reference
		byte[] b = ByteTools.shortToLEBytes( cce );
		byte[] recbytes = new byte[(cce * 8) + 2];
		recbytes[0] = b[0];
		recbytes[1] = b[1];
		int pos = 2;
		for( int i = 0; i < subexpression.size() && pos + 7 < recbytes.length; i++ )
		{
			Ptg p = (Ptg) subexpression.get( i );
			if( p instanceof PtgRef )
			{
				int[] rc = ((PtgRef) p).getRowCol();
				System.arraycopy( ByteTools.shortToLEBytes( (short) rc[0] ), 0, recbytes, pos, 2 );                // rw first
				System.arraycopy( ByteTools.shortToLEBytes( (short) rc[1] ), 0, recbytes, pos + 4, 2 );            // col first
				if( rc.length == 2 )
				{    // a single ref; repeat
					System.arraycopy( ByteTools.shortToLEBytes( (short) rc[0] ), 0, recbytes, pos + 2, 2 );        // rw last
					System.arraycopy( ByteTools.shortToLEBytes( (short) rc[1] ), 0, recbytes, pos + 6, 2 );        // col last
				}
				else
				{ // a range
					System.arraycopy( ByteTools.shortToLEBytes( (short) rc[2] ), 0, recbytes, pos + 2, 2 );        // rw last
					System.arraycopy( ByteTools.shortToLEBytes( (short) rc[3] ), 0, recbytes, pos + 6, 2 );        // col last
				}
				pos += 8;
			}
		}
/* KSC: TESTING		
		if (!Arrays.equals(recbytes, postExp))
			System.out.println("ISSUE!!!");*/
		return recbytes;
	}

	ArrayList refsheets = new ArrayList();

	void populateVals()
	{
		// 1st byte = ID, next 4 are ignored
		// cce= size of following sub-expressions
		cce = ByteTools.readShort( record[5], record[6] );
		// this is not really correct!
		if( record.length > 7 )
		{
			byte[] subexp;
			subexp = new byte[cce];
			System.arraycopy( record, 7, subexp, 0, cce );
			subexpression = ExpressionParser.parseExpression( subexp, this.parent_rec );
			// subexpression stack in form of:  REFERENCE, REFERENCE, OP [,REFERENCE, OP] ...
			// op can be one of:  PtgUnion [,] PtgIsect [ ] or PtgRange [:]
			// calculate subexpression to obtain ptgs
			try
			{
				Object o = FormulaCalculator.calculateFormula( this.subexpression );
				ArrayList components = new ArrayList();
				if( o != null && o instanceof Ptg[] )
				{
					ptgs = (Ptg[]) o;
					for( int i = 0; i < ptgs.length; i++ )
					{
						if( !refsheets.contains( ((PtgRef) ptgs[i]).getSheetName() ) )
						{
							refsheets.add( ((PtgRef) ptgs[i]).getSheetName() );
						}
						if( ptgs[i] instanceof PtgArea )
						{
							Ptg[] p = ptgs[i].getComponents();
							for( int j = 0; j < p.length; j++ )
							{
								components.add( p[j] );
							}
						}
						else
						{
							components.add( ptgs[i] );
						}
					}
				}
				else
				{    // often a single reference surrounded by parens
					for( int i = 0; i < subexpression.size(); i++ )
					{
						try
						{
							PtgRef pr = (PtgRef) subexpression.get( i );
							if( !refsheets.contains( pr.getSheetName() ) )
							{
								refsheets.add( pr.getSheetName() );
							}
							if( pr instanceof PtgArea )
							{
								Ptg[] pa = pr.getComponents();
								for( int j = 0; j < pa.length; j++ )
								{
									components.add( pa[j] );
								}
							}
							else
							{
								components.add( pr );
							}
						}
						catch( Exception e )
						{
						}
					}
				}
				ptgs = new Ptg[components.size()];
				components.toArray( ptgs );
			}
			catch( Exception e )
			{
				Logger.logErr( "PtgMemArea init: " + e.toString() );
			}

			//int z= subexpression.size();
			// to get # of references (PtgRefs) = stack size/2 + 1
		}
	}

	/**
	 * generate the bytes necessary to describe this PtgMemArea;
	 * extra data described by getPostRecord is necessary for completion
	 *
	 * @see getPostRecord
	 */
	@Override
	public byte[] getRecord()
	{
		int len = 0;
		for( int i = 0; i < subexpression.size(); i++ )
		{
			Ptg p = (Ptg) subexpression.get( i );
			len += p.getRecord().length;
		}
		cce = len;
		byte[] rec = new byte[len + 7];
		byte[] b = ByteTools.shortToLEBytes( (short) cce );
		rec[0] = 0x26;
		// bytes 1-4 are unused
		rec[5] = b[0];
		rec[6] = b[1];
		int offset = 7;
		for( int i = 0; i < subexpression.size(); i++ )
		{
			Ptg p = (Ptg) subexpression.get( i );
			System.arraycopy( p.getRecord(), 0, rec, offset, p.getRecord().length );
			offset += p.getRecord().length;
		}
		record = rec;
		return record;
	}

	@Override
	public int getLength()
	{
		return cce + 7;
	}

	//PtgRef[] comps = null; not used anymore; see note below

	/**
	 * Returns all of the cells of this range as PtgRef's.
	 * This includes empty cells, values, formulas, etc.
	 * Note the setting of parent-rec requires finding the cell
	 * the PtgRef refer's to.  If that is null then the PtgRef
	 * will exist, just with a null value.  This could cause issues when
	 * programatically populating cells.
	 * <p/>
	 * NOTE: now obtaining component ptgs is done in populateValues as it is
	 * a more complex operation than simply gathering all referenced ptgs
	 */
	@Override
	public Ptg[] getComponents()
	{
		/*if(comps!=null) // cache
	    		return comps;
    	
        ArrayList v = new ArrayList();
        try {
    	        for (int i= 0; i < ptgs.length; i++) {
    	        	if (ptgs[i] instanceof PtgArea) {
    	        		Ptg[] ps= ((PtgRef) ptgs[i]).getComponents();
    	        		for (int j= 0; j< ps.length; j++)
    	        			v.add(ps[j]);
    	        	} else if (ptgs[i] instanceof PtgRef) {
    	        		v.add((PtgRef) ptgs[i]);
    	        	} else { // it's a PtgName
    	        		Ptg[] pcomps= ((PtgName) ptgs[i]).getComponents();
    	        		for (int j= 0; j<pcomps.length; j++)
    	        			v.add(pcomps[j]);	
    	        	}	        		
    		    }
        }catch (Exception e){Logger.logInfo("calculating formula range value in PtgArea failed: " + e);}       
        comps = new PtgRef[v.size()];
        v.toArray(comps);z
        */
		return ptgs;
	}

	/**
	 * traverse through expression to retrieve set of ranges
	 * either discontiguous union (,), intersected ( ) or regular range (:)
	 */
	public String toString()
	{
		return FormulaParser.getExpressionString( subexpression ).substring( 1 );    // avoid "="
	}

	@Override
	public Object getValue()
	{
		try
		{
			double[] dub;
			try
			{
				dub = PtgCalculator.getDoubleValueArray( ptgs );
			}
			catch( CalculationException e )
			{
				return null;
			}
			double result = 0.0;
			for( int i = 0; i < dub.length; i++ )
			{
				result += dub[i];
			}
			return new Double( result );
		}
		catch( FunctionNotSupportedException e )
		{
			Logger.logWarn( "Function Unsupported error in PtgMemFunction: " + e );
			return null;
		}
	}

}
