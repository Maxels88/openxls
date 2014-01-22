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

import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.ExpressionParser;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.FastAddVector;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;
import java.util.Stack;

/**
 * PtgMemFunc refers to a reference subexpression that doesn't evaluate
 * to a constant reference.  This is still somewhat unclear to me how it functions,
 * or why it exists for that matter.
 * <p/>
 * This token encapsulates a reference subexpression that results in a non-constant cell address,
 * cell range address, or cell range list. For
 * <p/>
 * A little update on it.  Apparently this is used in situations where a record only contains one ptg, but needs to refer to a whole stack of them.
 * An example is in the Name record.  For a built in name that has both row & col repeat regions, the name expression is a ptgmemfunc, but contains
 * 2 ptgArea3d's.
 * Also used in Ai's (Chart Series Range Refs) when non-contiguous range refs are required
 * <p/>
 * PtgMemFunc basically represents a complex range and is used where only one PtgRef-type ptg is expected
 * <p/>
 * NOTE: that this represents a NON-CONSTANT expression while PtgMemArea represents a CONSTANT expression
 * <p/>
 * <pre>
 * 	OFFSET		NAME		SIZE		CONTENTS
 * -------------------------------------------------------
 * 	0			cce			2			The length of the reference subexpression
 */
public class PtgMemFunc extends GenericPtg
{

	public static final long serialVersionUID = 666555444333222l;

	Stack<?> subexpression = null; //

	Ptg[] ptgs = null;    // 20090905 KSC: can be PtgRef3d, PtgArea3d, PtgName  ...

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	@Override
	public boolean getIsReference()
	{
		return true;
	}    // 20100202 KSC

	int cce;
	byte[] subexp;

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		try
		{
			this.populateVals();
		}
		catch( Exception e )
		{
			Logger.logErr( "PtgMemFunc init: " + e.toString() );
		}
	}

	ArrayList<String> refsheets = new ArrayList<String>();

	void populateVals() throws Exception
	{
		cce = ByteTools.readShort( record[1], record[2] );
		subexp = new byte[cce];
		System.arraycopy( record, 3, subexp, 0, cce );
		// subexpression stack in form of:  REFERENCE, REFERENCE, OP [,REFERENCE, OP] ...
		// op can be one of:  PtgUnion [,] PtgIsect [ ] or PtgRange [:]
		subexpression = ExpressionParser.parseExpression( subexp, this.parent_rec );
		// try parsing/calculating on-demand rather than upon init
		//parseSubexpression();
	}

	/**
	 * update the record internally for ptgmemfunc
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
		byte[] b = new byte[len + 3];
		byte[] leng = ByteTools.shortToLEBytes( (short) len );
		b[0] = 0x29;
		b[1] = leng[0];
		b[2] = leng[1];
		int offset = 3;
		for( int i = 0; i < subexpression.size(); i++ )
		{
			Ptg p = (Ptg) subexpression.get( i );
			System.arraycopy( p.getRecord(), 0, b, offset, p.getRecord().length );
			offset += p.getRecord().length;
		}
		record = b;
		return record;
	}

	@Override
	public int getLength()
	{
		return cce + 3;
	}

	int calc_id = 1;

	@Override
	public Object getValue()
	{
		if( ptgs == null )
		{
			parseSubexpression();    // not parsed yet
		}
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

	PtgRef[] colrefs = null;

	/**
	 * return the ptg components for a certain column within a ptgArea()
	 *
	 * @param colNum
	 * @return all Ptg's within colNum
	 */
	public Ptg[] getColComponents( int colNum )
	{
		if( colrefs != null ) // cache
		{
			return colrefs;
		}
		FastAddVector v = new FastAddVector();
		Ptg[] allComponents = this.getComponents();
		for( int i = 0; i < allComponents.length; i++ )
		{
			PtgRef p = (PtgRef) allComponents[i];
//			 TODO: check rc sanity here
			int[] x = p.getIntLocation();
			if( x[1] == colNum )
			{
				v.add( p );
			}
		}
		colrefs = new PtgRef[v.size()];
		v.toArray( colrefs );
		return colrefs;
	}

	PtgRef[] comps = null;

	/**
	 * parses subexpression into ptgs array + traps referenced sheets
	 */
	private void parseSubexpression()
	{
		// calculate subexpression to obtain ptgs
		Object o = FormulaCalculator.calculateFormula( this.subexpression );
		ArrayList<Ptg> components = new ArrayList<Ptg>();
		if( o != null && o instanceof Ptg[] )
		{
			// Firstly: take subexpression and remove reference-tracked elements; calcualted elements are ref-tracked below
			for( int i = 0; i < subexpression.size(); i++ )
			{
				try
				{
					((PtgRef) subexpression.get( i )).removeFromRefTracker();
				}
				catch( Exception e )
				{
				}
			}
			ptgs = (Ptg[]) o;
			for( int i = 0; i < ptgs.length; i++ )
			{
				try
				{
					if( !refsheets.contains( ((PtgRef) ptgs[i]).getSheetName() ) )
					{
						refsheets.add( ((PtgRef) ptgs[i]).getSheetName() );
					}
					((PtgRef) ptgs[i]).addToRefTracker();
					if( ptgs[i] instanceof PtgArea & !(ptgs[i] instanceof PtgAreaErr3d) )
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
				catch( Exception e )
				{
					Logger.logErr( "PtgMemFunc init: " + e.toString() );
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
		if( ptgs == null )
		{
			parseSubexpression();    // not parsed yet
		}
		return ptgs;
	}

	/**
	 * @return Returns the firstPtg.
	 */
	public Ptg getFirstloc()
	{
		if( ptgs == null )
		{
			parseSubexpression();    // not parsed yet
		}
		if( ptgs != null )
		{
			return ptgs[0];
		}
		return null;
	}

	/**
	 * Ptgs upkeep their mapping in reference tracker, however, some ptgs
	 * are components of other Ptgs, such as individual ptg cells in a PtgArea.  These
	 * should not be stored in the RT.
	 */
	private boolean useReferenceTracker = true;

	public void setUseReferenceTracker( boolean b )
	{
		useReferenceTracker = b;
	}

	public boolean getUseReferenceTracker()
	{
		return useReferenceTracker;
	}

	/**
	 * @return Returns the lastPtg.
	 */
	public Stack<?> getSubExpression()
	{
		return subexpression;
	}

	/**
	 * given a complex range, parse and set this PtgMemFunc's associated ptgs
	 *
	 * @param String complexrange  String representing a complex range
	 */
	// some possibilities
	// a:b,c:d
	// a:b c:d,e:f g:h
	// a, b, c, d
	// Q34:Q36:Q35:Q37 Q36:Q38
	// name1:name2:name3:name4
	/*
	 * results of parsing complex ranges from Excel:
     *  @V@27, ), @V$27, ), " "		=(($V$27) ($V$27))
		Q27:Q29, Q28:Q30, ","		'=SUM((Q27:Q29,Q28:Q30))
		Q27:Q29, Q28:Q30, " "		'=SUM(Q27:Q29 Q28:Q30)
		Q27:Q29, Q28:Q30, ":"		'=SUM(Q27:Q29:Q28:Q30)
		Q34:Q36, Q35:Q37, ":", Q36:Q38, " ", Q37:Q39, ","	'=SUM((Q34:Q36:Q35:Q37 Q36:Q38,Q37:Q39))
		Q34:Q36, Q35:Q37, ":", ), Q36:Q38, " ", Q37:Q39, ","	'=SUM(((Q34:Q36:Q35:Q37) Q36:Q38,Q37:Q39))
		Q34:Q36, Q35:Q37, Q36:Q38, " ", ), ":", Q37:Q39, ","	'=SUM((Q34:Q36:(Q35:Q37 Q36:Q38),Q37:Q39))
		Q34:Q36, Q35:Q37, ":", Q36:Q38, Q37:Q39, ",", ), " "	'=SUM((Q34:Q36:Q35:Q37 (Q36:Q38,Q37:Q39)))
		Q34:Q36, Q35:Q37, ":", Q36:Q38, " ", ), Q37:Q39, ","	'=SUM(((Q34:Q36:Q35:Q37 Q36:Q38),Q37:Q39))
		Q34:Q36, Q35:Q37, Q36:Q38, " ", Q37:Q39, ",", ), ":"	'=SUM((Q34:Q36:(Q35:Q37 Q36:Q38,Q37:Q39)))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", " ", ","		'=SUM((Q40:Q42,Q41:Q43 Q42:Q44:Q43:Q45))
		Q40:Q42, Q41:Q43, ",", ), Q42:Q44, Q43:Q45, ":", " "	'=SUM(((Q40:Q42,Q41:Q43) Q42:Q44:Q43:Q45))
		Q40:Q42, Q41:Q43, " ", ), Q43:Q44, ":", Q45, ":", ","	'=SUM((Q40:Q42,(Q41:Q43 Q42):Q43:Q44:Q45))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", ), " ", ","	'=SUM((Q40:Q42,Q41:Q43 (Q42:Q44:Q43:Q45)))
		Q40:Q42, Q41:Q43, Q42:Q44, " ", ",", ), Q43:Q45, ":"	'=SUM(((Q40:Q42,Q41:Q43 Q42:Q44):Q43:Q45))
		Q40:Q42, Q41:Q43, Q42:Q44, Q43:Q45, ":", " ", ), ","	'=SUM((Q40:Q42,(Q41:Q43 Q42:Q44:Q43:Q45)))
		Q46, Q46, " "											'=-Q46 Q46
		Q47, ), Q47, ":"										'=-(Q47):Q47
		Q48, Q48, ), ":"										'=-Q48:(Q48)
     */
	@Override
	public void setLocation( String complexrange )
	{
		byte[] newData = new byte[3];    // 1st 3 bytes= id + cce (length of following data)
		String sheetName = "";
		WorkBook bk = null;
		ArrayList<String> sheets = new ArrayList<String>();
		try
		{
			bk = this.getParentRec().getWorkBook();
			for( int i = 0; i < bk.getSheetVect().size(); i++ )
			{
				sheets.add( bk.getWorkSheetByNumber( i ).getSheetName() );
			}
			sheetName = this.getParentRec().getSheet().getSheetName() + "!";
		}
		catch( Exception e )
		{
			//?
		}

		if( complexrange.startsWith( "(" ) && complexrange.endsWith( ")" ) )    // memfuncs are assumed to be wrapped in parens, no need to specify
		{
			complexrange = complexrange.substring( 0, complexrange.length() - 1 );
		}

		// KSC: TESTING: revert settng subsexpression here for now as tests fail ((;			
		//this.subexpression= new Stack();
		Stack<Comparable> refs = parseFmla( complexrange );
		try
		{
			// structure:
			// ref, ref, op, [op?] [, ref, op ...]
			String ref;
			while( refs.size() != 0 )
			{
				while( refs.size() > 0 )
				{
					if( refs.get( 0 ) instanceof Character )
					{ // it's an operator
						Character cOp = (Character) refs.get( 0 );
						if( cOp.charValue() == ',' )
						{
							PtgUnion pu = new PtgUnion();
							cce += pu.getRecord().length;
							newData = ByteTools.append( pu.getRecord(), newData );
							//this.subexpression.add(pu);
						}
						else if( cOp.charValue() == ' ' )
						{
							PtgIsect pi = new PtgIsect();
							cce += pi.getRecord().length;
							newData = ByteTools.append( pi.getRecord(), newData );
							//this.subexpression.add(pi);
						}
						else if( cOp.charValue() == ':' )
						{
							PtgRange pr = new PtgRange();
							cce += pr.getRecord().length;
							newData = ByteTools.append( pr.getRecord(), newData );
							//this.subexpression.add(pr);
						}
						else if( cOp.charValue() == ')' )
						{
							PtgParen pp = new PtgParen();
							cce += pp.getRecord().length;
							newData = ByteTools.append( pp.getRecord(), newData );
							//this.subexpression.add(pp);
						}
					}
					else
					{
						Object o = refs.get( 0 );
						if( o instanceof Ptg )
						{    // in the rare case of PtgMemFuncs which contain embedded formulas, Ptgs are already created (see parseFmla)
							cce += ((Ptg) o).getRecord().length;
							newData = ByteTools.append( ((Ptg) o).getRecord(), newData );
							//this.subexpression.add(o);
						}
						else
						{
							ref = (String) o;
							boolean isName = (this.getParentRec().getWorkBook().getName( ref ) != null);
							Ptg p = null;
							if( isName )
							{
								p = new PtgName();
								p.setParentRec( this.parent_rec );
								((PtgName) p).setName( ref );
								cce += p.getRecord().length;
								newData = ByteTools.append( p.getRecord(), newData );
							}
							else if( ref.indexOf( ":" ) > 0 )
							{ // TODO: handle in quote!!!
								if( ref.indexOf( "!" ) == -1 )
								{
									ref = sheetName + ref;
								}
								p = new PtgArea3d();
								p.setParentRec( this.parent_rec );
								p.setLocation( ref );
								cce += p.getRecord().length;
								newData = ByteTools.append( p.getRecord(), newData );
							}
							else
							{
								if( ref.indexOf( "!" ) == -1 )
								{
									ref = sheetName + ref;
								}
								p = new PtgRef3d();
								p.setParentRec( this.parent_rec );
								p.setLocation( ref );
								((PtgRef3d) p).setPtgType( PtgRef.REFERENCE );//important for charting/ptgmemfuncs in series/categories - will error on open otherwise
								cce += p.getRecord().length;
								newData = ByteTools.append( p.getRecord(), newData );
							}
							//this.subexpression.add(p);
						}
					}
					refs.remove( 0 );
				}
			}
		}
		catch( Exception e )
		{
			throw new IllegalArgumentException( "PtgMemFunc Error Parsing Location " + complexrange + ":" + e.toString() );
		}
		byte[] ix = ByteTools.shortToLEBytes( (short) cce );
		System.arraycopy( ix, 0, newData, 1, 2 );
		newData[0] = 41;    // ptgId
		record = newData;
		try
		{
			// KSC: don't re-parse as already have all the ptgs ... also, rw/col bytes for Excel2007 exceed maximums so conversion can't be 100%
			// KSC: TESTING: revert for now tests fail ((;			
			populateVals();
/*			cce = ByteTools.readShort(record[1], record[2]);
			subexp = new byte[cce];
	        System.arraycopy(record, 3, subexp, 0, cce);/**/
		}
		catch( Exception e )
		{
			Logger.logErr( "PtgMemFunc setLocation failed for: " + complexrange + " " + e.toString() );
		}

	}

	/**
	 * takes a range string which may contain operators:
	 * union [,] isect [ ] range [:} or paren
	 * plus range elements and/or named range
	 * parse and order each element correctly
	 * may be called recurrsively
	 * NOTE: may also be VERY complex, of type OFFSET(x,y,0):OFFSET(z,w,0)
	 * ALSO INDEX and INDIRECT ...
	 *
	 * @param complexrange
	 * @return ordered stack containing parsed range elements
	 */
	private Stack<Comparable> parseFmla( String complexrange )
	{
		Stack<Comparable> ops = new Stack<Comparable>();
		int lastOp = 0;
		boolean finishRange = false;
		Stack<Comparable> refs = new Stack<Comparable>();
		String ref = "";
		boolean inquote = false;
		String range = null;    // holds partial range
		for( int i = 0; i < complexrange.length(); i++ )
		{
			char c = complexrange.charAt( i );
			if( c == '\'' )
			{
				inquote = !inquote;
				ref += c;
			}
			else if( !inquote )
			{
				if( c == ',' || c == ' ' || c == ')' || (c == ':' && finishRange) )
				{    // it's an operand
					if( c == ' ' && lastOp == ' ' )
					{
						continue;    // skip 2nd space op (Isect)
					}
					if( finishRange )
					{ // add ref to rest of range
						refs.push( range + ref );
						if( !refs.isEmpty() && !ops.isEmpty() )
						{
							refs = handleOpearatorPreference( refs, ops );
						}
						while( !ops.isEmpty() )
						{
							refs.push( ops.pop() );
						}
						range = null;
						ref = "";
						finishRange = false;
						ops.push( new Character( (char) c ) );
					}
					else if( refs.isEmpty() )
					{    // no operands yet - put in 1st
						if( !ref.equals( "" ) )
						{
							refs.push( ref );
						}
						ref = "";
						ops.push( new Character( (char) c ) );
					}
					else
					{    // have all we need to process
						if( !ref.equals( "" ) )
						{
							refs.push( ref );
						}
						while( !ops.isEmpty() )
						{
							refs.push( ops.pop() );
						}
						ref = "";    // handle case of two spaces ... unfortunately
						ops.push( new Character( (char) c ) );
					}
					lastOp = c;
				}
				else if( c == ':' )
				{
					if( this.getParentRec().getWorkBook().getName( ref ) == null )
					{ // it's a regular range
						// check if the ref is a sheet name in a 3d ref
						if( !ref.equals( "" ) )
						{
							range = ref + c;
							finishRange = true;
						}
						else
						{ // happens in cases such as (opopop):ref:ref
							ops.push( new Character( (char) c ) );
						}
						ref = "";
					}
					else
					{    // it's a named range
						refs.push( ref );
						ref = "";
						ops.push( new Character( (char) c ) );
						finishRange = false;    // it's not a regular range
					}
				}
				else if( c == '(' )
				{
					int endparen = FormulaParser.getMatchOperator( complexrange, i, '(', ')' );
					if( endparen == -1 )
					{
						endparen = complexrange.length() - 1;
					}
					else if( !ref.equals( "" ) )
					{
						// rare case of a PtgMemFunc containing a formula:
						String f = ref + "(" + complexrange.substring( i + 1, endparen + 1 );
						ref = "";
						refs = mergeStacks( refs, FormulaParser.getPtgsFromFormulaString( this.getParentRec(), f, true ) );
						i = endparen;
						if( !ops.isEmpty() )
						{
							refs = handleOpearatorPreference( refs, ops );
						}
						while( !ops.isEmpty() )
						{
							refs.push( ops.pop() );
						}
						continue;
					}
					refs = mergeStacks( refs, parseFmla( complexrange.substring( i + 1, endparen + 1 ) ) );
					i = endparen;
					if( !ops.isEmpty() )
					{
						refs = handleOpearatorPreference( refs, ops );
					}
					while( !ops.isEmpty() )
					{
						refs.push( ops.pop() );
					}
				}
				else
				{
					ref += c;
				}
			}
			else
			{
				ref += c;
			}
		}
		// get any remaining
		if( finishRange )
		{ // add ref to rest of range
			// range op has more precedence than others ...
			if( !ops.isEmpty() && ((Character) ops.peek()).charValue() == ':' && !refs.isEmpty() &&
					refs.peek() instanceof Character )
			{
				while( refs.peek() instanceof Character )
				{
					if( ((Character) refs.peek()).charValue() != ':' )
					{
						ops.add( 0, refs.pop() );
					}
					else
					{
						break;
					}
				}
			}
			if( !ref.equals( "" ) )
			{
				refs.push( range + ref );
			}
			else
			{
				refs.push( range.substring( 0, range.length() - 1 ) );
				ops.push( ':' );
			}
		}
		else
		{
			if( !ref.equals( "" ) )
			{
				refs.push( ref );
			}
		}
		while( !ops.isEmpty() )
		{
			refs.push( ops.pop() );
		}
		return refs;
	}

	/**
	 * handle precedence of complex range operators:  : before , before ' '
	 *
	 * @param sourceStack
	 * @param destStack
	 */
	private static Stack<Comparable> handleOpearatorPreference( Stack<Comparable> refs, Stack<Comparable> ops )
	{
		char lastOp = ((Character) ops.pop()).charValue();
		if( refs.peek() instanceof Character )
		{
			char curOp = ((Character) refs.pop()).charValue();
			int group1 = rankPrecedence( lastOp );
			int group2 = rankPrecedence( curOp );
			if( group2 >= group1 )
			{
				ops.push( new Character( lastOp ) );
				refs.push( new Character( curOp ) );
			}
			else
			{
				ops.push( new Character( curOp ) );
				refs.push( new Character( lastOp ) );
			}

		}
		else
		{
			ops.push( new Character( lastOp ) );
		}
		return refs;
	}

	/**
	 * rank a Ptg Operator's precedence (lower
	 *
	 * @param curOp
	 * @return
	 */
	static int rankPrecedence( char curOp )
	{
		if( curOp == 0 )
		{
			return -1;
		}
		if( curOp == ')' )
		{
			return 6;
		}
		if( curOp == ':' )
		{
			return 5;
		}
		if( curOp == ',' || curOp == ' ' )    // same level????
		{
			return 4;
		}
		return 0;    // ' '
	}

	/**
	 * when parenthesed sub-functions
	 *
	 * @param first
	 * @param last
	 * @return
	 */
	private Stack<Comparable> mergeStacks( Stack<Comparable> first, Stack<Comparable> last )
	{
		first.addAll( last );
		return first;
	}

	/**
	 * traverse through expression to retrieve set of ranges
	 * either discontiguous union (,), intersected ( ) or regular range (:)
	 */
	public String toString()
	{
		return FormulaParser.getExpressionString( subexpression ).substring( 1 );    // avoid "="
	}

	/**
	 * return the boundsheet associated with this complex range
	 * <br>NOTE: since complex ranges may contain more than one sheet, this is incomplete for those instanaces
	 *
	 * @param b
	 * @return
	 */
	public Boundsheet[] getSheets( WorkBook b )
	{
		if( ptgs == null )
		{
			parseSubexpression();
		}
		if( this.refsheets != null || this.refsheets.size() != 0 )
		{
			try
			{
				Boundsheet[] sheets = new Boundsheet[this.refsheets.size()];
				for( int i = 0; i < sheets.length; i++ )
				{
					sheets[i] = b.getWorkSheetByName( (String) refsheets.get( i ) );
				}
				return sheets;
			}
			catch( Exception e )
			{
				; // TODO: report error?
			}
		}
		return null;

	}

	@Override
	public void close()
	{
		if( ptgs != null )
		{
			for( int i = 0; i < ptgs.length; i++ )
			{
				if( ptgs[i] instanceof PtgRef )
				{
					((PtgRef) ptgs[i]).close();
				}
				else
				{
					((GenericPtg) ptgs[i]).close();
				}
			}
		}
		ptgs = null;
		super.close();
	}
/*	protected void finalize() {
		this.close();
	}*/
}





