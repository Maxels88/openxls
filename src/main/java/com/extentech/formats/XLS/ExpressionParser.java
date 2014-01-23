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

import com.extentech.formats.XLS.formulas.FunctionConstants;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgAdd;
import com.extentech.formats.XLS.formulas.PtgArea;
import com.extentech.formats.XLS.formulas.PtgArea3d;
import com.extentech.formats.XLS.formulas.PtgAreaErr3d;
import com.extentech.formats.XLS.formulas.PtgAreaN;
import com.extentech.formats.XLS.formulas.PtgArray;
import com.extentech.formats.XLS.formulas.PtgAtr;
import com.extentech.formats.XLS.formulas.PtgBool;
import com.extentech.formats.XLS.formulas.PtgConcat;
import com.extentech.formats.XLS.formulas.PtgDiv;
import com.extentech.formats.XLS.formulas.PtgEQ;
import com.extentech.formats.XLS.formulas.PtgEndSheet;
import com.extentech.formats.XLS.formulas.PtgErr;
import com.extentech.formats.XLS.formulas.PtgExp;
import com.extentech.formats.XLS.formulas.PtgFunc;
import com.extentech.formats.XLS.formulas.PtgFuncVar;
import com.extentech.formats.XLS.formulas.PtgGE;
import com.extentech.formats.XLS.formulas.PtgGT;
import com.extentech.formats.XLS.formulas.PtgInt;
import com.extentech.formats.XLS.formulas.PtgIsect;
import com.extentech.formats.XLS.formulas.PtgLE;
import com.extentech.formats.XLS.formulas.PtgLT;
import com.extentech.formats.XLS.formulas.PtgMemArea;
import com.extentech.formats.XLS.formulas.PtgMemAreaA;
import com.extentech.formats.XLS.formulas.PtgMemAreaN;
import com.extentech.formats.XLS.formulas.PtgMemAreaNV;
import com.extentech.formats.XLS.formulas.PtgMemErr;
import com.extentech.formats.XLS.formulas.PtgMemFunc;
import com.extentech.formats.XLS.formulas.PtgMissArg;
import com.extentech.formats.XLS.formulas.PtgMlt;
import com.extentech.formats.XLS.formulas.PtgMystery;
import com.extentech.formats.XLS.formulas.PtgNE;
import com.extentech.formats.XLS.formulas.PtgName;
import com.extentech.formats.XLS.formulas.PtgNameX;
import com.extentech.formats.XLS.formulas.PtgNumber;
import com.extentech.formats.XLS.formulas.PtgParen;
import com.extentech.formats.XLS.formulas.PtgPercent;
import com.extentech.formats.XLS.formulas.PtgPower;
import com.extentech.formats.XLS.formulas.PtgRange;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.formats.XLS.formulas.PtgRef3d;
import com.extentech.formats.XLS.formulas.PtgRefErr;
import com.extentech.formats.XLS.formulas.PtgRefErr3d;
import com.extentech.formats.XLS.formulas.PtgRefN;
import com.extentech.formats.XLS.formulas.PtgStr;
import com.extentech.formats.XLS.formulas.PtgSub;
import com.extentech.formats.XLS.formulas.PtgUMinus;
import com.extentech.formats.XLS.formulas.PtgUPlus;
import com.extentech.formats.XLS.formulas.PtgUnion;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Stack;
import java.util.Vector;

/**

 */
public final class ExpressionParser implements java.io.Serializable
{
	private static final Logger log = LoggerFactory.getLogger( ExpressionParser.class );
	private static final long serialVersionUID = 4745215965823234010L;
	/*  All of the operand values

		Section of binary operator PTG's.  These pop the two
		top values out of a stack and perform an operation on
		them before pushing back in
	*/
	// really "special" one, read all about it.
	public static final short ptgExp = 0x1;
	public static final short ptgAdd = 0x3;
	public static final short ptgSub = 0x4;
	public static final short ptgMlt = 0x5;
	public static final short ptgDiv = 0x6;
	public static final short ptgPower = 0x7;
	public static final short ptgConcat = 0x8;
	public static final short ptgLT = 0x09;
	public static final short ptgLE = 0x0a;
	public static final short ptgEQ = 0x0b;
	public static final short ptgGE = 0x0c;
	public static final short ptgGT = 0x0d;
	public static final short ptgNE = 0x0e;
	public static final short ptgIsect = 0x0f;
	public static final short ptgUnion = 0x10;
	public static final short ptgRange = 0x11;
	//End of binary operator PTG's

	//Unary Operator tokens
	public static final short ptgUPlus = 0x12;
	public static final short ptgUMinus = 0x13; //todo
	public static final short ptgPercent = 0x14; //todo

	// Controls
	public static final short ptgParen = 0x15;
	public static final short ptgAtr = 0x19;
	// End of Controls

	// Constant operators
	public static final short ptgMissArg = 0x16;
	public static final short ptgStr = 0x17;
	public static final short ptgEndSheet = 0x1b;
	public static final short ptgErr = 0x1c;
	public static final short ptgBool = 0x1d;
	public static final short ptgInt = 0x1e;
	public static final short ptgNum = 0x1f;
	// End of Constant Operators

	public static final short ptgArray = 0x20;
	public static final short ptgFunc = 0x21;
	public static final short ptgFuncVar = 0x22;
	public static final short ptgName = 0x23;
	public static final short ptgRef = 0x24;
	public static final short ptgArea = 0x25;
	public static final short ptgMemArea = 0x26;
	public static final short ptgMemErr = 0x27;
	public static final short ptgMemFunc = 0x29;
	public static final short ptgRefErr = 0x2a;
	public static final short ptgAreaErr = 0x2b;
	public static final short ptgRefN = 0x2c;
	public static final short ptgAreaN = 0x2d;
	public static final short ptgNameX = 0x39;
	public static final short ptgRef3d = 0x3a;
	public static final short ptgArea3d = 0x3b;
	public static final short ptgRefErr3d = 0x3c;

	// who knows, added to fix broken Named ranges -jm 03/26/04
	public static final short ptgAreaErr3d = 0x3d;
	public static final short ptgMemAreaA = 0x66;
	public static final short ptgMemAreaNV = 0x4e;
	public static final short ptgMemAreaN = 0x2e;

	/**
	 * Parse the byte array, create component Ptg's and insert
	 * them into a stack.
	 * <p/>
	 * <p/>
	 * Feb 8, 2010
	 *
	 * @param function
	 * @param rec
	 * @return
	 */
	public static Stack parseExpression( byte[] function, XLSRecord rec )
	{
		return ExpressionParser.parseExpression( function, rec, function.length );
	}

	/**
	 * Parse the byte array, create component Ptg's and insert them into
	 * a stack.
	 * <p/>
	 * Feb 8, 2010
	 *
	 * @param function
	 * @param rec
	 * @param expressionLen
	 * @return
	 */
	public static Stack parseExpression( byte[] function, XLSRecord rec, int expressionLen )
	{
		Stack stack = new Stack();
		short ptg = 0x0;
		int ptgLen = 0;
		boolean hasArrays = false;
		/* Not really needed
        //boolean hasPtgExtraMem= false;
        //PtgMemArea pma= null;*/
		CompatibleVector arrayLocs = new CompatibleVector();
		if( expressionLen > function.length )
		{
			expressionLen = function.length; // deal with out of spec formulas (testJapanese:Open25.xls) -jm
		}
		// KSC: shared formula changes for peformance: now PtgRefN's/PtgAreaN's are instantiated and reference-tracked (of a sort) ...
		XLSRecord p = rec;    // parent

		// iterate the expression and create Ptgs.
		for( int i = 0; i < expressionLen; )
		{
			// check if the 40 bit is set, is it a Array class?
			if( (function[i] & 0x40) == 0x40 )
			{
				//rec is a value class
				//we need to strip the high-order bits and set the 0x20 bit
				ptg = (short) ((function[i] | 0x20) & 0x3f);
			}
			else
			{
				// the bit is already set, just strip the high order bits
				// rec may be an array class.  need to figure rec one out.
				ptg = (short) (function[i] & 0x3f);
			}
			switch( ptg )
			{

				case ptgExp:
					log.debug( "ptgExp Located" );
					if( i == 0 )
					{// MUST BE THE ONLY PTG in the formula expression
						PtgExp px = new PtgExp();
						ptgLen = px.getLength();
						byte[] b = new byte[ptgLen];
						if( (ptgLen + i) <= function.length )
						{
							System.arraycopy( function, (i), b, 0, ptgLen );
						}
						px.setParentRec( p );
						px.init( b );
						stack.push( px );
						break;
					}
					// ptgStr is one of the only ptg's that varies in length, so there is some special handling
					// going on for it.
				case ptgStr:
					log.debug( "ptgStr Located" );
					int x = i;
					x += 1; // move past the opcode to the cch
					ptgLen = function[x] & 0xff; // this is the cch
					short theGrbit = function[x + 1];// this is the grbit;
					if( (theGrbit & 0x1) == 0x1 )
					{
						// unicode string
						ptgLen = ptgLen * 2;
					}
					ptgLen += 3; // include the PtgId, cch, & grbit;
					PtgStr pst = new PtgStr();
					byte[] b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pst.init( b );
					pst.setParentRec( p );
					stack.push( pst );
					break;
				/* */

				case ptgMemAreaA:
					log.debug( "ptgMemAreaA Located" + function[i] );
					x = i;
					x += 5; // move past the opcode & reserved to the cce
					ptgLen = ByteTools.readShort( function[x], function[x + 1] ); // this is the cce
					ptgLen += 7; // include the PtgId, cce, & reserv;
					PtgMemAreaA pmema = new PtgMemAreaA();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmema.init( b );
					pmema.setParentRec( p );
					stack.push( pmema );
					break;

				case ptgMemAreaN:
					log.debug( "ptgMemAreaN Located" + function[i] );
					PtgMemAreaN pmemn = new PtgMemAreaN();
					ptgLen = pmemn.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmemn.init( b );
					pmemn.setParentRec( p );
					stack.push( pmemn );
					break;

				case ptgMemAreaNV:
					log.debug( "ptgMemAreaNV Located" + function[i] );
					x = i;
					x += 5; // move past the opcode & reserved to the cce
					ptgLen = ByteTools.readShort( function[x], function[x + 1] ); // this is the cce
					ptgLen += 7; // include the PtgId, cce, & reserv;
					PtgMemAreaNV pmemv = new PtgMemAreaNV();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmemv.init( b );
					pmemv.setParentRec( p );
					stack.push( pmemv );
					break;

//				ptgMemArea also varies in length...							
				case ptgMemArea:
					log.debug( "ptgMemArea Located" + function[i] );
					ptgLen = 7;
					PtgMemArea pmem = new PtgMemArea();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmem.init( b );
					// now grab the rest of the "extra data" that defines the ptgmemarea
					// these are separate ptgs (PtgArea, PtgRef's ... plus PtgUnions) 
					// that comprise the PtgMemArea coordinates  
					pmem.setParentRec( p );
					i += ptgLen;    // after PtgMemArea record, get subexpression
					ptgLen = pmem.getnTokens();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmem.setSubExpression( b );
					stack.push( pmem );
					//hasPtgExtraMem= true;	// has a PtgExtraMem structure after end of parsed expression: The PtgExtraMem structure specifies a range that corresponds to a PtgMemArea as specified in RgbExtra.
					//pma= pmem;	// save for later
					break;

				case ptgMemFunc:
					log.debug( "ptgMemFunc Located" );
					PtgMemFunc pmemf = new PtgMemFunc();
					x = i;
					x += 1; // move past the opcode to the cce
					ptgLen = ByteTools.readShort( function[x], function[x + 1] ); // this is the cce
					ptgLen += 3; // include the PtgId, cce, & reserv;
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pmemf.setParentRec( p );
					pmemf.init( b );
					stack.push( pmemf );
					break;

				case ptgInt:
					log.debug( "ptgInt Located" );
					PtgInt pi = new PtgInt();
					ptgLen = pi.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pi.init( b );
					pi.setParentRec( p );
					stack.push( pi );
					break;

				case ptgErr:
					log.debug( "ptgErr Located" );
					PtgErr perr = new PtgErr();
					ptgLen = perr.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					perr.init( b );
					perr.setParentRec( p );
					stack.push( perr );
					break;

				case ptgNum:
					log.debug( "ptgNum Located" );
					PtgNumber pnum = new PtgNumber();
					ptgLen = pnum.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pnum.init( b );
					pnum.setParentRec( p );
					stack.push( pnum );
					break;

				case ptgBool:
					log.debug( "ptgBool Located" );
					PtgBool pboo = new PtgBool();
					ptgLen = pboo.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pboo.init( b );
					pboo.setParentRec( p );
					stack.push( pboo );
					break;

				case ptgName:
					log.debug( "ptgName Located" );
					PtgName pn = new PtgName();
					ptgLen = pn.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pn.setParentRec( p );
					pn.init( b );
					pn.addListener();
					stack.push( pn );
					int chk = (i + ptgLen);
					if( chk < function.length )
					{
						if( function[i + ptgLen] == 0x0 )
						{
								log.warn( "Undocumented Name Record mystery byte encountered in Formula: " );
							i++;
						}
					}
					break;

				case ptgNameX:
					log.debug( "ptgNameX Located" );
						log.warn( "referencing external spreadsheets unsupported." );
					PtgNameX pnx = new PtgNameX();
					ptgLen = pnx.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pnx.init( b );
					pnx.setParentRec( p );
					pnx.addListener();
					stack.push( pnx );
					break;

				case ptgRef:
					log.debug( "ptgRef Located " );
					PtgRef pt = new PtgRef();
					ptgLen = pt.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pt.setParentRec( p );    // parent rec must be set before init
					pt.init( b );
					pt.addToRefTracker();
					stack.push( pt );
					break;

				case ptgArray:
					hasArrays = true;
					log.debug( "ptgArray Located " );
					PtgArray pa = new PtgArray();
					ptgLen = 8;  //7 len + id
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pa.init( b );    //setArrVals(b);	// 20090820 KSC: b represents base record not array values
					Integer ingr = stack.size();    // constant value array for PtgArray appears at end of stack see hasArrays below
					arrayLocs.add( ingr );
					stack.push( pa );
					break;

				case ptgRefN:
					log.debug( "ptgRefN Located " );
					PtgRefN ptn = new PtgRefN( false );
					ptgLen = ptn.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					ptn.setParentRec( rec );    // parent rec must be set before init
					ptn.init( b );
					if( rec.getOpcode() == XLSConstants.SHRFMLA )
					{
						ptn.addToRefTracker();
					}
					stack.push( ptn );
					break;

				case ptgArea:
					log.debug( "ptgArea Located " );
					PtgArea pg = new PtgArea();
					ptgLen = pg.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pg.setParentRec( p );    // parent rec must be set before init
					pg.init( b );
					pg.addToRefTracker();
					stack.push( pg );
					break;

				case ptgArea3d:
					log.debug( "ptgArea3d Located " );
					PtgArea3d pg3 = new PtgArea3d();
					ptgLen = pg3.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pg3.init( b, p ); // we need this to init the sub-ptgs correctly
					pg3.addListener();
					pg3.addToRefTracker();
					stack.push( pg3 );
					break;

				case ptgAreaN:
					log.debug( "ptgAreaN Located " );
					PtgAreaN pgn = new PtgAreaN();
					ptgLen = pgn.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pgn.setParentRec( rec );
					pgn.init( b );
					if( rec.getOpcode() == XLSConstants.SHRFMLA )
					{
						pgn.addToRefTracker();
					}
					stack.push( pgn );
					break;

				case ptgAreaErr3d:
					log.debug( "ptgAreaErr3d Located" );
					PtgAreaErr3d ptfa = new PtgAreaErr3d();
					ptgLen = ptfa.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					ptfa.setParentRec( p );
					ptfa.init( b );
					//ptfa.addToRefTracker();
					stack.push( ptfa );
					break;

				case ptgRefErr3d:
					log.debug( "ptgRefErr3d Located" );
					PtgRefErr3d ptfr = new PtgRefErr3d();
					ptgLen = ptfr.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					ptfr.setParentRec( p );
					ptfr.init( b );
					//ptfr.addToRefTracker();
					stack.push( ptfr );
					break;

				case ptgMemErr:
					log.debug( "ptgMemErr Located" );
					PtgMemErr pm = new PtgMemErr();
					ptgLen = pm.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pm.setParentRec( p );
					pm.init( b );
					stack.push( pm );
					break;

				case ptgRefErr:
					log.debug( "ptgRefErr Located" );
					PtgRefErr pr = new PtgRefErr();
					ptgLen = pr.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pr.setParentRec( p );    // parent rec must be set before init
					pr.init( b );
					stack.push( pr );
					break;

				case ptgEndSheet:
					log.debug( "ptgEndSheet Located" );
					PtgEndSheet prs = new PtgEndSheet();
					ptgLen = prs.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					prs.init( b );
					prs.setParentRec( p );
					stack.push( prs );
					break;

				case ptgRef3d:
					log.debug( "ptgRef3d Located" );
					PtgRef3d pr3 = new PtgRef3d();
					ptgLen = pr3.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pr3.setParentRec( p );
					pr3.init( b );
					pr3.addListener();
					pr3.addToRefTracker();
					stack.push( pr3 );
					// if an External Link i.e. defined in another workbook, flag formula as such
					if( pr3.isExternalLink() && (p.getOpcode() == XLSConstants.FORMULA) )
					{
						((Formula) p).setIsExternalRef( true );
					}
					break;
                /*
                 * PtgAtr is another one of the ugly size-changing ptg's
                 */
				case ptgAtr:
					PtgAtr pat = new PtgAtr( (byte) 0 );
					log.debug( "PtgAtr Located" );
					ptgLen = pat.getLength();
					if( (function[i + 1] & 0x4) == 0x4 )
					{
						ptgLen = ByteTools.readShort( function[i + 2], function[i + 3] );
						ptgLen++; // one extra for some undocumented reason
						ptgLen = ptgLen * 2; //seems to be two bytes per...
						ptgLen += 4; // add the cch & grbit
					}
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pat.init( b );
					pat.init();
					pat.setParentRec( p );
					stack.push( pat );
					break;

				case ptgFunc:
					log.debug( "ptgFunc Located" );
					PtgFunc ptf = new PtgFunc();
					ptgLen = ptf.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					ptf.init( b );
					ptf.setParentRec( p );
					stack.push( ptf );
					break;

				case ptgFuncVar:
					log.debug( "ptgFuncVar Located" );
					PtgFuncVar ptfv = new PtgFuncVar();
					ptgLen = ptfv.getLength();
					b = new byte[ptgLen];
					if( ((ptgLen) + (i)) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}

					ptfv.init( b );
					ptfv.setParentRec( p );
					if( ptfv.getFunctionId() == FunctionConstants.XLF_INDIRECT )
					{
/*TESTING NEW WAY:                    	
// New way does not account for expanded shared formula references, unfortunately  so keep original for new
 * 	    
  						Stack indirectStack= new Stack();
                    	int z= stack.size()-1;
                    	int nparams= 1;
                    	for (; z > 0 && nparams > 0; z--) {
                    		Ptg p= (Ptg) stack.get(z);
							if (p instanceof PtgAtr) {
								continue;
							}
                            if(p.getIsOperator()||p.getIsControl()||p.getIsFunction()){
                                if(p.getIsControl() ){
                                    if(p.getOpcode() == 0x15) { // its a parens!   
                                        // the parens is already pop'd so just return and it is gone...
                                        continue;
                                    }
                                }
                                int t= 0;
                                // make sure we have the correct amount popped back in..
                                if (p.getIsBinaryOperator()) t=2;
                                if (p.getIsUnaryOperator()) t=1;
                    			if (p.getIsStandAloneOperator()) t=0;
                                if (p.getOpcode() == 0x22 || p.getOpcode() == 0x42 || p.getOpcode() == 0x62){t=p.getNumParams();}// it's a ptgfunkvar!
                                if (p.getOpcode() == 0x21 || p.getOpcode() == 0x41 ||  p.getOpcode() == 0x61){t=p.getNumParams();}// guess that ptgfunc is not only one..
                                nparams+=t-1;
                            	if (nparams==0)
                            		break;
                            } else {
                            	nparams--;
                            	if (nparams==0)
                            		break;
                            }
                    	}
                    	indirectStack.addAll(stack.subList(z, stack.size()));
                    	indirectStack.push(ptfv);
                    	rec.getWorkBook().addIndirectFormulaStack(indirectStack);	// must save and calculate indirect reference AFTER all formulas/cells have been added ...
// original is below                    	                    	
 /**/            			
/**/
						if( rec.getOpcode() == XLSConstants.FORMULA )
						{
							((Formula) rec).setContainsIndirectFunction( true );
						}
						else if( rec.getOpcode() == XLSConstants.SHRFMLA )
						{
							((Shrfmla) rec).setContainsIndirectFunction( true );
						}
 /**/
					}
					stack.push( ptfv );
					break;

				case ptgAdd:
					log.debug( "ptgAdd Located" );
					PtgAdd pad = new PtgAdd();
					ptgLen = pad.getLength();
					b = new byte[ptgLen];
					//if((ptgLen+i) <= function.length)
					System.arraycopy( function, (i), b, 0, ptgLen );
					pad.init( b );
					pad.setParentRec( p );
					stack.push( pad );
					break;

				case ptgMissArg:
					log.debug( "ptgMissArg Located" );
					PtgMissArg pmar = new PtgMissArg();
					ptgLen = pmar.getLength();
					b = new byte[ptgLen];
					//if((ptgLen+i) <= function.length) 
					System.arraycopy( function, (i), b, 0, ptgLen );
					pmar.init( b );
					pmar.setParentRec( p );
					stack.push( pmar );
					break;

				case ptgSub:
					log.debug( "PtgSub Located" );
					PtgSub psb = new PtgSub();
					ptgLen = psb.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					psb.init( b );
					psb.setParentRec( p );
					stack.push( psb );
					break;

				case ptgMlt:
					log.debug( "PtgMlt Located" );
					PtgMlt pml = new PtgMlt();
					ptgLen = pml.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pml.init( b );
					pml.setParentRec( p );
					stack.push( pml );
					break;

				case ptgDiv:
					log.debug( "PtgDiv Located" );
					PtgDiv pdiv = new PtgDiv();
					ptgLen = pdiv.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pdiv.init( b );
					pdiv.setParentRec( p );
					stack.push( pdiv );
					break;

				case ptgUPlus:
					log.debug( "PtgUPlus Located" );
					PtgUPlus puplus = new PtgUPlus();
					ptgLen = puplus.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					puplus.init( b );
					puplus.setParentRec( p );
					stack.push( puplus );
					break;

				case ptgUMinus:
					log.debug( "PtgUminus Located" );
					PtgUMinus puminus = new PtgUMinus();
					ptgLen = puminus.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					puminus.init( b );
					puminus.setParentRec( p );
					stack.push( puminus );
					break;

				case ptgPercent:
					log.debug( "ptgPercent Located" );
					PtgPercent pperc = new PtgPercent();
					ptgLen = pperc.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pperc.init( b );
					pperc.setParentRec( p );
					stack.push( pperc );
					break;

				case ptgPower:
					log.debug( "PtgPower Located" );
					PtgPower pow = new PtgPower();
					ptgLen = pow.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pow.init( b );
					pow.setParentRec( p );
					stack.push( pow );
					break;

				case ptgConcat:
					log.debug( "PtgConcat Located" );
					PtgConcat pcon = new PtgConcat();
					ptgLen = pcon.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pcon.init( b );
					pcon.setParentRec( p );
					stack.push( pcon );
					break;

				case ptgLT:
					log.debug( "PtgLT Located" );
					PtgLT plt = new PtgLT();
					ptgLen = plt.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					plt.init( b );
					plt.setParentRec( p );
					stack.push( plt );
					break;

				case ptgLE:
					log.debug( "PtgLE Located" );
					PtgLE ple = new PtgLE();
					ptgLen = ple.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					ple.init( b );
					ple.setParentRec( p );
					stack.push( ple );
					break;

				case ptgEQ:
					log.debug( "PtgEQ Located" );
					PtgEQ peq = new PtgEQ();
					ptgLen = peq.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					peq.init( b );
					peq.setParentRec( p );
					stack.push( peq );
					break;

				case ptgGE:
					log.debug( "PtgGE Located" );
					PtgGE pge = new PtgGE();
					ptgLen = pge.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pge.init( b );
					pge.setParentRec( p );
					stack.push( pge );
					break;

				case ptgGT:
					log.debug( "PtgGT Located" );
					PtgGT pgt = new PtgGT();
					ptgLen = pgt.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pgt.init( b );
					pgt.setParentRec( p );
					stack.push( pgt );
					break;

				case ptgNE:
					log.debug( "PtgNE Located" );
					PtgNE pne = new PtgNE();
					ptgLen = pne.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}

					pne.init( b );
					pne.setParentRec( p );
					stack.push( pne );
					break;

				case ptgIsect:
					log.debug( "PtgIsect Located" );
					PtgIsect pist = new PtgIsect();
					ptgLen = pist.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}

					pist.init( b );
					pist.setParentRec( p );
					stack.push( pist );
					break;

				case ptgUnion:
					log.debug( "ptgUnion Located" );
					PtgUnion pun = new PtgUnion();
					ptgLen = pun.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pun.init( b );
					pun.setParentRec( p );
					stack.push( pun );
					break;

				case ptgRange:
					log.debug( "ptgRange Located" );
					PtgRange pran = new PtgRange();
					ptgLen = pran.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}
					pran.init( b );
					pran.setParentRec( p );
					stack.push( pran );
					break;

				case ptgParen:
					log.debug( "PtgParens Located" );
					PtgParen pp = new PtgParen();
					ptgLen = pp.getLength();
					b = new byte[ptgLen];
					if( (ptgLen + i) <= function.length )
					{
						System.arraycopy( function, (i), b, 0, ptgLen );
					}

					pp.init( b );
					pp.setParentRec( p );
					stack.push( pp );
					break;

				default:
					PtgMystery pmy = new PtgMystery();
					ptgLen = function.length - i;
					b = new byte[ptgLen];
						log.warn( "Unsupported Formula Function: 0x" + Integer.toHexString( ptg ) + " length: " + ptgLen );
					System.arraycopy( function, i, b, 0, ptgLen );
					pmy.init( b );
					pmy.setParentRec( p );
					stack.push( pmy );
					break;
			}
			i += ptgLen;
		}
		if( hasArrays && (rec instanceof Formula) )
		{    // Array Recs handle extra data differently
			// array data is appended to end of expression
			// for each array in the function list,
			// get saved ptgArray var (stored in stack var),
			// grab data and parse array components
			int startPos = expressionLen;
			for( int i = 0; i < arrayLocs.size(); i++ )
			{
				Integer ingr = (Integer) arrayLocs.elementAt( i );
				PtgArray parr = (PtgArray) stack.elementAt( ingr );

				// have to assume that remaining data all goes for this ptgarray
				// since length is variable and can only be ascertained by parsing
				// if multiple arrays are present, actual array length will be returned via setArrVals
				byte[] b = new byte[function.length - startPos]; // get "extra" array data
				System.arraycopy( function, startPos, b, 0, b.length );
				try
				{
					parr.setParentRec( rec );
					startPos += parr.setArrVals( b );
				}
				catch( Exception e )
				{//TODO: this needs to be caught due to "name" records being parsed incorrectly.  The problem has to do with the lenght of the name record not including the extra 7 bytes of space.  Temporary fix for infoteria
					log.debug( "ExpressionParser.parseExpression: Array: " + e );
				}
			}
		} /* no need to keep PtgExtraMem as can regenerate easily else
        	if (hasPtgExtraMem && rec instanceof Formula) {
        	//The PtgExtraMem structure specifies a range that corresponds to a PtgMemArea as specified in RgbExtra.)
        	//     count (2 bytes): An unsigned integer that specifies the areas within the range.
        	// 	   array (variable): An array of Ref8U that specifies the range. The number of elements MUST be equal to count.
        		pma.setPostExpression(function, expressionLen);
        }*/
		log.debug( "finished formula" );
		return stack;

	}

	/*  Returns the ptg that matches the string location sent to it.
		rec can either be in the format "C5" or a range, such as "C4:D9"

	*/
	public static List getPtgsByLocation( String loc, Stack expression ) throws FormulaNotFoundException
	{
		List lv = new Vector();
		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			if( o == null )
			{
				throw new FormulaNotFoundException( "Couldn't get Ptg at: " + loc );
			}
			if( o instanceof Byte )
			{
				// do nothing
			}
			else if( o instanceof Ptg )
			{
				Ptg part = (Ptg) o;
				String lo = part.getLocation();
				if( lo == null )
				{
					lo = "none";
				}
				String comp = loc;
				if( loc.indexOf( "!" ) > -1 )
				{ // the sheet is referenced
					if( lo.indexOf( "!" ) == -1 )
					{ // and the ptg does not have sheet referenced
						comp = loc.substring( loc.indexOf( "!" ) + 1 );
					}
				}

				if( comp.equalsIgnoreCase( lo ) )
				{
					lv.add( part );
				}
				else
				{
					// try fq location
					lo = part.toString();
					if( loc.equalsIgnoreCase( lo ) )
					{
						lv.add( part );

					}
					else if( o instanceof PtgRef3d )
					{// gotta look into the first & last
						// already checked
					}
					else if( o instanceof PtgArea )
					{// gotta look into the first & last
						Ptg first = ((PtgArea) o).getFirstPtg();
						Ptg last = ((PtgArea) o).getLastPtg();
						if( first.getLocation().equalsIgnoreCase( loc ) )
						{
							lv.add( first );
						}
						if( last.getLocation().equalsIgnoreCase( loc ) )
						{
							lv.add( last );
						}
					}
				}
			}
		}
		return lv;
	}

	/**
	 * returns the position in the expression stack for the ptg associated with this location
	 *
	 * @param loc        String
	 * @param expression
	 * @return
	 * @throws FormulaNotFoundException
	 */
	public static int getExpressionLocByLocation( String loc, Stack expression ) throws FormulaNotFoundException
	{

		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			if( o == null )
			{
				throw new FormulaNotFoundException( "Couldn't get Ptg at: " + loc );
			}
			if( o instanceof Byte )
			{
				// do nothing
			}
			else if( o instanceof Ptg )
			{
				Ptg part = (Ptg) o;
				String lo = part.getLocation();
				if( loc.equalsIgnoreCase( lo ) )
				{
					return i;
				}
				// try full location
				lo = part.toString();
				if( loc.equalsIgnoreCase( lo ) )
				{
					return i;
				}
				if( o instanceof PtgArea )
				{// gotta look into the first & last
					Ptg first = ((PtgArea) o).getFirstPtg();
					Ptg last = ((PtgArea) o).getLastPtg();
					if( first.getLocation().equalsIgnoreCase( loc ) )
					{
						return i;
					}
					if( last.getLocation().equalsIgnoreCase( loc ) )
					{
						return i;
					}
				}
			}
		}
		return -1;
	}

	/**
	 * returns the position in the expression stack for the desired ptg
	 *
	 * @param ptg        Ptg to lookk up
	 * @param expression
	 * @return
	 * @throws FormulaNotFoundException
	 */
	public static int getExpressionLocByPtg( Ptg ptg, Stack expression ) throws FormulaNotFoundException
	{

		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			if( o == null )
			{
				throw new FormulaNotFoundException( "Couldn't get Ptg at: " + ptg.toString() );
			}
			if( o instanceof Byte )
			{
				// do nothing
			}
			else if( o instanceof Ptg )
			{
				if( o.equals( ptg ) )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * getCellRangePtgs handles locating which cells are refereced in an expression stack.
	 * <p/>
	 * Essentially the use is we can check a formula if it refereces a cell that is moving, then we have
	 * the ability to manipulate these ranges in whatever way makes sense.
	 *
	 * @return an array of ptgs that are location based (ptgRef, PtgArea)
	 * @expression = a Stack of ptgs that represent an excel calculation.
	 */
	public static Ptg[] getCellRangePtgs( Stack expression ) throws FormulaNotFoundException
	{
		Vector ret = new Vector();
		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			if( o == null )
			{
				throw new FormulaNotFoundException( "Couldn't get Ptg at: " + i );
			}
			if( o instanceof Byte )
			{
				// do nothing
			}
			else if( o instanceof Ptg )
			{
				Ptg part = (Ptg) o;
				// handle shared formula range
				if( part instanceof PtgExp )
				{
					String lox = part.getLocation();
					PtgRef ref = new PtgRef();
					ref.setParentRec( part.getParentRec() );    // must be done before setLocation
					ref.setLocation( lox );
					ret.add( ref );
				}
				else if( (part instanceof PtgRefErr) || (part instanceof PtgAreaErr3d) )
				{
					ret.add( "#REF!" );
				}
				else if( part instanceof PtgMemFunc )
				{
					//Ptg[] p= getCellRangePtgs(((PtgMemFunc)part).getSubExpression());
					Ptg[] p = part.getComponents();
					for( Ptg aP : p )
					{
						ret.add( aP );
					}
				}
				else
				{
					String lox = part.getLocation();
					if( lox != null )
					{
						ret.add( part );
					}
				}
			}
		}
		Ptg[] retp = new Ptg[ret.size()];
		return (Ptg[]) ret.toArray( retp );
	}

}