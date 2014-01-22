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

import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.toolkit.ByteTools;

/**
 * Displays "special" attributes like spaces and "optimized SUMs"
 * <p/>
 * Offset Size Contents
 * 0 1 19H
 * 1 1 Attribute type flags:
 * 01H = This is a tAttrVolatile token (volatile function)
 * 02H = This is a tAttrIf token (IF function control)
 * 04H = This is a tAttrChoose token (CHOOSE function control)
 * 08H = This is a tAttrSkip token (skip part of token array)
 * 10H = This is a tAttrSum token (SUM function with one parameter)
 * 20H = This is a tAttrAssign token (assignment-style formula in a macro sheet)
 * 40H = This is a tAttrSpace token (spaces and carriage returns, BIFF3-BIFF8)
 * 41H = This is a tAttrSpaceVolatile token (BIFF3-BIFF8, see below)
 * 2 var. Additional information dependent on the attribute type
 * <p/>
 * tAttrSpace:
 * 0 1 19H
 * 1 1 40H (identifier for the tAttrSpace token), or
 * 41H (identifier for the tAttrSpaceVolatile token)
 * 2 1 Type and position of the inserted character(s):
 * 00H = Spaces before the next token (not allowed before tParen token)
 * 01H = Carriage returns before the next token (not allowed before tParen token)
 * 02H = Spaces before opening parenthesis (only allowed before tParen token)
 * 03H = Carriage returns before opening parenthesis (only allowed before tParen token)
 * 04H = Spaces before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
 * 05H = Carriage returns before closing parenthesis (only allowed before tParen, tFunc, and tFuncVar tokens)
 * 06H = Spaces following the equality sign (only in macro sheets)
 * 3 1 Number of inserted spaces or carriage returns
 *
 * @see Ptg
 * @see Formula
 */
public class PtgAtr extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2825828785221803436L;
	byte grbit = 0x0;
	//    int bitAttrSemi     = 0;		// changed to BitAttrVolatile
	int bitAttrVolatile = 0;
	int bitAttrIf = 0;
	int bitAttrChoose = 0;
	int bitAttrGoto = 0;        // == bitAttrSkip
	int bitAttrSum = 0;
	int bitAttrAssign = 0;        // changed from bitAttrBaxcel
	int bitAttrSpace = 0;
	int bitAttrSpaceVolatile = 0;    // added

	public PtgAtr()
	{
	}

	public PtgAtr( byte type )
	{
		record = new byte[4];
		record[0] = 0x19;
		record[1] = type;
	}

	public String getString()
	{
	    /* We may not want any text from this record.. */
		this.init();
		if( bitAttrIf > 0 )
		{
			return "";  //"IF("; this is already taken care of by another ptg.
		}
		if( bitAttrSum > 0 )
		{
			return "SUM(";
		}
// 	if(bitAttrSemi      > 0)   return "SEMI(";
		if( bitAttrVolatile > 0 )
		{
			return "";
		}
		if( bitAttrAssign > 0 )
		{
			return " EQUALS ";
		}
		if( bitAttrChoose > 0 )
		{
			return "CHOOSE(";
		}
		if( bitAttrSpace > 0 )
		{
			return " ";
		}
		if( bitAttrGoto > 0 )
		{
			return ""; // this may be wrong, but as far as I can tell it is just internal for calc purposes.
		}
		return "UNKNOWN(";

	}

	public String toString()
	{
		return this.getString() + this.getString2();
	}

	/**
	 * return the human-readable String representation of
	 * the "closing" portion of this Ptg
	 * such as a closing parenthesis.
	 */
	public String getString2()
	{
		if( bitAttrIf > 0 )
		{
			return "";  //")"; this is already taken care of by another ptg.
		}
		if( bitAttrSum > 0 )
		{
			return ")";
		}
//        if(bitAttrSemi      > 0)   return ")";
		if( bitAttrVolatile > 0 )
		{
			return "";
		}
		if( bitAttrAssign > 0 )
		{
			return ")";
		}
		if( bitAttrChoose > 0 )
		{
			return ")";
		}
		if( bitAttrGoto > 0 )
		{
			return ""; // this may be wrong, but as far as I can tell it is just internal for calc purposes.
		}
		return "";
	}

	public boolean getIsControl()
	{
		this.init();
		if( getIsPrimitiveOperator() )
		{
			return false;
		}
	    /* TODO: Rework bitAttrIf.  It optimizes the calculation of if statements
        * should not normally be a big deal, but saves the calculation of one of
        * the result fields if needed.*/
		if( bitAttrIf > 0 )
		{
			return false;
		}
		if( bitAttrSum > 0 )
		{
			return true;
		}
//        if(bitAttrSemi   > 0)return false;
		if( bitAttrVolatile > 0 )
		{
			return false;
		}
		if( bitAttrAssign > 0 )
		{
			return true;
		}
		if( bitAttrGoto > 0 )
		{
			return false;
		}
		if( bitAttrChoose > 0 )
		{
			return false;
		}
		return false;
	}

	/**
	 * is the space special -- does it go between vars?
	 * for now we say sure why not.
	 */
	public boolean getIsPrimitiveOperator()
	{
		this.init();
		if( bitAttrSpace > 0 )
		{
			return true;
		}
		return false;
	}

	public boolean getIsUnaryOperator()
	{
		if( bitAttrIf > 0 )
		{
			return false;
		}
		if( bitAttrChoose > 0 )
		{
			return false;
		}
		return true;
	}

	public boolean getIsOperator()
	{
		return false;
	}

	public boolean getIsSpace()
	{
		if( bitAttrSpace > 0 )
		{
			return true;
		}
		return false;
	}

	public boolean getIsOperand()
	{
		//
		// if(bitAttrSpace > 0)return true;
// Old version?		if(bitAttrSemi   > 0)return true; // it just shows that this is a volatile function
		if( bitAttrVolatile > 0 )
		{
			return true; // it just shows that this is a volatile function
		}

		return false;
	}

	/*
		Sets the grbit for the record
	*///[25, 2, 10, 0]		grbit= 2
	public void init()
	{
		grbit = this.getRecord()[1];
        /* john, the following syntax was not reliable, switched with syntax below....
           bitAttrIf       = ((grbit &     0x2)  >>  4);
        */
		// 20060501 KSC: Changed bitAttrVolatile operation + some names
//        if ((grbit & 0x1)== 0x1){bitAttrSemi = 1;}        
		if( (grbit & 0x1) == 0x1 )
		{
			bitAttrVolatile = 1;
		}  // volatile= a function that needs to be recalculated always, such as NOW()
		if( (grbit & 0x2) == 0x2 )
		{
			bitAttrIf = 1;
		}
		if( (grbit & 0x4) == 0x4 )
		{
			bitAttrChoose = 1;
		}
		if( (grbit & 0x8) == 0x8 )
		{
			bitAttrGoto = 1;
		}
		if( (grbit & 0x10) == 0x10 )
		{
			bitAttrSum = 1;
		}
		if( (grbit & 0x20) == 0x20 )
		{
			bitAttrAssign = 1;
		}    // changed name from bitAttrBaxcel
		if( (grbit & 0x40) == 0x40 )
		{
			bitAttrSpace = 1;
		}
		if( (grbit & 0x41) == 0x41 )
		{
			bitAttrSpaceVolatile = 1;
		}

	}

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 * <p/>
	 * public String getString(){
	 * byte[] br = this.getRecord();
	 * byte[] db = new byte[br.length-1]; // strip opcode
	 * System.arraycopy(br, 1, db, 0, db.length);
	 * return ""; // new String(db);
	 * }
	 */
	public int getLength( byte[] b )
	{
		if( (b[0] & 0x4) == 0x4 )
		{
			int i = ByteTools.readShort( b[1], b[2] );
			i += 4;
			return i;
		}
		return getLength();
	}

	public int getLength()
	{
		return PTG_ATR_LENGTH;
	}

	Ptg[] alloperands = null; // cached!

	/*
		Calculate the value of this ptg.
	*/
	public Ptg calculatePtg( Ptg[] pthing ) throws FunctionNotSupportedException, CalculationException
	{
		Ptg returnPtg = null;
		if( this.bitAttrSum > 0 )
		{
			returnPtg = MathFunctionCalculator.calcSum( pthing );
		}
		return returnPtg;
	}
}