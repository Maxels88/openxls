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

import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

import java.util.Locale;

/**
 * PtgFunc is a fuction operator that refers to the header file in order to
 * use the correct function.
 * <p/>
 * PtgFuncVar is only used with a variable number of arguments.
 * <p/>
 * Opcode = 22h
 * <p/>
 * <pre>
 * Offset      Bits    Name        Mask        Contents
 * --------------------------------------------------------
 * 0           6-0     cargs       7Fh         The number of arguments to the function
 * 7       fPrompt     80h         =1, function prompts the user
 * 1           14-0    iftab       7FFFh       The index to the function table
 * see GenericPtgFunc for details
 * 15      fCE         8000h       This function is a command equivalent
 * </pre>
 *
 * @see Ptg
 * @see GenericPtgFunc
 */
public class PtgFuncVar extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1478629759437556620L;
	public static int LENGTH = 3;
	byte ptgId;
	String loc;
	byte cargs;
	boolean fprompt;
	short iftab;
	boolean fCE;

	public PtgFuncVar( int funcType, int numArgs, XLSRecord parentRec )
	{
		this( funcType, numArgs );
		setParentRec( parentRec );
	}

	public PtgFuncVar( int funcType, int numArgs )
	{
		byte[] recbyte = new byte[4];
		// 20100609 KSC: there are three types of funcvars:
		// 22H (tFuncVarR), 42H (tFuncVarV), 62H (tFuncVarA)
		// tFUncVarR = reference return value (most common?)
		// tFuncVarV = value type of return value (ROW, SUM)
		// tFuncVarA = Array return type (TREND)
		// TODO: figure out which other functions are type V or type A
		switch( funcType )
		{
			case FunctionConstants.XLF_ROW:        // ROW
			case FunctionConstants.xlfColumn:        // COLUMN
			case FunctionConstants.xlfIndex:    // INDEX	
			case FunctionConstants.xlfVlookup:    // VLOOKUP
			case FunctionConstants.xlfSumproduct:    // SUMPRODUCT
				recbyte[0] = 0x42;
				break;
			default:        // default= tFuncVarR
				recbyte[0] = 0x22;
		}
		byte[] b = ByteTools.shortToLEBytes( (short) funcType );
		recbyte[1] = (byte) numArgs;
		recbyte[2] = b[0];
		recbyte[3] = b[1];
		init( recbyte );
	}

	public PtgFuncVar()
	{
		// placeholder
	}

	@Override
	public boolean getIsFunction()
	{
		return true;
	}

	/**
	 * Returns the number of Params to pass to the Ptg
	 */
	@Override
	public int getNumParams()
	{
		return cargs;
	}

	/**
	 * set the number of parmeters in the FuncVar record
	 *
	 * @param byte nParams
	 */
	// 20060131 KSC: Added to set # params separately from init
	public void setNumParams( byte nParams )
	{
		record[1] = nParams;
		populateVals();
	}

	// should be handled by super?
	@Override
	public byte getOpcode()
	{
		return ptgId;
	}

	/**
	 * GetString - is this toString, what is it returning?
	 */
	@Override
	public String getString()
	{
		if( iftab != FunctionConstants.xlfADDIN )
		{
			String f = null;
			if( Locale.JAPAN.equals( Locale.getDefault() ) )
			{
				f = FunctionConstants.getJFunctionString( iftab );
			}
			if( f == null )
			{
				f = FunctionConstants.getFunctionString( iftab );
			}
			return f;
		}
		return getAddInFunctionString();
	}

	// KSC: added to handle string version of add-in formulas
	private String getAddInFunctionString()
	{
		if( (vars != null) && (vars[0] instanceof PtgNameX) )
		{
			return vars[0].toString() + "(";
		}
		return "(";
	}

	@Override
	public String getString2()
	{
		return ")";
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = 0x22;
		record = b;
		fprompt = false;
		fCE = false;
		populateVals();
	}

	/**
	 * Get the function ID for this PtgFuncVar
	 *
	 * @return function Id
	 */
	public short getFunctionId()
	{
		return iftab;
	}

	/**
	 * parse all the values out of the byte array and
	 * populate the classes values
	 */
	private void populateVals()
	{

		cargs = record[1];
		if( (cargs & 0x80) == 0x80 )
		{ // is fprompt set?
			fprompt = true;
		}
		cargs = (byte) (cargs & 0x7f);
		iftab = ByteTools.readShort( record[2], record[3] );
		if( (iftab & 0x8000) == 0x8000 )
		{ // is fCE set?
			fCE = true;
		}
		iftab = (short) (iftab & 0x7fff); // cut out the fCE
	}

	public int getVal()
	{
		return iftab;
	}

	// THis will have to be modified when we start modifying the record.
	@Override
	public byte[] getRecord()
	{
		return record;
	}

	@Override
	public int getLength()
	{
		return PTG_FUNCVAR_LENGTH;
	}

	@Override
	public Ptg calculatePtg( Ptg[] pthings ) throws FunctionNotSupportedException, CalculationException
	{
		Ptg[] ptgarr = new Ptg[pthings.length + 1];
		ptgarr[0] = this;
		// add this into the array so the functionHandler has a handle to the function
		System.arraycopy( pthings, 0, ptgarr, 1, pthings.length );
		Ptg resPtg = FunctionHandler.calculateFunction( ptgarr );
		return resPtg;
	}

	/**
	 * return String representation of function id for this funcvar
	 */
	public String toString()
	{
		return "FUNCVAR " + iftab;
	}

	/**
	 * given this specific Func Var, ensure that it's parameters are of the correct Ptg type
	 * <br>Value, Reference or Array
	 * <br>This is necessary when functions are added via String
	 * <br>NOTE: eventually all FuncVars which require a specific type of parameter will be handled here
	 *
	 * @see FormulaParser.adjustParameterIds
	 */
	public void adjustParameterIds()
	{
		if( vars == null )
		{
			return; // no parameters to worry about
		}
		switch( iftab )
		{
			case FunctionConstants.xlfVlookup:
				setParameterType( 0, PtgRef.VALUE );
				setParameterType( 1, PtgRef.REFERENCE );
				setParameterType( 2, PtgRef.VALUE );
				setParameterType( 3, PtgRef.VALUE );
				break;
			case FunctionConstants.xlfColumn:
			case FunctionConstants.XLF_ROW:
				setParameterType( 0, PtgRef.REFERENCE );
				break;
			case FunctionConstants.xlfIndex:
				setParameterType( 0, PtgRef.REFERENCE );
				setParameterType( 1, PtgRef.VALUE );
				break;
			case FunctionConstants.XLF_SUM_IF:
				setParameterType( 0, PtgRef.REFERENCE );
				break;
			case FunctionConstants.xlfSumproduct:
				setParameterType( 0, PtgRef.ARRAY );
				setParameterType( 1, PtgRef.ARRAY );
				break;
			default:
				break;
		}
	}

	/**
	 * utility for adjustParameterIds to set the PtgRef-type or PtgName-type pareameter to the correct type
	 * either PtgRef.REFERENCE, PtgRef.VALUE or PtgRef.ARRAY
	 * dependent upon the function they are used in
	 *
	 * @param n
	 * @param type
	 */
	private void setParameterType( int n, short type )
	{
		if( vars.length > n )
		{
			if( vars[n] instanceof PtgArea )
			{
				((PtgArea) vars[n]).setPtgType( type );
			}
			else if( vars[n] instanceof PtgRef )
			{
				((PtgRef) vars[n]).setPtgType( type );
			}
			else if( vars[n] instanceof PtgName )
			{
				((PtgName) vars[n]).setPtgType( type );
			}
		}
	}
}