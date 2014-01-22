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
 * use the correct function.  PtgFunc is only used with fixed number of argument
 * functions
 * <p/>
 * Opcode = 21h
 * <p/>
 * <pre>
 * Offset      Name        Size        Contents
 * --------------------------------------------------------
 * 0           iftab       2           The index to the function table
 * see GenericPtgFunc for details
 * </pre>
 *
 * @see Ptg
 * @see GenericPtgFunc
 */
public class PtgFunc extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4435263700288188538L;
	/* 20060425 KSC: see FunctionConstants for Formula Consolidation details 
	public static String[][] recArr =  {
		{"Pi","19"},
		{"Round","27"},
		{"Rept","30"},
		{"Mid","31"},
		{"Mod","39"},
		{"Rand","63"},
		{"Date","65"},
		{"Time","66"},
		{"Day","67"},
		{"Now","74"},
		{"Atan2","97"},
		{"Log","109"},
		{"Left","115"},
		{"Right","116"},
		{"Replace","119"},
		{"Exact","117"},
		{"Trim","118"},
		{"Roundup","212"},
		{"MMult","165"},
		{"RoundDown","213"},
		{"today","221"},
		{"Combin","276"},
		{"Floor","285"},
		{"Celing","288"},
		{"Power","337"},
		{"Countif","246"},
		{"Hour", String.valueOf(FunctionHandler.xlfHour)},
		{"Minute", String.valueOf(FunctionHandler.xlfMinute)},
		{"Month", String.valueOf(FunctionHandler.xlfMonth)},
		{"Year", String.valueOf(FunctionHandler.xlfYear)},
		{"Second", String.valueOf(FunctionHandler.xlfSecond)},
		{"Weekday", String.valueOf(FunctionHandler.xlfWeekday)},
		{"Time", String.valueOf(FunctionHandler.xlfTime)},
		{"Now", String.valueOf(FunctionHandler.xlfNow)},
		{"Quartile", String.valueOf(FunctionHandler.xlfQuartile)},
		{"Frequency", String.valueOf(FunctionHandler.xlfFrequency)},
		{"Linest", String.valueOf(FunctionHandler.xlfLinest)},
		{"Correl", String.valueOf(FunctionHandler.xlfCorrel)},
		{"Slope", String.valueOf(FunctionHandler.xlfSlope)},
		{"Intercept", String.valueOf(FunctionHandler.xlfIntercept)},
		{"Pearson", String.valueOf(FunctionHandler.xlfPearson)},
		{"Rsq", String.valueOf(FunctionHandler.xlfRsq)},
		{"Steyx", String.valueOf(FunctionHandler.xlfSteyx)},
		{"Forecast", String.valueOf(FunctionHandler.xlfForecast)},
		{"Trend", String.valueOf(FunctionHandler.xlfTrend)},
		{"Covar", String.valueOf(FunctionHandler.xlfCovar)},
		{"IsNumber", String.valueOf(FunctionHandler.xlfIsnumber)},
		{"DAVERAGE", String.valueOf(FunctionHandler.xlfDaverage)},
		{"DCOUNT", String.valueOf(FunctionHandler.xlfDcount)},
		{"DCOUNTA", String.valueOf(FunctionHandler.xlfDcounta)},
		{"DGET", String.valueOf(FunctionHandler.xlfDget)},
		{"DMIN", String.valueOf(FunctionHandler.xlfDmin)},
		{"DMAX", String.valueOf(FunctionHandler.xlfDmax)},
		{"DPRODUCT", String.valueOf(FunctionHandler.xlfDproduct)},
		{"DSTDEVP", String.valueOf(FunctionHandler.xlfDstdevp)},
		{"DSTDEV", String.valueOf(FunctionHandler.xlfDstdev)},
		{"DSUM", String.valueOf(FunctionHandler.xlfDsum)},
		{"DVAR", String.valueOf(FunctionHandler.xlfDvar)},
		{"DVARP", String.valueOf(FunctionHandler.xlfDvarp)},
		{"SQRT", String.valueOf(FunctionHandler.xlfSqrt)},
		
	};
	*/
	private int iftab = -1;

	public PtgFunc( int funcType, XLSRecord parentRec )
	{
		this( funcType );
		this.setParentRec( parentRec );
	}

	public PtgFunc( int funcType )
	{
		byte[] recbyte = new byte[3];
//    	recbyte[0]= 0x21;
		recbyte[0] = 0x41;        // 20060126 - KSC: Excel seems to need this code for PtgFunc
		byte[] b = ByteTools.shortToLEBytes( (short) funcType );
		recbyte[1] = b[0];
		recbyte[2] = b[1];
		this.init( recbyte );
	}

	public PtgFunc()
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
	 * Unfortunately this seems to vary depending on the formula.
	 * fill in the non-1's as you get them.
	 */
	@Override
	public int getNumParams()
	{
		// 20060425 KSC: Formula consolidation - see FunctionConstants for details
		return FunctionConstants.getNumParams( iftab );
	}

	/*
			if (iftab == 10) return 0; // na
			if (iftab == 19) return 0; // Pi
			if (iftab == 27) return 2; // Round
			if (iftab == 30) return 2; // rept
			if (iftab == 31) return 3; // Mid
			if (iftab == 39) return 2; // Mod
			if (iftab >= 40 && iftab <= 45) return 3; // Dxxx formulas
			if (iftab == 47) return 3; // DVar
			if (iftab == 63) return 0; // Rand
			if (iftab == 65) return 3; // Date
			if (iftab == 66) return 3; // Time
			if (iftab == 67) return 1; // Day
			if (iftab == 74) return 0; // now
			if (iftab == 97) return 2; // Atan2
			if (iftab == 109) return 2; // Log
			if (iftab == 115) return 2; // Left
			if (iftab == 116) return 2; // Right
			if (iftab == 118) return 1; // Trim
			if (iftab == 119) return 4; // Log
			if (iftab == 117) return 2; // Exact
			if (iftab == 165) return 2;  //TODO:
			if (iftab == 189) return 3; // DProduct
			if (iftab == 195) return 3; // DStdDev
			if (iftab == 196) return 3; // DVarP
			if (iftab == 199) return 3; // DCountA
			if (iftab == 212) return 2; // Roundup
			if (iftab == 213) return 2; // Rounddown
			if (iftab == 221) return 0; // today
			if (iftab == 235) return 3; // DGet
			if (iftab == 276) return 2; // Combin
			if (iftab == 285) return 2; // Floor
			if (iftab == 288) return 2; // Ceiling
			if (iftab == 337) return 2; // Power
			if (iftab == FunctionHandler.xlfTime) return 3;
			if (iftab == FunctionHandler.xlfQuartile) return 2;
			if (iftab == FunctionHandler.xlfFrequency) return 2;
			if (iftab == FunctionHandler.xlfCorrel) return 2;
			if (iftab == FunctionHandler.xlfCovar) return 2;
			if (iftab == FunctionHandler.xlfSlope) return 2;
			if (iftab == FunctionHandler.xlfIntercept) return 2;
			if (iftab == FunctionHandler.xlfPearson) return 2;
			if (iftab == FunctionHandler.xlfRsq) return 2;
			if (iftab == FunctionHandler.xlfSteyx) return 2;
			if (iftab == FunctionHandler.xlfForecast) return 2;
			if (iftab == FunctionHandler.xlfTrend) return 2;
			if (iftab == FunctionHandler.xlfIsnumber) return 1;
			return 1; //if we are lucky
		}
	*/
	@Override
	public boolean getIsUnaryOperator()
	{
		return true;
	}

	@Override
	public String getString()
	{
		short iftb = (short) iftab;
		String f = null;
		if( Locale.JAPAN.equals( Locale.getDefault() ) )
		{
			f = FunctionConstants.getJFunctionString( iftb );
		}
		if( f == null )
		{
			f = FunctionConstants.getFunctionString( iftb );
		}
		return f;
	}

	@Override
	public String getString2()
	{
		return ")";
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	private void populateVals()
	{
		iftab = ByteTools.readShort( record[1], record[2] );
	}

	public int getVal()
	{
		return iftab;
	}

	public void setVal( int i )
	{
		iftab = i;
		this.updateRecord();
	}

	@Override
	public void updateRecord()
	{
		byte[] tmp = new byte[1];
		tmp[0] = record[0];
		byte[] brow = ByteTools.cLongToLEBytes( iftab );
		tmp = ByteTools.append( brow, tmp );
		record = tmp;
	}

	@Override
	public int getLength()
	{
		return PTG_FUNC_LENGTH;
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
	 * given this specific Func, ensure that it's parameters are of the correct Ptg type
	 * <br>Value, Reference or Array
	 * <br>This is necessary when functions are added via String
	 * <br>NOTE: This method is a stub; eventually ALL Functions which require
	 * specific types of parameters will be handled here
	 *
	 * @see FormulaParser.adjustParameterIds
	 */
	public void adjustParameterIds()
	{
		if( vars == null )
		{
			return; // no parameters to worry about
		}
		/*TODO Eventually will have a list of Function Id's which require a certain type of parameter*/
		switch( iftab )
		{
			case FunctionConstants.xlfRows:
				for( int i = 0; i < vars.length; i++ )
				{
					if( vars[0] instanceof PtgRef )
					{
						((PtgRef) vars[0]).setPtgType( PtgRef.REFERENCE );
					}
					else if( vars[0] instanceof PtgName )
					{
						((PtgName) vars[0]).setPtgType( PtgRef.REFERENCE );
					}
					break;
				}
		}
	}

}