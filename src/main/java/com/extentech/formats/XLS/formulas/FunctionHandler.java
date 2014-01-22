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

import java.util.Locale;

/**
 * Function Handler takes an array of PTG's with a header PtgFunc or PtgFuncVar, calcuates
 * those ptgs in the determined way, then return a relevant PtgValue
 * <p/>
 * Descriptions of these functions are available on the msdn site,
 * http://msdn.microsoft.com/library/default.asp?url=/library/en-us/office97/html/S88F9.asp
 */

public class FunctionHandler
{

	/*
		Calculates the function and returns a relevant Ptg as a value
		This is the main entry method.

	*/
	public static final Ptg calculateFunction( Ptg[] ptgs ) throws FunctionNotSupportedException, CalculationException
	{
		Ptg funk; // the function identifier
		Ptg[] operands; // the ptgs acted upon by the function
		int funkId = 0;   //what function are we calling?

		funk = ptgs[0];
		// if ptgs are missing parent_recs, populate from funk
		XLSRecord bpar = funk.getParentRec();
		if( bpar != null )
		{
			for( int t = 0; t < ptgs.length; t++ )
			{
				if( ptgs[t].getParentRec() == null )
				{
					ptgs[t].setParentRec( bpar );
				}
			}
		}

		int oplen = ptgs.length - 1;
		operands = new Ptg[oplen];
		System.arraycopy( ptgs, 1, operands, 0, oplen );
		if( (funk.getOpcode() == 0x21) || (funk.getOpcode() == 0x41) || (funk.getOpcode() == 0x61) )
		{  // ptgfunc
			return calculatePtgFunc( funk, funkId, operands );
		}
		if( (funk.getOpcode() == 0x22) || (funk.getOpcode() == 0x42) || (funk.getOpcode() == 0x62) )
		{ // ptgfuncvar
			return calculatePtgFuncVar( funk, funkId, operands );
		}
		return null;
	}

	/*
		Keep the calculation of ptgfunc & ptgfuncvar seperate in case any differences show up
	*/
	public static final Ptg calculatePtgFunc( Ptg funk, int funkId, Ptg[] operands ) throws
	                                                                                 FunctionNotSupportedException,
	                                                                                 CalculationException
	{
		PtgFunc pf = (PtgFunc) funk;
		funkId = pf.getVal();
		return parse_n_calc( funk, funkId, operands );
	}

	/**
	 * Keep the calculation of ptgfunc & ptgfuncvar seperate in case any differences show up
	 *
	 * @throws CalculationException
	 */
	public static final Ptg calculatePtgFuncVar( Ptg funk, int funkId, Ptg[] operands ) throws
	                                                                                    FunctionNotSupportedException,
	                                                                                    CalculationException
	{
		PtgFuncVar pf = (PtgFuncVar) funk;
		funkId = pf.getVal();
		// Handle Add-in Formulas - which have a name operand 1st
		if( funkId == FunctionConstants.xlfADDIN )
		{  // XL flag that formula is an add-in
			//	must pop the PtgNameX record to get the correct function id
			String s = "";
			boolean foundit = false;
			if( operands[0] instanceof PtgNameX )
			{
				int index = ((PtgNameX) operands[0]).getVal();
				s = pf.getParentRec().getSheet().getWorkBook().getExternalName( index );
			}
			else if( operands[0] instanceof PtgName )
			{
				s = ((PtgName) operands[0]).getStoredName();
			}
			if( s.startsWith( "_xlfn." ) )
			{    // Excel "new" functions
				s = s.substring( 6 );
			}
			if( Locale.JAPAN.equals( Locale.getDefault() ) )
			{
				for( int y = 0; y < FunctionConstants.jRecArr.length; y++ )
				{
					if( s.equalsIgnoreCase( FunctionConstants.jRecArr[y][0] ) )
					{
						funkId = Integer.valueOf( FunctionConstants.jRecArr[y][1] );
						y = FunctionConstants.jRecArr.length;  // exit loop
						foundit = true;
					}
				}
			}
			if( !foundit )
			{
				for( int y = 0; y < FunctionConstants.recArr.length; y++ )
				{    // Use FunctionConstants instead of PtFuncVar
					if( s.equalsIgnoreCase( FunctionConstants.recArr[y][0] ) )
					{
						funkId = Integer.valueOf( FunctionConstants.recArr[y][1] );
						y = FunctionConstants.recArr.length;    // exit loop
					}
				}
			}
			if( funkId == 255 )// it's not found
			{
				throw new FunctionNotSupportedException( s );
			}

//			now get rid of PtgNameX operand before calling function
			Ptg[] ops = new Ptg[operands.length - 1];
			System.arraycopy( operands, 1, ops, 0, operands.length - 1 );
			operands = new Ptg[ops.length];
			System.arraycopy( ops, 0, operands, 0, ops.length );
		}    // end KSC added
		return parse_n_calc( funk, funkId, operands );
	}

	/**
	 * *********************************************************************************
	 * *
	 * Your standard big case statement, calling methods based on what the funkid is. *
	 * You will notice that these are seperated out into packages based on the MS     *
	 * documentation (Link above).  Each package calls a different class full of      *
	 * static method calls.   There are a lot of these :-)                            *
	 * *
	 * PLEASE:  Remove function from comment list when you enable it!!!!           *
	 * *
	 *
	 * @throws CalculationException ***********************************************************************************
	 */
	public static final Ptg parse_n_calc( Ptg function, int functionId, Ptg[] operands ) throws
	                                                                                     FunctionNotSupportedException,
	                                                                                     CalculationException
	{
		Ptg resultPtg = null;
		Ptg[] resultArrPtg = null;

		switch( functionId )
		{
			/********************************************
			 *   Database and List package functions    **
			 ********************************************/
			case FunctionConstants.xlfDaverage:
				resultPtg = DatabaseCalculator.calcDAverage( operands );
				break;

			case FunctionConstants.xlfDcount:
				resultPtg = DatabaseCalculator.calcDCount( operands );
				break;

			case FunctionConstants.xlfDcounta:
				resultPtg = DatabaseCalculator.calcDCountA( operands );
				break;

			case FunctionConstants.xlfDget:
				resultPtg = DatabaseCalculator.calcDGet( operands );
				break;

			case FunctionConstants.xlfDmax:
				resultPtg = DatabaseCalculator.calcDMax( operands );
				break;

			case FunctionConstants.xlfDmin:
				resultPtg = DatabaseCalculator.calcDMin( operands );
				break;

			case FunctionConstants.xlfDproduct:
				resultPtg = DatabaseCalculator.calcDProduct( operands );
				break;

			case FunctionConstants.xlfDstdev:
				resultPtg = DatabaseCalculator.calcDStdDev( operands );
				break;

			case FunctionConstants.xlfDstdevp:
				resultPtg = DatabaseCalculator.calcDStdDevP( operands );
				break;

			case FunctionConstants.xlfDsum:
				resultPtg = DatabaseCalculator.calcDSum( operands );
				break;

			case FunctionConstants.xlfDvar:
				resultPtg = DatabaseCalculator.calcDVar( operands );
				break;

			case FunctionConstants.xlfDvarp:
				resultPtg = DatabaseCalculator.calcDVarP( operands );
				break;

			/********************************************
			 *   Date and time functions         *********
			 ********************************************/
			case FunctionConstants.xlfDate:
				resultPtg = DateTimeCalculator.calcDate( operands );
				break;

			case FunctionConstants.xlfDay:
				resultPtg = DateTimeCalculator.calcDay( operands );
				break;

			case FunctionConstants.xlfDays360:
				resultPtg = DateTimeCalculator.calcDays360( operands );
				break;

			case FunctionConstants.xlfHour:
				resultPtg = DateTimeCalculator.calcHour( operands );
				break;

			case FunctionConstants.xlfMinute:
				resultPtg = DateTimeCalculator.calcMinute( operands );
				break;

			case FunctionConstants.xlfMonth:
				resultPtg = DateTimeCalculator.calcMonth( operands );
				break;

			case FunctionConstants.xlfYear:
				resultPtg = DateTimeCalculator.calcYear( operands );
				break;

			case FunctionConstants.xlfSecond:
				resultPtg = DateTimeCalculator.calcSecond( operands );
				break;

			case FunctionConstants.xlfTimevalue:
				resultPtg = DateTimeCalculator.calcTimevalue( operands );
				break;

			case FunctionConstants.xlfWeekday:
				resultPtg = DateTimeCalculator.calcWeekday( operands );
				break;

			case FunctionConstants.xlfWEEKNUM:
				resultPtg = DateTimeCalculator.calcWeeknum( operands );
				break;

			case FunctionConstants.xlfWORKDAY:
				resultPtg = DateTimeCalculator.calcWorkday( operands );
				break;

			case FunctionConstants.xlfYEARFRAC:
				resultPtg = DateTimeCalculator.calcYearFrac( operands );
				break;

			case FunctionConstants.xlfNow:
				resultPtg = DateTimeCalculator.calcNow( operands );
				break;

			case FunctionConstants.xlfTime:
				resultPtg = DateTimeCalculator.calcTime( operands );
				break;

			case FunctionConstants.xlfToday:
				resultPtg = DateTimeCalculator.calcToday( operands );
				break;

			case FunctionConstants.xlfDatevalue:
				resultPtg = DateTimeCalculator.calcDateValue( operands );
				break;

			case FunctionConstants.xlfEDATE:
				resultPtg = DateTimeCalculator.calcEdate( operands );
				break;

			case FunctionConstants.xlfEOMONTH:
				resultPtg = DateTimeCalculator.calcEOMonth( operands );
				break;

			case FunctionConstants.xlfNETWORKDAYS:
				resultPtg = DateTimeCalculator.calcNetWorkdays( operands );
				break;

			/********************************************
			 *   DDE and External functions           ****
			 ********************************************/

			/********************************************
			 *   Engineering functions           *********
			 ********************************************/
			case FunctionConstants.xlfBIN2DEC:
				resultPtg = EngineeringCalculator.calcBin2Dec( operands );
				break;

			case FunctionConstants.xlfBIN2HEX:
				resultPtg = EngineeringCalculator.calcBin2Hex( operands );
				break;

			case FunctionConstants.xlfBIN2OCT:
				resultPtg = EngineeringCalculator.calcBin2Oct( operands );
				break;

			case FunctionConstants.xlfDEC2BIN:
				resultPtg = EngineeringCalculator.calcDec2Bin( operands );
				break;

			case FunctionConstants.xlfDEC2HEX:
				resultPtg = EngineeringCalculator.calcDec2Hex( operands );
				break;

			case FunctionConstants.xlfDEC2OCT:
				resultPtg = EngineeringCalculator.calcDec2Oct( operands );
				break;

			case FunctionConstants.xlfHEX2BIN:
				resultPtg = EngineeringCalculator.calcHex2Bin( operands );
				break;

			case FunctionConstants.xlfHEX2DEC:
				resultPtg = EngineeringCalculator.calcHex2Dec( operands );
				break;

			case FunctionConstants.xlfHEX2OCT:
				resultPtg = EngineeringCalculator.calcHex2Oct( operands );
				break;

			case FunctionConstants.xlfOCT2BIN:
				resultPtg = EngineeringCalculator.calcOct2Bin( operands );
				break;

			case FunctionConstants.xlfOCT2DEC:
				resultPtg = EngineeringCalculator.calcOct2Dec( operands );
				break;

			case FunctionConstants.xlfOCT2HEX:
				resultPtg = EngineeringCalculator.calcOct2Hex( operands );
				break;

			case FunctionConstants.xlfCOMPLEX:
				resultPtg = EngineeringCalculator.calcComplex( operands );
				break;

			case FunctionConstants.xlfGESTEP:
				resultPtg = EngineeringCalculator.calcGEStep( operands );
				break;

			case FunctionConstants.xlfDELTA:
				resultPtg = EngineeringCalculator.calcDelta( operands );
				break;

			case FunctionConstants.xlfIMAGINARY:
				resultPtg = EngineeringCalculator.calcImaginary( operands );
				break;

			case FunctionConstants.xlfIMREAL:
				resultPtg = EngineeringCalculator.calcImReal( operands );
				break;

			case FunctionConstants.xlfIMARGUMENT:
				resultPtg = EngineeringCalculator.calcImArgument( operands );
				break;

			case FunctionConstants.xlfIMABS:
				resultPtg = EngineeringCalculator.calcImAbs( operands );
				break;

			case FunctionConstants.xlfIMDIV:
				resultPtg = EngineeringCalculator.calcImDiv( operands );
				break;

			case FunctionConstants.xlfIMCONJUGATE:
				resultPtg = EngineeringCalculator.calcImConjugate( operands );
				break;

			case FunctionConstants.xlfIMCOS:
				resultPtg = EngineeringCalculator.calcImCos( operands );
				break;

			case FunctionConstants.xlfIMSIN:
				resultPtg = EngineeringCalculator.calcImSin( operands );
				break;

			case FunctionConstants.xlfIMEXP:
				resultPtg = EngineeringCalculator.calcImExp( operands );
				break;

			case FunctionConstants.xlfIMSUB:
				resultPtg = EngineeringCalculator.calcImSub( operands );
				break;

			case FunctionConstants.xlfIMSUM:
				resultPtg = EngineeringCalculator.calcImSum( operands );
				break;

			case FunctionConstants.xlfIMPRODUCT:
				resultPtg = EngineeringCalculator.calcImProduct( operands );
				break;

			case FunctionConstants.xlfIMLN:
				resultPtg = EngineeringCalculator.calcImLn( operands );
				break;

			case FunctionConstants.xlfIMLOG10:
				resultPtg = EngineeringCalculator.calcImLog10( operands );
				break;

			case FunctionConstants.xlfIMLOG2:
				resultPtg = EngineeringCalculator.calcImLog2( operands );
				break;

			case FunctionConstants.xlfIMPOWER:
				resultPtg = EngineeringCalculator.calcImPower( operands );
				break;

			case FunctionConstants.xlfIMSQRT:
				resultPtg = EngineeringCalculator.calcImSqrt( operands );
				break;

			case FunctionConstants.xlfCONVERT:
				resultPtg = EngineeringCalculator.calcConvert( operands );
				break;

			case FunctionConstants.xlfERF:
				resultPtg = EngineeringCalculator.calcErf( operands );
				break;

			/********************************************
			 *   Financial functions             *********
			 ********************************************/

			case FunctionConstants.xlfDb:
				resultPtg = FinancialCalculator.calcDB( operands );
				break;

			case FunctionConstants.xlfDdb:
				resultPtg = FinancialCalculator.calcDDB( operands );
				break;

			case FunctionConstants.xlfPmt:
				resultPtg = FinancialCalculator.calcPmt( operands );
				break;

			// KSC: Added
			case FunctionConstants.xlfAccrintm:
				resultPtg = FinancialCalculator.calcAccrintm( operands );
				break;

			case FunctionConstants.xlfAccrint:
				resultPtg = FinancialCalculator.calcAccrint( operands );
				break;

			case FunctionConstants.xlfCoupDayBS:
				resultPtg = FinancialCalculator.calcCoupDayBS( operands );
				break;

			case FunctionConstants.xlfCoupDays:
				resultPtg = FinancialCalculator.calcCoupDays( operands );
				break;

			case FunctionConstants.xlfNpv:
				resultPtg = FinancialCalculator.calcNPV( operands );
				break;

			case FunctionConstants.xlfPv:
				resultPtg = FinancialCalculator.calcPV( operands );
				break;

			case FunctionConstants.xlfFv:
				resultPtg = FinancialCalculator.calcFV( operands );
				break;

			case FunctionConstants.xlfIpmt:
				resultPtg = FinancialCalculator.calcIPMT( operands );
				break;

			case FunctionConstants.xlfCumIPmt:
				resultPtg = FinancialCalculator.calcCumIPmt( operands );
				break;

			case FunctionConstants.xlfCumPrinc:
				resultPtg = FinancialCalculator.calcCumPrinc( operands );
				break;

			case FunctionConstants.xlfCoupNCD:
				resultPtg = FinancialCalculator.calcCoupNCD( operands );
				break;

			case FunctionConstants.xlfCoupDaysNC:
				resultPtg = FinancialCalculator.calcCoupDaysNC( operands );
				break;

			case FunctionConstants.xlfCoupPCD:
				resultPtg = FinancialCalculator.calcCoupPCD( operands );
				break;

			case FunctionConstants.xlfCoupNUM:
				resultPtg = FinancialCalculator.calcCoupNum( operands );
				break;

			case FunctionConstants.xlfDollarDE:
				resultPtg = FinancialCalculator.calcDollarDE( operands );
				break;

			case FunctionConstants.xlfDollarFR:
				resultPtg = FinancialCalculator.calcDollarFR( operands );
				break;

			case FunctionConstants.xlfEffect:
				resultPtg = FinancialCalculator.calcEffect( operands );
				break;

			case FunctionConstants.xlfRECEIVED:
				resultPtg = FinancialCalculator.calcReceived( operands );
				break;

			case FunctionConstants.xlfINTRATE:
				resultPtg = FinancialCalculator.calcINTRATE( operands );
				break;

			case FunctionConstants.xlfIrr:
				resultPtg = FinancialCalculator.calcIRR( operands );
				break;

			case FunctionConstants.xlfMirr:
				resultPtg = FinancialCalculator.calcMIRR( operands );
				break;

			case FunctionConstants.xlfXIRR:
				resultPtg = FinancialCalculator.calcXIRR( operands );
				break;

			case FunctionConstants.xlfXNPV:
				resultPtg = FinancialCalculator.calcXNPV( operands );
				break;

			case FunctionConstants.xlfRate:
				resultPtg = FinancialCalculator.calcRate( operands );
				break;

			case FunctionConstants.xlfYIELD:
				resultPtg = FinancialCalculator.calcYIELD( operands );
				break;

			case FunctionConstants.xlfPRICE:
				resultPtg = FinancialCalculator.calcPRICE( operands );
				break;

			case FunctionConstants.xlfPRICEDISC:
				resultPtg = FinancialCalculator.calcPRICEDISC( operands );
				break;

			case FunctionConstants.xlfPRICEMAT:
				resultPtg = FinancialCalculator.calcPRICEMAT( operands );
				break;

			case FunctionConstants.xlfDISC:
				resultPtg = FinancialCalculator.calcDISC( operands );
				break;

			case FunctionConstants.xlfNper:
				resultPtg = FinancialCalculator.calcNPER( operands );
				break;

			case FunctionConstants.xlfSln:
				resultPtg = FinancialCalculator.calcSLN( operands );
				break;

			case FunctionConstants.xlfSyd:
				resultPtg = FinancialCalculator.calcSYD( operands );
				break;

			case FunctionConstants.xlfDURATION:
				resultPtg = FinancialCalculator.calcDURATION( operands );
				break;

			case FunctionConstants.xlfMDURATION:
				resultPtg = FinancialCalculator.calcMDURATION( operands );
				break;

			case FunctionConstants.xlfTBillEq:
				resultPtg = FinancialCalculator.calcTBillEq( operands );
				break;

			case FunctionConstants.xlfTBillPrice:
				resultPtg = FinancialCalculator.calcTBillPrice( operands );
				break;

			case FunctionConstants.xlfTBillYield:
				resultPtg = FinancialCalculator.calcTBillYield( operands );
				break;

			case FunctionConstants.xlfYieldDisc:
				resultPtg = FinancialCalculator.calcYieldDisc( operands );
				break;

			case FunctionConstants.xlfYieldMat:
				resultPtg = FinancialCalculator.calcYieldMat( operands );
				break;

			case FunctionConstants.xlfPpmt:
				resultPtg = FinancialCalculator.calcPPMT( operands );
				break;

			case FunctionConstants.xlfFVSchedule:
				resultPtg = FinancialCalculator.calcFVSCHEDULE( operands );
				break;

			case FunctionConstants.xlfIspmt:
				resultPtg = FinancialCalculator.calcISPMT( operands );
				break;

			case FunctionConstants.xlfAmorlinc:
				resultPtg = FinancialCalculator.calcAmorlinc( operands );
				break;

			case FunctionConstants.xlfAmordegrc:
				resultPtg = FinancialCalculator.calcAmordegrc( operands );
				break;

			case FunctionConstants.xlfOddFPrice:
				resultPtg = FinancialCalculator.calcODDFPRICE( operands );
				break;

			case FunctionConstants.xlfOddFYield:
				resultPtg = FinancialCalculator.calcODDFYIELD( operands );
				break;

			case FunctionConstants.xlfOddLPrice:
				resultPtg = FinancialCalculator.calcODDLPRICE( operands );
				break;

			case FunctionConstants.xlfOddLYield:
				resultPtg = FinancialCalculator.calcODDLYIELD( operands );
				break;

			case FunctionConstants.xlfNOMINAL:
				resultPtg = FinancialCalculator.calcNominal( operands );
				break;

			case FunctionConstants.xlfVdb:
				resultPtg = FinancialCalculator.calcVDB( operands );
				break;
			/********************************************
			 *   Information functions           *********
			 ********************************************/

			case FunctionConstants.xlfCell:
				resultPtg = InformationCalculator.calcCell( operands );
				break;

			case FunctionConstants.xlfInfo:
				resultPtg = InformationCalculator.calcInfo( operands );
				break;

			case FunctionConstants.XLF_IS_NA:
				resultPtg = InformationCalculator.calcIsna( operands );
				break;

			case FunctionConstants.XLF_IS_ERROR:
				resultPtg = InformationCalculator.calcIserror( operands );
				break;

			case FunctionConstants.xlfIserr:
				resultPtg = InformationCalculator.calcIserr( operands );
				break;

			case FunctionConstants.xlfErrorType:
				resultPtg = InformationCalculator.calcErrorType( operands );
				break;

			case FunctionConstants.xlfNa:
				resultPtg = InformationCalculator.calcNa( operands );
				break;

			case FunctionConstants.xlfIsblank:
				resultPtg = InformationCalculator.calcIsBlank( operands );
				break;

			case FunctionConstants.xlfIslogical:
				resultPtg = InformationCalculator.calcIsLogical( operands );
				break;

			case FunctionConstants.xlfIsnontext:
				resultPtg = InformationCalculator.calcIsNonText( operands );
				break;

			case FunctionConstants.xlfIstext:
				resultPtg = InformationCalculator.calcIsText( operands );
				break;

			case FunctionConstants.xlfIsref:
				resultPtg = InformationCalculator.calcIsRef( operands );
				break;

			case FunctionConstants.xlfN:
				resultPtg = InformationCalculator.calcN( operands );
				break;

			case FunctionConstants.xlfIsnumber:
				resultPtg = InformationCalculator.calcIsNumber( operands );
				break;

			case FunctionConstants.xlfISEVEN:
				resultPtg = InformationCalculator.calcIsEven( operands );
				break;

			case FunctionConstants.xlfISODD:
				resultPtg = InformationCalculator.calcIsOdd( operands );
				break;

			case FunctionConstants.xlfType:
				resultPtg = InformationCalculator.calcType( operands );
				break;

			/********************************************
			 *   Logical functions       *****
			 ********************************************/
			case FunctionConstants.xlfAnd:
				resultPtg = LogicalCalculator.calcAnd( operands );
				break;

			case FunctionConstants.xlfFalse:
				resultPtg = LogicalCalculator.calcFalse( operands );
				break;

			case FunctionConstants.xlfTrue:
				resultPtg = LogicalCalculator.calcTrue( operands );
				break;

			case FunctionConstants.XLF_IS:
				resultPtg = LogicalCalculator.calcIf( operands );
				break;

			case FunctionConstants.xlfNot:
				resultPtg = LogicalCalculator.calcNot( operands );
				break;

			case FunctionConstants.xlfOr:
				resultPtg = LogicalCalculator.calcOr( operands );
				break;

			case FunctionConstants.xlfIFERROR:
				resultPtg = LogicalCalculator.calcIferror( operands );
				break;

			/********************************************
			 *   Lookup and reference functions       *****
			 ********************************************/
			case FunctionConstants.xlfAddress:
				resultPtg = LookupReferenceCalculator.calcAddress( operands );
				break;

			case FunctionConstants.xlfAreas:
				resultPtg = LookupReferenceCalculator.calcAreas( operands );
				break;

			case FunctionConstants.xlfChoose:
				resultPtg = LookupReferenceCalculator.calcChoose( operands );
				break;

			case FunctionConstants.xlfColumn:
				if( operands.length == 0 )
				{
					operands = new Ptg[1];
					operands[0] = function;
				}
				resultPtg = LookupReferenceCalculator.calcColumn( operands );
				break;

			case FunctionConstants.xlfColumns:
				resultPtg = LookupReferenceCalculator.calcColumns( operands );
				break;

			case FunctionConstants.xlfHyperlink:
				resultPtg = LookupReferenceCalculator.calcHyperlink( operands );
				break;

			case FunctionConstants.xlfIndex:
				resultPtg = LookupReferenceCalculator.calcIndex( operands );
				break;

			case FunctionConstants.XLF_INDIRECT:
				resultPtg = LookupReferenceCalculator.calcIndirect( operands );
				break;

			case FunctionConstants.XLF_ROW:
				if( operands.length == 0 )
				{
					operands = new Ptg[1];
					operands[0] = function;
				}
				resultPtg = LookupReferenceCalculator.calcRow( operands );
				break;

			case FunctionConstants.xlfRows:
				resultPtg = LookupReferenceCalculator.calcRows( operands );
				break;

			case FunctionConstants.xlfTranspose:
				resultPtg = LookupReferenceCalculator.calcTranspose( operands );
				break;

			case FunctionConstants.xlfLookup:
				resultPtg = LookupReferenceCalculator.calcLookup( operands );
				// KSC: Clear out lookup caches!
//			function.getParentRec().getWorkBook().getRefTracker().clearLookupCaches();
				break;

			case FunctionConstants.xlfHlookup:
				resultPtg = LookupReferenceCalculator.calcHlookup( operands );
				// KSC: Clear out lookup caches!
//			function.getParentRec().getWorkBook().getRefTracker().clearLookupCaches();
				break;

			case FunctionConstants.xlfVlookup:
				resultPtg = LookupReferenceCalculator.calcVlookup( operands );
				break;

			case FunctionConstants.xlfMatch:
				resultPtg = LookupReferenceCalculator.calcMatch( operands );
				break;

			case FunctionConstants.xlfOffset:
				resultPtg = LookupReferenceCalculator.calcOffset( operands );
				break;

			/********************************************
			 *   Math & Trigonometry functions       *****
			 ********************************************/

			case FunctionConstants.XLF_SUM:
				resultPtg = MathFunctionCalculator.calcSum( operands );
				break;

			case FunctionConstants.XLF_SUM_IF:
				resultPtg = MathFunctionCalculator.calcSumif( operands );
				break;

			case FunctionConstants.xlfSUMIFS:
				resultPtg = MathFunctionCalculator.calcSumIfS( operands );
				break;

			case FunctionConstants.xlfSumproduct:
				resultPtg = MathFunctionCalculator.calcSumproduct( operands );
				break;

			case FunctionConstants.xlfExp:
				resultPtg = MathFunctionCalculator.calcExp( operands );
				break;

			case FunctionConstants.xlfAbs:
				resultPtg = MathFunctionCalculator.calcAbs( operands );
				break;

			case FunctionConstants.xlfAcos:
				resultPtg = MathFunctionCalculator.calcAcos( operands );
				break;

			case FunctionConstants.xlfAcosh:
				resultPtg = MathFunctionCalculator.calcAcosh( operands );
				break;

			case FunctionConstants.xlfAsin:
				resultPtg = MathFunctionCalculator.calcAsin( operands );
				break;

			case FunctionConstants.xlfAsinh:
				resultPtg = MathFunctionCalculator.calcAsinh( operands );
				break;

			case FunctionConstants.xlfAtan:
				resultPtg = MathFunctionCalculator.calcAtan( operands );
				break;

			case FunctionConstants.xlfAtan2:
				resultPtg = MathFunctionCalculator.calcAtan2( operands );
				break;

			case FunctionConstants.xlfAtanh:
				resultPtg = MathFunctionCalculator.calcAtanh( operands );
				break;

			case FunctionConstants.xlfCeiling:
				resultPtg = MathFunctionCalculator.calcCeiling( operands );
				break;

			case FunctionConstants.xlfCombin:
				resultPtg = MathFunctionCalculator.calcCombin( operands );
				break;

			case FunctionConstants.xlfCos:
				resultPtg = MathFunctionCalculator.calcCos( operands );
				break;

			case FunctionConstants.xlfCosh:
				resultPtg = MathFunctionCalculator.calcCosh( operands );
				break;

			case FunctionConstants.xlfDegrees:
				resultPtg = MathFunctionCalculator.calcDegrees( operands );
				break;

			case FunctionConstants.xlfEven:
				resultPtg = MathFunctionCalculator.calcEven( operands );
				break;

			case FunctionConstants.xlfFact:
				resultPtg = MathFunctionCalculator.calcFact( operands );
				break;

			case FunctionConstants.xlfDOUBLEFACT:
				resultPtg = MathFunctionCalculator.calcFactDouble( operands );
				break;

			case FunctionConstants.xlfFloor:
				resultPtg = MathFunctionCalculator.calcFloor( operands );
				break;

			case FunctionConstants.xlfGCD:
				resultPtg = MathFunctionCalculator.calcGCD( operands );
				break;

			case FunctionConstants.xlfInt:
				resultPtg = MathFunctionCalculator.calcInt( operands );
				break;

			case FunctionConstants.xlfLCM:
				resultPtg = MathFunctionCalculator.calcLCM( operands );
				break;

			case FunctionConstants.xlfMROUND:
				resultPtg = MathFunctionCalculator.calcMRound( operands );
				break;

			case FunctionConstants.xlfMmult:
				resultPtg = MathFunctionCalculator.calcMMult( operands );
				break;

			case FunctionConstants.xlfMULTINOMIAL:
				resultPtg = MathFunctionCalculator.calcMultinomial( operands );
				break;

			case FunctionConstants.xlfLn:
				resultPtg = MathFunctionCalculator.calcLn( operands );
				break;

			case FunctionConstants.xlfLog:
				resultPtg = MathFunctionCalculator.calcLog( operands );
				break;

			case FunctionConstants.xlfLog10:
				resultPtg = MathFunctionCalculator.calcLog10( operands );
				break;

			case FunctionConstants.xlfMod:
				resultPtg = MathFunctionCalculator.calcMod( operands );
				break;

			case FunctionConstants.xlfOdd:
				resultPtg = MathFunctionCalculator.calcOdd( operands );
				break;

			case FunctionConstants.xlfPi:
				resultPtg = MathFunctionCalculator.calcPi( operands );
				break;

			case FunctionConstants.xlfPower:
				resultPtg = MathFunctionCalculator.calcPower( operands );
				break;

			case FunctionConstants.xlfProduct:
				resultPtg = MathFunctionCalculator.calcProduct( operands );
				break;

			case FunctionConstants.xlfQUOTIENT:
				resultPtg = MathFunctionCalculator.calcQuotient( operands );
				break;

			case FunctionConstants.xlfRadians:
				resultPtg = MathFunctionCalculator.calcRadians( operands );
				break;

			case FunctionConstants.xlfRand:
				resultPtg = MathFunctionCalculator.calcRand( operands );
				break;

			case FunctionConstants.xlfRANDBETWEEN:
				resultPtg = MathFunctionCalculator.calcRandBetween( operands );
				break;

			case FunctionConstants.xlfRoman:
				resultPtg = MathFunctionCalculator.calcRoman( operands );
				break;

			case FunctionConstants.xlfRound:
				resultPtg = MathFunctionCalculator.calcRound( operands );
				break;

			case FunctionConstants.xlfRounddown:
				resultPtg = MathFunctionCalculator.calcRoundDown( operands );
				break;

			case FunctionConstants.xlfRoundup:
				resultPtg = MathFunctionCalculator.calcRoundUp( operands );
				break;

			case FunctionConstants.xlfSign:
				resultPtg = MathFunctionCalculator.calcSign( operands );
				break;

			case FunctionConstants.xlfSin:
				resultPtg = MathFunctionCalculator.calcSin( operands );
				break;

			case FunctionConstants.xlfSinh:
				resultPtg = MathFunctionCalculator.calcSinh( operands );
				break;

			case FunctionConstants.xlfSqrt:
				resultPtg = MathFunctionCalculator.calcSqrt( operands );
				break;

			case FunctionConstants.xlfSQRTPI:
				resultPtg = MathFunctionCalculator.calcSqrtPi( operands );
				break;

			case FunctionConstants.xlfTan:
				resultPtg = MathFunctionCalculator.calcTan( operands );
				break;

			case FunctionConstants.xlfTanh:
				resultPtg = MathFunctionCalculator.calcTanh( operands );
				break;

			case FunctionConstants.xlfTrunc:
				resultPtg = MathFunctionCalculator.calcTrunc( operands );
				break;

			/********************************************
			 *   Statistical functions       *****
			 ********************************************/

			case FunctionConstants.XLF_COUNT:
				resultPtg = StatisticalCalculator.calcCount( operands );
				break;

			case FunctionConstants.xlfCounta:
				resultPtg = StatisticalCalculator.calcCountA( operands );
				break;

			case FunctionConstants.xlfCountblank:
				resultPtg = StatisticalCalculator.calcCountBlank( operands );
				break;

			case FunctionConstants.xlfCountif:
				resultPtg = StatisticalCalculator.calcCountif( operands );
				break;

			case FunctionConstants.xlfCOUNTIFS:
				resultPtg = StatisticalCalculator.calcCountIfS( operands );
				break;

			case FunctionConstants.XLF_MIN:
				resultPtg = StatisticalCalculator.calcMin( operands );
				break;

			case FunctionConstants.xlfMinA:
				resultPtg = StatisticalCalculator.calcMinA( operands );
				break;

			case FunctionConstants.XLF_MAX:
				resultPtg = StatisticalCalculator.calcMax( operands );
				break;

			case FunctionConstants.xlfMaxA:
				resultPtg = StatisticalCalculator.calcMaxA( operands );
				break;

			case FunctionConstants.xlfNormdist:
				resultPtg = StatisticalCalculator.calcNormdist( operands );
				break;

			case FunctionConstants.xlfNormsdist:
				resultPtg = StatisticalCalculator.calcNormsdist( operands );
				break;

			case FunctionConstants.xlfNormsinv:
				resultPtg = StatisticalCalculator.calcNormsInv( operands );
				break;

			case FunctionConstants.xlfNorminv:
				resultPtg = StatisticalCalculator.calcNormInv( operands );
				break;

			case FunctionConstants.XLF_AVERAGE:
				resultPtg = StatisticalCalculator.calcAverage( operands );
				break;

			case FunctionConstants.xlfAVERAGEIF:
				resultPtg = StatisticalCalculator.calcAverageIf( operands );
				break;

			case FunctionConstants.xlfAVERAGEIFS:
				resultPtg = StatisticalCalculator.calcAverageIfS( operands );
				break;

			case FunctionConstants.xlfAvedev:
				resultPtg = StatisticalCalculator.calcAveDev( operands );
				break;

			case FunctionConstants.xlfAverageA:
				resultPtg = StatisticalCalculator.calcAverageA( operands );
				break;

			case FunctionConstants.xlfMedian:
				resultPtg = StatisticalCalculator.calcMedian( operands );
				break;

			case FunctionConstants.xlfMode:
				resultPtg = StatisticalCalculator.calcMode( operands );
				break;

			case FunctionConstants.xlfQuartile:
				resultPtg = StatisticalCalculator.calcQuartile( operands );
				break;

			case FunctionConstants.xlfRank:
				resultPtg = StatisticalCalculator.calcRank( operands );
				break;

			case FunctionConstants.xlfStdev:
				resultPtg = StatisticalCalculator.calcStdev( operands );
				break;

			case FunctionConstants.xlfVar:
				resultPtg = StatisticalCalculator.calcVar( operands );
				break;

			case FunctionConstants.xlfVarp:
				resultPtg = StatisticalCalculator.calcVarp( operands );
				break;

			case FunctionConstants.xlfCovar:
				resultPtg = StatisticalCalculator.calcCovar( operands );
				break;

			case FunctionConstants.xlfCorrel:
				resultPtg = StatisticalCalculator.calcCorrel( operands );
				break;

			case FunctionConstants.xlfFrequency:
				resultPtg = StatisticalCalculator.calcFrequency( operands );
				break;

			case FunctionConstants.xlfLinest:
				resultPtg = StatisticalCalculator.calcLineSt( operands );
				break;

			case FunctionConstants.xlfSlope:
				resultPtg = StatisticalCalculator.calcSlope( operands );
				break;

			case FunctionConstants.xlfIntercept:
				resultPtg = StatisticalCalculator.calcIntercept( operands );
				break;

			case FunctionConstants.xlfPearson:
				resultPtg = StatisticalCalculator.calcPearson( operands );
				break;

			case FunctionConstants.xlfRsq:
				resultPtg = StatisticalCalculator.calcRsq( operands );
				break;

			case FunctionConstants.xlfSteyx:
				resultPtg = StatisticalCalculator.calcSteyx( operands );
				break;

			case FunctionConstants.xlfForecast:
				resultPtg = StatisticalCalculator.calcForecast( operands );
				break;

			case FunctionConstants.xlfTrend:
				resultPtg = StatisticalCalculator.calcTrend( operands );
				break;

			case FunctionConstants.xlfLarge:
				resultPtg = StatisticalCalculator.calcLarge( operands );
				break;

			case FunctionConstants.xlfSmall:
				resultPtg = StatisticalCalculator.calcSmall( operands );
				break;

			/********************************************
			 *   Text functions                                *****
			 ********************************************/
			/*
			 * these DBCS functions are not working yet
		case FunctionConstants.xlfAsc:
			resultPtg= TextCalculator.calcAsc(operands);
			break;
			
		case FunctionConstants.xlfDbcs:
			resultPtg= TextCalculator.calcJIS(operands);
			break;
			*/
			case FunctionConstants.xlfChar:
				resultPtg = TextCalculator.calcChar( operands );
				break;

			case FunctionConstants.xlfClean:
				resultPtg = TextCalculator.calcClean( operands );
				break;

			case FunctionConstants.xlfCode:
				resultPtg = TextCalculator.calcCode( operands );
				break;

			case FunctionConstants.xlfConcatenate:
				resultPtg = TextCalculator.calcConcatenate( operands );
				break;

			case FunctionConstants.xlfDollar:
				resultPtg = TextCalculator.calcDollar( operands );
				break;

			case FunctionConstants.xlfExact:
				resultPtg = TextCalculator.calcExact( operands );
				break;

			case FunctionConstants.xlfFind:
				resultPtg = TextCalculator.calcFind( operands );
				break;

			// DBCS functions are not working 100% yet
			case FunctionConstants.xlfFindb:
				resultPtg = TextCalculator.calcFindB( operands );
				break;

			case FunctionConstants.xlfFixed:
				resultPtg = TextCalculator.calcFixed( operands );
				break;

			case FunctionConstants.xlfLeft:
				resultPtg = TextCalculator.calcLeft( operands );
				break;

			case FunctionConstants.xlfLeftb:
				resultPtg = TextCalculator.calcLeftB( operands );
				break;

			case FunctionConstants.xlfLen:
				resultPtg = TextCalculator.calcLen( operands );
				break;

			case FunctionConstants.xlfLenb:
				resultPtg = TextCalculator.calcLenB( operands );
				break;

			case FunctionConstants.xlfLower:
				resultPtg = TextCalculator.calcLower( operands );
				break;

			case FunctionConstants.xlfUpper:
				resultPtg = TextCalculator.calcUpper( operands );
				break;

			case FunctionConstants.xlfMid:
				resultPtg = TextCalculator.calcMid( operands );
				break;

			case FunctionConstants.xlfProper:
				resultPtg = TextCalculator.calcProper( operands );
				break;

			case FunctionConstants.xlfReplace:
				resultPtg = TextCalculator.calcReplace( operands );
				break;

			case FunctionConstants.xlfRept:
				resultPtg = TextCalculator.calcRept( operands );
				break;

			case FunctionConstants.xlfRight:
				resultPtg = TextCalculator.calcRight( operands );
				break;

			case FunctionConstants.xlfSearch:
				resultPtg = TextCalculator.calcSearch( operands );
				break;

			case FunctionConstants.xlfSearchb:
				resultPtg = TextCalculator.calcSearchB( operands );
				break;

			case FunctionConstants.xlfSubstitute:
				resultPtg = TextCalculator.calcSubstitute( operands );
				break;

			case FunctionConstants.xlfT:
				resultPtg = TextCalculator.calcT( operands );
				break;

			case FunctionConstants.xlfTrim:
				resultPtg = TextCalculator.calcTrim( operands );
				break;

			case FunctionConstants.xlfText:
				resultPtg = TextCalculator.calcText( operands );
				break;

			case FunctionConstants.xlfValue:
				resultPtg = TextCalculator.calcValue( operands );
				break;

			default:
				String s = FunctionConstants.getFunctionString( (short) functionId );
				if( (s != null) && !s.equals( "" ) )
				{
					s = s.substring( 0, s.length() - 1 );
				}
				else
				{
					s = new String( Integer.toHexString( (int) functionId ) );
				}
				//throw new FunctionNotSupportedException( (!.equals(""))?FunctionConstants.getFunctionString((short)funkId).substring(0, ):Integer.toHexString((int)funkId));
				throw new FunctionNotSupportedException( s );    // 20081118 KSC: add a little more info ...

		}

		return resultPtg;
	}

	/****************************************************************************
	 *                                                                           *
	 *   The following section is made up of the calcuations for each of the     *
	 *   function types.  These map directly to the name of the function         *
	 *   declared in the header and is called from the coresponding switch       *
	 *   statement above.  ENJOY!                                                *
	 *                                                                           *
	 *****************************************************************************/

	/****************************************
	 *                                       *
	 *   Excel function numbers              *
	 *                                       *
	 ****************************************/
/*	
	public static final int xlfCount    = 0;
	public static final int xlfIf	    = 1;
	public static final int xlfIsna     = 2;
	public static final int xlfIserror  = 3;
	public static final int xlfSum      = 4;
	public static final int xlfAverage  = 5;
	public static final int xlfMin      = 6;
	public static final int xlfMax      = 7;
	public static final int xlfRow      = 8;
	public static final int xlfColumn   = 9;
	public static final int xlfNa       = 10;
	public static final int xlfNpv      = 11;
	public static final int xlfStdev    = 12;
	public static final int xlfDollar   = 13;
	public static final int xlfFixed    = 14;
	public static final int xlfSin      = 15;
	public static final int xlfCos      = 16;
	public static final int xlfTan      = 17;
	public static final int xlfAtan     = 18;
	public static final int xlfPi       = 19;
	public static final int xlfSqrt     = 20;
	public static final int xlfExp      = 21;
	public static final int xlfLn       = 22;
	public static final int xlfLog10    = 23;
	public static final int xlfAbs      = 24;
	public static final int xlfInt      = 25;
	public static final int xlfSign     = 26;
	public static final int xlfRound    = 27;
	public static final int xlfLookup   = 28;
	public static final int xlfIndex    = 29;
	public static final int xlfRept     = 30;
	public static final int xlfMid      = 31;
	public static final int xlfLen      = 32;
	public static final int xlfValue    = 33;
	public static final int xlfTrue     = 34;
	public static final int xlfFalse    = 35;
	public static final int xlfAnd      = 36;
	public static final int xlfOr       = 37;
	public static final int xlfNot      = 38;
	public static final int xlfMod      = 39;
	public static final int xlfDcount   = 40;
	public static final int xlfDsum     = 41;
	public static final int xlfDaverage = 42;
	public static final int xlfDmin     = 43;
	public static final int xlfDmax     = 44;
	public static final int xlfDstdev   = 45;
	public static final int xlfVar      = 46;
	public static final int xlfDvar     = 47;
	public static final int xlfText     = 48;
	public static final int xlfLinest   = 49;
	public static final int xlfTrend    = 50;
	public static final int xlfLogest   = 51;
	public static final int xlfGrowth   = 52;
	public static final int xlfGoto     = 53;
	public static final int xlfHalt     = 54;
	public static final int xlfPv       = 56;
	public static final int xlfFv       = 57;
	public static final int xlfNper     = 58;
	public static final int xlfPmt      = 59;
	public static final int xlfRate     = 60;
	public static final int xlfMirr     = 61;
	public static final int xlfIrr      = 62;
	public static final int xlfRand     = 63;
	public static final int xlfMatch    = 64;
	public static final int xlfDate     = 65;
	public static final int xlfTime     = 66;
	public static final int xlfDay      = 67;
	public static final int xlfMonth    = 68;
	public static final int xlfYear     = 69;
	public static final int xlfWeekday  = 70;
	public static final int xlfHour     = 71;
	public static final int xlfMinute   = 72;
	public static final int xlfSecond   = 73;
	public static final int xlfNow      = 74;
	public static final int xlfAreas    = 75;
	public static final int xlfRows     = 76;
	public static final int xlfColumns  = 77;
	public static final int xlfOffset   = 78;
	public static final int xlfAbsref   = 79;
	public static final int xlfRelref   = 80;
	public static final int xlfArgument = 81;
	public static final int xlfSearch   = 82;
	public static final int xlfTranspose = 83;
	public static final int xlfError    = 84;
	public static final int xlfStep     = 85;
	public static final int xlfType     = 86;
	public static final int xlfEcho     = 87;
	public static final int xlfSetName  = 88;
	public static final int xlfCaller   = 89;
	public static final int xlfDeref    = 90;
	public static final int xlfWindows  = 91;
	public static final int xlfSeries   = 92;
	public static final int xlfDocuments = 93;
	public static final int xlfActiveCell = 94;
	public static final int xlfSelection = 95;
	public static final int xlfResult   = 96;
	public static final int xlfAtan2    = 97;
	public static final int xlfAsin     = 98;
	public static final int xlfAcos     = 99;
	public static final int xlfChoose   = 100;
	public static final int xlfHlookup  = 101;
	public static final int xlfVlookup  = 102;
	public static final int xlfLinks    = 103;
	public static final int xlfInput    = 104;
	public static final int xlfIsref    = 105;
	public static final int xlfGetFormula = 106;
	public static final int xlfGetName  = 107;
	public static final int xlfSetValue = 108;
	public static final int xlfLog      = 109;
	public static final int xlfExec     = 110;
	public static final int xlfChar     = 111;
	public static final int xlfLower    = 112;
	public static final int xlfUpper    = 113;
	public static final int xlfProper   = 114;
	public static final int xlfLeft     = 115;
	public static final int xlfRight    = 116;
	public static final int xlfExact    = 117;
	public static final int xlfTrim     = 118;
	public static final int xlfReplace  = 119;
	public static final int xlfSubstitute = 120;
	public static final int xlfCode     = 121;
	public static final int xlfNames    = 122;
	public static final int xlfDirectory = 123;
	public static final int xlfFind     = 124;
	public static final int xlfCell     = 125;
	public static final int xlfIserr    = 126;
	public static final int xlfIstext   = 127;
	public static final int xlfIsnumber = 128;
	public static final int xlfIsblank  = 129;
	public static final int xlfT        = 130;
	public static final int xlfN        = 131;
	public static final int xlfFopen    = 132;
	public static final int xlfFclose   = 133;
	public static final int xlfFsize    = 134;
	public static final int xlfFreadln  = 135;
	public static final int xlfFread    = 136;
	public static final int xlfFwriteln = 137;
	public static final int xlfFwrite   = 138;
	public static final int xlfFpos     = 139;
	public static final int xlfDatevalue = 140;
	public static final int xlfTimevalue = 141;
	public static final int xlfSln      = 142;
	public static final int xlfSyd      = 143;
	public static final int xlfDdb      = 144;
	public static final int xlfGetDef   = 145;
	public static final int xlfReftext  = 146;
	public static final int xlfTextref  = 147;
	public static final int xlfIndirect = 148;
	public static final int xlfRegister = 149;
	public static final int xlfCall     = 150;
	public static final int xlfAddBar   = 151;
	public static final int xlfAddMenu  = 152;
	public static final int xlfAddCommand = 153;
	public static final int xlfEnableCommand = 154;
	public static final int xlfCheckCommand = 155;
	public static final int xlfRenameCommand = 156;
	public static final int xlfShowBar  = 157;
	public static final int xlfDeleteMenu = 158;
	public static final int xlfDeleteCommand = 159;
	public static final int xlfGetChartItem = 160;
	public static final int xlfDialogBox = 161;
	public static final int xlfClean    = 162;
	public static final int xlfMdeterm  = 163;
	public static final int xlfMinverse = 164;
	public static final int xlfMmult    = 165;
	public static final int xlfFiles    = 166;
	public static final int xlfIpmt     = 167;
	public static final int xlfPpmt     = 168;
	public static final int xlfCounta   = 169;
	public static final int xlfCancelKey = 170;
	public static final int xlfInitiate = 175;
	public static final int xlfRequest  = 176;
	public static final int xlfPoke     = 177;
	public static final int xlfExecute  = 178;
	public static final int xlfTerminate = 179;
	public static final int xlfRestart  = 180;
	public static final int xlfHelp     = 181;
	public static final int xlfGetBar   = 182;
	public static final int xlfProduct  = 183;
	public static final int xlfFact     = 184;
	public static final int xlfGetCell  = 185;
	public static final int xlfGetWorkspace = 186;
	public static final int xlfGetWindow = 187;
	public static final int xlfGetDocument = 188;
	public static final int xlfDproduct = 189;
	public static final int xlfIsnontext = 190;
	public static final int xlfGetNote  = 191;
	public static final int xlfNote     = 192;
	public static final int xlfStdevp   = 193;
	public static final int xlfVarp     = 194;
	public static final int xlfDstdevp  = 195;
	public static final int xlfDvarp    = 196;
	public static final int xlfTrunc    = 197;
	public static final int xlfIslogical = 198;
	public static final int xlfDcounta  = 199;
	public static final int xlfDeleteBar = 200;
	public static final int xlfUnregister = 201;
	public static final int xlfUsdollar = 204;
	public static final int xlfFindb    = 205;
	public static final int xlfSearchb  = 206;
	public static final int xlfReplaceb = 207;
	public static final int xlfLeftb    = 208;
	public static final int xlfRightb   = 209;
	public static final int xlfMidb     = 210;
	public static final int xlfLenb     = 211;
	public static final int xlfRoundup  = 212;
	public static final int xlfRounddown = 213;
	public static final int xlfAsc      = 214;
	public static final int xlfDbcs     = 215;
	public static final int xlfRank     = 216;
	public static final int xlfAddress  = 219;
	public static final int xlfDays360  = 220;
	public static final int xlfToday    = 221;
	public static final int xlfVdb      = 222;
	public static final int xlfMedian   = 227;
	public static final int xlfSumproduct = 228;
	public static final int xlfSinh     = 229;
	public static final int xlfCosh     = 230;
	public static final int xlfTanh     = 231;
	public static final int xlfAsinh    = 232;
	public static final int xlfAcosh    = 233;
	public static final int xlfAtanh    = 234;
	public static final int xlfDget     = 235;
	public static final int xlfCreateObject = 236;
	public static final int xlfVolatile = 237;
	public static final int xlfLastError = 238;
	public static final int xlfCustomUndo = 239;
	public static final int xlfCustomRepeat = 240;
	public static final int xlfFormulaConvert = 241;
	public static final int xlfGetLinkInfo = 242;
	public static final int xlfTextBox  = 243;
	public static final int xlfInfo     = 244;
	public static final int xlfGroup    = 245;
	public static final int xlfGetObject = 246;
	public static final int xlfDb       = 247;
	public static final int xlfPause    = 248;
	public static final int xlfResume   = 251;
	public static final int xlfFrequency = 252;
	public static final int xlfAddToolbar = 253;
	public static final int xlfDeleteToolbar = 254;
	public static final int xlfADDIN	= 255;	// KSC: Added; Excel function ID for add-ins
	public static final int xlfResetToolbar = 256;
	public static final int xlfEvaluate = 257;
	public static final int xlfGetToolbar = 258;
	public static final int xlfGetTool  = 259;
	public static final int xlfSpellingCheck = 260;
	public static final int xlfErrorType = 261;
	public static final int xlfAppTitle = 262;
	public static final int xlfWindowTitle = 263;
	public static final int xlfSaveToolbar = 264;
	public static final int xlfEnableTool = 265;
	public static final int xlfPressTool = 266;
	public static final int xlfRegisterId = 267;
	public static final int xlfGetWorkbook = 268;
	public static final int xlfAvedev   = 269;
	public static final int xlfBetadist = 270;
	public static final int xlfGammaln  = 271;
	public static final int xlfBetainv  = 272;
	public static final int xlfBinomdist = 273;
	public static final int xlfChidist  = 274;
	public static final int xlfChiinv  = 275;
	public static final int xlfCombin   = 276;
	public static final int xlfConfidence = 277;
	public static final int xlfCritbinom = 278;
	public static final int xlfEven     = 279;
	public static final int xlfExpondist = 280;
	public static final int xlfFdist    = 281;
	public static final int xlfFinv     = 282;
	public static final int xlfFisher   = 283;
	public static final int xlfFisherinv = 284;
	public static final int xlfFloor    = 285;
	public static final int xlfGammadist =  286;
	public static final int xlfGammainv     = 287;
	public static final int xlfCeiling  = 288;
	public static final int xlfHypgeomdist = 289;
	public static final int xlfLognormdist = 290;
	public static final int xlfLoginv   = 291;
	public static final int xlfNegbinomdist = 292;
	public static final int xlfNormdist = 293;
	public static final int xlfNormsdist    = 294;
	public static final int xlfNorminv  = 295;
	public static final int xlfNormsinv = 296;
	public static final int xlfStandardize = 297;
	public static final int xlfOdd      = 298;
	public static final int xlfPermut   = 299;
	public static final int xlfPoisson  = 300;
	public static final int xlfTdist    = 301;
	public static final int xlfWeibull  = 302;
	public static final int xlfSumxmy2  = 303;
	public static final int xlfSumx2my2 = 304;
	public static final int xlfSumx2py2 = 305;
	public static final int xlfChitest  = 306;
	public static final int xlfCorrel   = 307;
	public static final int xlfCovar    = 308;
	public static final int xlfForecast = 309;
	public static final int xlfFtest    = 310;
	public static final int xlfIntercept = 311;
	public static final int xlfPearson  = 312;
	public static final int xlfRsq      = 313;
	public static final int xlfSteyx    = 314;
	public static final int xlfSlope    = 315;
	public static final int xlfTtest    = 316;
	public static final int xlfProb     = 317;
	public static final int xlfDevsq    = 318;
	public static final int xlfGeomean  = 319;
	public static final int xlfHarmean  = 320;
	public static final int xlfSumsq    = 321;
	public static final int xlfKurt     = 322;
	public static final int xlfSkew     = 323;
	public static final int xlfZtest    = 324;
	public static final int xlfLarge    = 325;
	public static final int xlfSmall    = 326;
	public static final int xlfQuartile = 327;
	public static final int xlfPercentile = 328;
	public static final int xlfPercentrank = 329;
	public static final int xlfMode     = 330;
	public static final int xlfTrimmean = 331;
	public static final int xlfTinv     = 332;
	public static final int xlfMovieCommand = 334;
	public static final int xlfGetMovie = 335;
	public static final int xlfConcatenate = 336;
	public static final int xlfPower    = 337;
	public static final int xlfPivotAddData = 338;
	public static final int xlfGetPivotTable = 339;
	public static final int xlfGetPivotField = 340;
	public static final int xlfGetPivotItem = 341;
	public static final int xlfRadians  = 342;
	public static final int xlfDegrees  = 343;
	public static final int xlfSubtotal = 344;
	public static final int xlfSumif    = 345;
	public static final int xlfCountif  = 346;
	public static final int xlfCountblank = 347;
	public static final int xlfScenarioGet = 348;
	public static final int xlfOptionsListsGet = 349;
	public static final int xlfIspmt    = 350;
	public static final int xlfDatedif  = 351;
	public static final int xlfDatestring = 352;
	public static final  int xlfNumberstring = 353;
	public static final int xlfRoman    = 354;
	public static final int xlfOpenDialog = 355;
	public static final int xlfSaveDialog = 356;
	public static final int xlfViewGet  = 357;
	public static final int xlfGetPivotData = 358;
	public static final int xlfHyperlink = 359;
	public static final int xlfPhonetic     = 360;
	public static final int xlfAverageA     = 361;
	public static final int xlfMaxA     = 362;
	public static final int xlfMinA     = 363;
	public static final int xlfStDevPA  = 364;
	public static final int xlfVarPA    = 365;
	public static final int xlfStDevA   = 366;
	public static final int xlfVarA     = 367;
	// KSC: ADD-IN formulas - use any index; name must be present in FunctionConstants.addIns
	// Financial Formulas
	public static final int xlfAccrintm= 368;
	public static final int xlfAccrint= 369;
	public static final int xlfCoupDayBS= 370;
	public static final int xlfCoupDays= 371;
	public static final int xlfCumIPmt= 372;
	public static final int xlfCumPrinc= 373;
	public static final int xlfCoupNCD= 374;
	public static final int xlfCoupDaysNC= 375;
	public static final int xlfCoupPCD= 376;
	public static final int xlfCoupNUM= 377;
	public static final int xlfDollarDE= 378;
	public static final int xlfDollarFR= 379;
	public static final int xlfEffect= 380;
	public static final int xlfINTRATE= 381;
	public static final int xlfXIRR= 382;
	public static final int xlfXNPV= 383;
	public static final int xlfYIELD= 384;
	public static final int xlfPRICE= 385;
	public static final int xlfPRICEDISC= 386;
	public static final int xlfPRICEMAT= 387;
	public static final int xlfDURATION= 388;
	public static final int xlfMDURATION= 389;
	public static final int xlfTBillEq= 390;
	public static final int xlfTBillPrice= 391;
	public static final int xlfTBillYield= 392;
	public static final int xlfYieldDisc= 393;
	public static final int xlfYieldMat= 394;
	public static final int xlfFVSchedule= 395;
	public static final int xlfAmorlinc= 396;
	public static final int xlfAmordegrc= 397;
	public static final int xlfOddFPrice= 398;
	public static final int xlfOddLPrice= 399;
	public static final int xlfOddFYield= 400;
	public static final int xlfOddLYield= 401;
	public static final int xlfNOMINAL= 402;
	public static final int xlfDISC= 403;
	public static final int xlfRECEIVED= 404;
	// Engineering Formulas
	public static final int xlfBIN2DEC= 405; 
	public static final int xlfBIN2HEX= 406; 
	public static final int xlfBIN2OCT= 407; 
	public static final int xlfDEC2BIN= 408; 
	public static final int xlfDEC2HEX= 409; 
	public static final int xlfDEC2OCT= 410; 
	public static final int xlfHEX2BIN= 411; 
	public static final int xlfHEX2DEC= 412; 
	public static final int xlfHEX2OCT= 413; 
	public static final int xlfOCT2BIN= 414; 
	public static final int xlfOCT2DEC= 415; 
	public static final int xlfOCT2HEX= 416;
	public static final int xlfCOMPLEX= 417;
	public static final int xlfGESTEP= 	418;
	public static final int xlfDELTA= 	419;
	public static final int xlfIMAGINARY= 420;
	public static final int xlfIMABS=	421;
	public static final int xlfIMDIV=	422;
	public static final int xlfIMCONJUGATE= 423;
	public static final int xlfIMCOS=	424;
	public static final int xlfIMSIN=	425;
	public static final int xlfIMREAL=	426;
	public static final int xlfIMEXP=	427;
	public static final int xlfIMSUB=	428;
	public static final int xlfIMSUM=	429;
	public static final int xlfIMPRODUCT= 430;
	public static final int xlfIMLN=	431;
	public static final int xlfIMLOG10= 432;
	public static final int xlfIMLOG2=	433;
	public static final int xlfIMPOWER=	434;
	public static final int xlfIMSQRT=	435;
	public static final int xlfIMARGUMENT= 436;
	public static final int xlfCONVERT= 437;
	// Math Add-in Formulas
	public static final int xlfDOUBLEFACT= 438;
	public static final int xlfGCD=		439;
	public static final int xlfLCM=		440;
	public static final int xlfMROUND=	441;
	public static final int xlfMULTINOMIAL= 442;
	public static final int xlfQUOTIENT=	443;
	public static final int xlfRANDBETWEEN= 444;
	public static final int xlfSERIESSUM=	445;
	public static final int xlfSQRTPI=		446;
*/
}