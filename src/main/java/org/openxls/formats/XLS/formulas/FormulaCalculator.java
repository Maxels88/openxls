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
package org.openxls.formats.XLS.formulas;

import org.openxls.formats.XLS.FunctionNotSupportedException;
import org.openxls.formats.XLS.XLSRecord;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Stack;

/**
 * Formula Calculator.
 * <p/>
 * Translates an excel calc stack into a value.  The stack is in a modified
 * reverse polish notation.  For details please look into the Excel 97 reference for BIFF8.
 * <p/>
 * Formula Calculators do not exist per formula, rather it uses a factory pattern.
 * Put a formula in & calculate it.
 * <p/>
 * Actual calculation methods exist within the OperatorPtg's themselves, so this just handles
 * the grunt work of parsing and passing ptgs back and forth along with formatting the output.
 *
 * @see Formula
 */
public class FormulaCalculator
{
	private static final Logger log = LoggerFactory.getLogger( FormulaCalculator.class );

	/**
	 * Calculates the value of calcStac  This is handled by
	 * running through the stack, adding operands to tempstack until
	 * an operator PTG is found. At that point pass the relevant ptg's from
	 * tempstack into the calculate method of the operator PTG.  The operator
	 * ptg should return a valid value PTG.
	 */
	public static Object calculateFormula( Stack expression ) throws FunctionNotSupportedException
	{
		int sz = expression.size();
		Ptg[] stck = new Ptg[sz];
		stck = (Ptg[]) expression.toArray( stck );
		Stack calcStack = new Stack();
		for( int t = 0; t < sz; t++ )
		{
			// flip the stack TODO: investigate why needed
			calcStack.add( 0, stck[t] );
		}
		Stack tempstack = new Stack();
		while( !calcStack.isEmpty() )
		{
			// loop while there are Ptgs
			handlePtg( calcStack, tempstack );
		}
		Ptg finalptg = (Ptg) tempstack.pop();
		return finalptg.getValue();
	}

	/**
	 * Calculates the final Ptg result of calcStack.  This is handled by
	 * running through the stack, adding operands to tempstack until
	 * an operator PTG is found. At that point pass the relevant ptg's from
	 * tempstack into the calculate method of the operator PTG.  The operator
	 * ptg should return a valid value PTG.
	 */
	public static Ptg calculateFormulaPtg( Stack expression ) throws FunctionNotSupportedException
	{
		int sz = expression.size();
		Ptg[] stck = new Ptg[sz];
		stck = (Ptg[]) expression.toArray( stck );
		Stack calcStack = new Stack();
		for( int t = 0; t < sz; t++ )
		{ // flip the stack TODO: investigate why needed
			calcStack.add( 0, stck[t] );
		}
		Stack tempstack = new Stack();
		Stack newstck = calcStack;
		while( !newstck.isEmpty() )
		{// loop while there are Ptgs
			handlePtg( newstck, tempstack );
		}
		Ptg finalptg = (Ptg) tempstack.pop();
		return finalptg;
	}

	/**
	 * This is a very similar method to the handle ptg method in formula parser.
	 * Instead of creating a tree however it calculates in the order recommended by
	 * the book of knowledge (excel developers guide).  That is, FILO.  First In Last Out.
	 * We also don't really care about things like parens, they are just for display purposes.
	 */
	static void handlePtg( Stack newstck, Stack vals ) throws FunctionNotSupportedException
	{
		Ptg p = (Ptg) newstck.pop();

		if( p.getIsOperator() || p.getIsControl() || p.getIsFunction() )
		{
			// Get rid of the parens ptgs
			if( p.getIsControl() && !vals.isEmpty() )
			{
				if( p.getOpcode() == 0x15 )
				{
					// its a parens!
					// the parens is already pop'd so just return and it is gone...
					return;
				}
				// we didn't use it, back it goes.
				log.trace( "opr: " + p.toString() );
			}
			// make sure we have the correct amount popped back in..
			int t = 0;

			if( p.getIsBinaryOperator() )
			{
				t = 2;
			}

			if( p.getIsUnaryOperator() )
			{
				t = 1;
			}

			if( p.getIsStandAloneOperator() )
			{
				t = 0;
			}

			if( (p.getOpcode() == 0x22) || (p.getOpcode() == 0x42) || (p.getOpcode() == 0x62) )
			{
				t = p.getNumParams();
			}
			else if( (p.getOpcode() == 0x21) || (p.getOpcode() == 0x41) || (p.getOpcode() == 0x61) )
			{
				t = p.getNumParams();
			}// guess that ptgfunc is not only one..

			Ptg[] vx = new Ptg[t];

			for( int x = 0; x < t; x++ )
			{
				// Ok, at this point we know how many operands this current function/method requires (t), so pop them off the vals stack
				// into a new one for this function/operands processing.
				vx[(t - 1) - x] = (Ptg) vals.pop();
			}

			Ptg result;
			try
			{
				// FIXME: We need to know if we are inside a SUMPRPODUCT here (and possibly other funky functions) where primitive operations
				// FIXME: have different results - case in point: SUMPRODUCT( 0+ (A1:A10>0) )  - this is a coercion function within Excel.
				result = p.calculatePtg( vx );
			}
			catch( CalculationException e )
			{
				result = new PtgErr( e.getErrorCode() );
				if( e.getName().equals( "#CIR_ERR!" ) )
				{
					((PtgErr) p).setCircularError( true );
				}
			}

        	/* useful for debugging*/
			String adr = "";
			XLSRecord parentRec = result.getParentRec();
			if( parentRec != null )
			{
				adr = "addr: " + parentRec.getCellAddress();
			}
			log.trace( "{} val: {}", adr, result.toString() );
			vals.push( result );// push it back on the stack

		}
		else if( p.getIsOperand() )
		{
			log.trace( "opr: {}", p.toString() );

			vals.push( p );
		}
		else if( p instanceof PtgAtr )
		{
			PtgAtr ptgAtr = (PtgAtr) p;
			// FIXME: "Probably" ?
			// this is probably just a space at this point, don't output error message
			log.debug( "Possible 'space' - PtgAtr - ignoring for now: {}", ptgAtr.toString() );
		}
		else
		{
			throw new FunctionNotSupportedException( "WARNING: Calculating Formula failed: Unsupported/Incorrect Ptg Type: 0x" + p.getOpcode() + " " + p
					.getString() );
		}
	}
}