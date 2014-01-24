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
import com.extentech.formats.XLS.Array;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.InvalidRecordException;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.XLSRecordFactory;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Locale;
import java.util.Stack;
import java.util.Vector;

//import com.sun.org.apache.xerces.internal.impl.xs.identity.Selector.Matcher;

/**
 * Formula Parser.
 * <p/>
 * Translates Excel-compatible Strings into Biff8/ExtenXLS Compatible Formulas and vice-versa.
 *
 * @see Formula
 */
public final class FormulaParser
{
	private static final Logger log = LoggerFactory.getLogger( FormulaParser.class );
	/**
	 * getPtgsFromFormulaString
	 * returns ordered stack of Ptgs parsed from formula string fmla
	 *
	 * @param Formula form	formula record
	 * @param String  fmla		string rep. of current state of formula (either original "=F(Y)" or a recurrsed state e.g. "Y")
	 * @returns Stack    ordered Ptgs that represent formula expression
	 */
	public static Stack getPtgsFromFormulaString( XLSRecord form, String fmla )
	{
		return getPtgsFromFormulaString( form, fmla, true );
	}

	/**
	 * getPtgsFromFormulaString is the main entry point for parsing a string and creating a formula.
	 * The formula passed in at this point can either be an existing formula with an expression, or a
	 * templated formula with no expression. The string gets parsed and entered
	 * as the expression for the formula.
	 *
	 * @param    Formula form	formula record
	 * @param    String fmla		string rep. of current state of formula (either original "=F(Y)" or a recurrsed state e.g. "Y")
	 * @param    boolean bIsCompleteExpression	truth of "formula fmla represents a complete formula i.e. we are not currently in a recurrsed state"
	 * @returns Stack    ordered Ptgs that represent formula expression 
	 */
	/**
	 * should handle all sorts of variations of formula strings such as:
	 * =PV(C17,1+-(1*1)-9, 0, 1)
	 * =100*0.5
	 * =(B2-B3)*B4
	 * =SUM(IF(A1:A10=B1:B10, 1, 0))
	 * =IF(B4<=10,"10", if(b4<=100, "15", "20"))
	 * ="STRING"&IF(A<>"",A,"N/A")&" - &IF(C<>"",C,"N/A")&" Result "
	 * <p/>
	 * in basic essence, handles signatures such as
	 * a op f(b op c, d, uop e ...) op g
	 * <p/>
	 * where op is any binary operator, uop is a unary operator f is a formula
	 * ...
	 */
	protected static Stack getPtgsFromFormulaString( XLSRecord form, String fmla, boolean bMergeWithLast )
	{
		Object[] operands = new Object[2];

		fmla = fmla.trim();
		if( fmla.startsWith( "=" ) )
		{
			fmla = fmla.substring( 1 );
		}
		fmla = fmla.trim();

		if( fmla.equals( "" ) )
		{    // 20081120 KSC: Handle Missing Argument
			Stack s = new Stack();
			s.add( new PtgMissArg() );
			return s;
		}

		if( fmla.startsWith( "{" ) )
		{ // must process array formula first, PtgArray expects full function string
			PtgArray pa = new PtgArray();
			pa.setParentRec( form );
			int endarray = getMatchOperator( fmla, 0, '{', '}' );
			pa.setVal( fmla.substring( 0, endarray + 1 ) );
			fmla = fmla.substring( endarray + 1 );
			Stack s = new Stack();
			s.add( pa );
			operands[0] = s;
			bMergeWithLast = false;
		}
// TODO: complex ranges??  	  
		boolean inQuote = false;
		boolean inRange = false;
		boolean inOp = false;
		String prefix = "";
		Stack ops = new Stack();
		String op = "";
		for( int i = 0; i < fmla.length(); i++ )
		{
			char c = fmla.charAt( i );
			if( (c == '"') || (c == '\'') )
			{ // get to ending quote
				if( inQuote )
				{
					inQuote = (c != prefix.trim().charAt( 0 ));    // if start quote == end quote, inQuote is false 
				}
				else
				{
					inQuote = true;
				}
				if( inQuote )
				{
					if( inOp )
					{
						inOp = false;
						if( !op.equals( "" ) )
						{
							ops.add( 0, op );
						}
						op = "";
						if( (operands[0] != null) && (operands[1] != null) )
						{
							addOperands( form, /*functionStack, */operands, ops );
						}
					}
				}
				prefix += c;
			}
			else if( inQuote )
			{
				prefix += c;
			}
			else if( c == ':' )
			{
				if( i > 0 )
				{
					inRange = true;
				}
				prefix += c;
			}
			else if( c == '(' )
			{        // found a formula?? check out
				// if the parenthesis is part of a complex range, keep going i.e. keep entire expression together 
				if( inRange )
				{
					prefix += c;
					continue;
				}
				if( inOp )
				{
					inOp = false;
					if( !op.equals( "" ) )
					{
						ops.add( 0, op );
					}
					op = "";
					if( !ops.isEmpty() && (operands[0] != null) && (operands[1] != null) )
					{
						addOperands( form, /*functionStack, */operands, ops );
					}
				}
				String funcName = "";
				for( int k = prefix.length() - 1; k >= 0; k-- )
				{
					if( Character.isLetterOrDigit( prefix.charAt( k ) ) || Character.toString( prefix.charAt( k ) ).equals( "." ) )
					{
						funcName = prefix.charAt( k ) + funcName;
						prefix = prefix.substring( 0, k );
					}
					else
					{
						break;
					}
				}
				// prefix= anything before function name
				if( !prefix.trim().equals( "" ) )
				{
					if( operands[0] == null )
					{
						operands[0] = prefix.trim();
					}
					else
					{
						operands[1] = prefix.trim();
					}
					prefix = "";
				}
				// function name should = part just before parents
				if( !funcName.equals( "" ) )
				{
					Ptg funcPtg = null;
					funcPtg = getFuncPtg( funcName, form );

					// do we have a valid function Ptg?
					if( funcPtg != null )
					{    // yes, then handle function paramters i.e. evertyhing between the parentheses
						int endparen = getMatchOperator( fmla, i, '(', ')' );
						if( endparen < (fmla.length() - 1) )
						{
							if( fmla.charAt( endparen + 1 ) == ':' )
							{    // it's a VERY complex range :)
								inRange = true;
								prefix = funcName + fmla.substring( i, endparen + 1 );    // keep function name together ...
								i = endparen;
								continue;
							}
						}
						// things like: xyz + f(x)
						if( !ops.isEmpty() && (operands[0] != null) && (operands[1] != null) )
						{
							addOperands( form, /*functionStack, */operands, ops );
							inOp = false;
						}    // have [xyz, xyz, OP] + [abc, def, ghi]
						// parse function
						Stack s = parseFunctionPtg( form, funcName, fmla.substring( i + 1, endparen ), funcPtg );
						if( operands[0] == null )
						{
							operands[0] = s;
						}
						else
						{
							operands[1] = s;
						}
						//functionStack.addAll(parseFunctionPtg(form, funcName, fmla.substring(i+1, endparen), funcPtg));
						i = endparen;    // inc. pointer to past processing point
					}
					else    // else, we have *something* in front of the parentheses ... 
					{
						throw new FunctionNotSupportedException( funcName + " is not a supported function" );
					}
				}
				else
				{        // enclosing parens
					// complexities occur for complex ranges and enclosing parens ...
					int endparen = getMatchOperator( fmla, i, '(', ')' );
					if( (endparen == -1) || ((endparen < (fmla.length() - 1)) && (fmla.charAt( endparen + 1 ) == ':')) )
					{    // it's a VERY complex range :)  
						inRange = true;
						prefix = "(";
						continue;
					}
					String f = fmla.substring( i + 1, endparen );    // the statement less the parenthesis
					i = endparen;    // skip parens ...
					// see if the enclosed expression is a complex range - must parse as 1 unit, rather than parsing particular ptgs
					if( FormulaParser.isComplexRange( '(' + f + ')' ) )
					{
						Stack s = new Stack();
						s.push( parseSinglePtg( form, '(' + f + ')', false ) );
						s.push( parseSinglePtg( form, ")", false ) );
						if( operands[0] == null )
						{
							operands[0] = s;
						}
						else
						{
							operands[1] = s;
						}
					}
					else
					{    // embedded functions, keep parsing
						Stack s = getPtgsFromFormulaString( form, f, true );    // flag as a complete expression
						s.push( new PtgParen() );    // add ending parens to stack
						if( operands[0] == null )
						{
							operands[0] = s;
						}
						else
						{
							operands[1] = s;
						}
						if( !ops.isEmpty() )
						{
							addOperands( form, /*functionStack, */operands, ops );
						}
					}
				}
			}
			else
			{    // see if we have found an operataor
				if( !Character.isJavaIdentifierPart( c ) && (c != ' ') && (c != '%') )
				{
					//  if (inRange && !Character.isJavaIdentifierPart(c) && c!=',' && c!=' ')
					if( inRange )
					{
						if( (c != ',') && (c != ' ') && (c != ')') && (c != '!') )
						{
							inRange = false;
							if( !prefix.trim().equals( "" ) )
							{
								if( operands[0] == null )
								{
									operands[0] = prefix.trim();
								}
								else
								{
									operands[1] = prefix.trim();
								}
								prefix = "";
							}
						}
						else
						{
							prefix += c;
							continue;
						}
					}
					if( (c != '!') && (c != '#') && (c != '.') )
					{    // ignore !
						// FOUND AN OPERATOR - ready to add operands yet?
						inOp = true;
						if( !prefix.trim().equals( "" ) )
						{
							if( operands[0] == null )
							{
								operands[0] = prefix.trim();
							}
							else
							{
								operands[1] = prefix.trim();
							}
							prefix = "";
						}
						if( !ops.isEmpty() && (operands[0] != null) && (operands[1] != null) )
						{
							addOperands( form, /*functionStack, */operands, ops );
						}
						if( !op.equals( "" ) )
						{
							if( !((c == '=') || (c == '>')) )
							{
								ops.add( 0, op );
								op = "";
							}
						}
						op += c;    // >,<,-,/, ,+
						continue;
					}
				}
				else if( inOp )
				{
					inOp = false;
					if( !ops.isEmpty() && (operands[0] != null) && (operands[1] != null) )
					{
						addOperands( form, /*functionStack, */operands, ops );
					}
					if( !op.equals( "" ) )
					{
						ops.add( 0, op );
					}
					op = "";
				}
				prefix += c;
			}

		}
		if( !prefix.trim().equals( "" ) )
		{    // get any remaining elements 
			if( operands[0] == null )
			{
				operands[0] = prefix.trim();
			}
			else
			{
				operands[1] = prefix.trim();
			}
			prefix = "";
		}

		if( !op.equals( "" ) )
		{
			ops.add( 0, op );
		}
		addOperands( form, operands, ops );
//if (((Stack)operands[0]).isEmpty())
//return functionStack;
		return (Stack) operands[0];

	}

	/**
	 * Given two operands (objects) and operators (can be up to 2 if there is a unary operator present)
	 * organize and add to functionStack in reverse polish notation i.e.
	 * OPERAND, OPERAND, OP [...]
	 *
	 * @param form
	 * @param functionStack
	 * @param operands
	 * @param ops
	 */
	private static void addOperands( XLSRecord form, /*Stack functionStack, */Object[] operands, Stack ops )
	{
		Ptg pOp = null;
		if( !ops.isEmpty() )
		{
			pOp = parseSinglePtg( form, (String) ops.pop(), (operands[1] == null) );
		}

		Stack s = new Stack();
		s.addAll( handleOperatorPrecedence( form, /*functionStack, */operands, pOp ) );
		operands[0] = s;
		//functionStack.clear();
		if( !ops.isEmpty() )
		{
			pOp = parseSinglePtg( form, (String) ops.pop(), true );
			s = new Stack();
			s.addAll( handleOperatorPrecedence( form, /*functionStack, */operands, pOp ) );
			operands[0] = s;
		}

	}

	private static Stack handleOperatorPrecedence( XLSRecord form, /*Stack functionStack, */Object[] operands, Ptg pOp )
	{
		Stack functionStack = new Stack();
		if( operands[0] instanceof Stack )
		{
			functionStack = (Stack) operands[0];
			operands[0] = null;
		}
		if( !functionStack.isEmpty() && (pOp != null) )
		{
			Ptg lastOp = (Ptg) functionStack.peek();
			if( (lastOp != null) && lastOp.getIsOperator() )
			{
				if( lastOp.getIsOperator() )
				{
					functionStack.pop();    //= lastOp
					int group1 = rankPrecedence( pOp );
					int group2 = rankPrecedence( lastOp );
					if( group2 >= group1 )
					{
						functionStack.push( lastOp );
					}
					else
					{    // current op has higher priority
						if( operands[0] != null )
						{
							if( operands[0] instanceof String )
							{
								functionStack.push( parseSinglePtg( form, (String) operands[0], functionStack.isEmpty() ) );
							}
							else
							{
								functionStack.addAll( (Stack) operands[0] );
							}
						}
						if( operands[1] != null )
						{
							if( operands[1] instanceof String )
							{
								functionStack.push( parseSinglePtg( form, (String) operands[1], functionStack.isEmpty() ) );
							}
							else
							{
								functionStack.addAll( (Stack) operands[1] );
							}
						}
						operands[0] = null;
						operands[1] = null;
						functionStack.push( pOp );
						pOp = lastOp;
					}
				}
			}
		}
		if( operands[0] != null )
		{
			if( operands[0] instanceof String )
			{
				functionStack.push( parseSinglePtg( form, (String) operands[0], functionStack.isEmpty() ) );
			}
			else
			{
				functionStack.addAll( (Stack) operands[0] );
			}
		}
		if( operands[1] != null )
		{
			if( operands[1] instanceof String )
			{
				functionStack.push( parseSinglePtg( form, (String) operands[1], functionStack.isEmpty() ) );
			}
			else
			{
				functionStack.addAll( (Stack) operands[1] );
			}
		}
		operands[0] = null;
		operands[1] = null;
		if( pOp != null )
		{
			functionStack.push( pOp );
		}
		return functionStack;
	}

	/**
	 * merge last stacks to ensure operator order is correct
	 *
	 * @param functionStack
	 * @return
	 */
	private static Stack mergeStacks( Stack prevStack, Stack curStack, boolean bIsCompleteExpression )
	{
		if( prevStack.isEmpty() )
		{
			return curStack;
		}

		Ptg lastOp = (Ptg) prevStack.peek();
		Ptg curOp = (curStack.isEmpty() ? null : ((Ptg) curStack.peek()));
		int group1 = rankPrecedence( lastOp );
		int group2 = rankPrecedence( curOp );
		if( (group1 >= 0) && ((group1 < group2) || (group2 == -1)) )
		{
			lastOp = (Ptg) prevStack.pop();
			curStack.push( lastOp );  				
/*  				while (curOp.getIsOperator()) {
  					lastOp= curOp;
  					curStack.push(curOp);	
  					if (!prevStack.isEmpty()) {
  						curOp = (Ptg) prevStack.pop();					
  						// handle precedence 
  				  		group1=rankPrecedence(curOp); 
  				  		group2=rankPrecedence(lastOp);
  				  		if (!(group1>=0 && (group1 < group2 || group2==-1)))
  				  			break;
  					}else		
  						return curStack;
  				}
*/
		}
		prevStack.addAll( curStack );
		return prevStack;
	}

	/**
	 * parse and add to Stack a valid Excel function represented by funcPtg and fmla string
	 * called from getPtgsFromFormulaString
	 *
	 * @param form    formula record
	 * @param fmla    function parameters in the form of (x, y, z)
	 * @param funcPtg function data for the formula represented by fmla
	 * @return Stack    ordered parsed Stack of Ptgs
	 * @para func        function name f
	 */
	private static Stack parseFunctionPtg( XLSRecord form, String func, String fmla, Ptg funcPtg )
	{
		Stack returnStack = new Stack();
		fmla = fmla.trim();
		int nParens = 0;
		boolean enclosing = false;
		// change:  only remove 1 set of parens:
		if( (fmla.length() > 0) && (fmla.charAt( 0 ) == '(') )
		{
			if( getMatchOperator( fmla, 0, '(', ')' ) == (fmla.length() - 1) )
			{    // then strip enclosing parens
				nParens++;
				enclosing = true;
			}
			fmla = fmla.trim();
		}

		// NOTE: all memfuncs/complex ranges are enclosed by parentheses 
		// IF enclosed by parens, DO NOT split apart into operands:
		int funcLen = 1;
		if( enclosing )
		{
			returnStack.addAll( FormulaParser.getPtgsFromFormulaString( form, fmla, true ) );
		}
		else
		{
			CompatibleVector cv = splitFunctionOperands( fmla );
			funcLen = cv.size();
			// loop through the operands to the function and recurse
			for( int y = 0; y < cv.size(); y++ )
			{
				String s = (String) cv.elementAt( y );        // flag as a complete expression
				returnStack.addAll( FormulaParser.getPtgsFromFormulaString( form, s, true ) );
			}
		}

		// Handle PtgFuncVar-specifics such as number of parameters and add-in PtgNameX record
		if( funcPtg instanceof PtgFuncVar )
		{
			if( ((PtgFuncVar) funcPtg).getVal() == FunctionConstants.xlfADDIN )
			{
				// if an add-in, must add PtgNameX to stack
				PtgNameX pn = new PtgNameX();
				pn.setParentRec( form );
				pn.setName( func );
				returnStack.add( 0, pn );    // add to bottom of stack
				funcLen++;
				funcPtg.setParentRec( form ); // nec. to resolve external name
			}
			((PtgFuncVar) funcPtg).setNumParams( (byte) funcLen );
		}
		returnStack.push( funcPtg );
		return returnStack;
	}

	/**
	 * combine two stacks of Ptgs, popping the operator of the sourceStack and
	 * adding it to the end of the destination stack to 
	 * ensure it is in the correct order in the destination stack
	 *
	 * @param sourceStack
	 * @param destStack
	 * /
	private static Stack addPtgStacks(Stack sourceStack, Stack destStack) {
	Ptg opPtg = (Ptg) sourceStack.pop();
	Ptg lastOp= (destStack.isEmpty()?null:((Ptg)destStack.peek()));	

	// handle precedence:  unaries before ^(power) before *, / before +, - before &(concat), before comparisons (=, <>, <=, >=, <, >)
	int group1=rankPrecedence(opPtg); 
	int group2=rankPrecedence(lastOp);

	if (group1>=0 && (group1 < group2 || group2==-1)) {
	while (opPtg.getIsOperator()) {
	lastOp= opPtg;
	destStack.push(opPtg);	
	if (!sourceStack.isEmpty()) {
	opPtg = (Ptg) sourceStack.pop();					
	// handle precedence 
	group1=rankPrecedence(opPtg); 
	group2=rankPrecedence(lastOp);
	if (!(group1>=0 && (group1 < group2 || group2==-1)))
	break;
	}else		
	return destStack;
	}
	} 
	sourceStack.push(opPtg);

	// after sorting out operators, assemble two stacks into one
	Stack nwstack = destStack;
	destStack = new Stack();
	destStack.addAll(sourceStack);
	destStack.addAll(nwstack);

	return destStack;
	}
	 */

	/**
	 * rank a Ptg Operator's precedence (lower
	 *
	 * @param curOp
	 * @return
	 */
	static int rankPrecedence( Ptg curOp )
	{
		if( curOp == null )
		{
			return -1;
		}
//		if (curOp==null || !curOp.getIsOperator()) return -1;
		if( (curOp instanceof PtgUMinus) || (curOp instanceof PtgUPlus) )
		{
			return 7;
		}
		if( curOp instanceof PtgPercent )
		{
			return 6;
		}
		if( curOp instanceof PtgPower )
		{
			return 5;
		}
		if( (curOp instanceof PtgMlt) || (curOp instanceof PtgDiv) )
		{
			return 4;
		}
		if( (curOp instanceof PtgAdd) || (curOp instanceof PtgSub) )
		{
			return 3;
		}
		if( curOp instanceof PtgConcat )
		{
			return 2;
		}
		if( (curOp instanceof PtgEQ) || (curOp instanceof PtgNE) ||
				(curOp instanceof PtgLE) || (curOp instanceof PtgLT) ||
				(curOp instanceof PtgGE) || (curOp instanceof PtgGT) )
		{
			return 1;
		}
//  		else if (curOp instanceof PtgParen)
//  			return 0;
		return -1;
	}

	/*
	* getMatchOperator takes a string and starting operator location.
    * It then parses the string and determines which closing operator
    * matches the opening parens specified by startParenLoc.  Returns
    * -1 if it cannot find a match.
    */
	public static int getMatchOperator( String input, int startParenLoc, char matchOpenChar, char matchCloseChar )
	{
		// 20081112 KSC: do a different way as it wasn't working for all cases
		int openCnt = 0;
		for( int i = startParenLoc; i < input.length(); i++ )
		{
			if( (input.charAt( i ) == '"') || (input.charAt( i ) == '\'') )
			{// handle quoted strings within input (quoted strings may of course contain match chars ...
				char endquote = input.charAt( i );
				while( ++i < input.length() )
				{
					if( input.charAt( i ) == endquote )
					{
						break;
					}
				}
			}
			if( i == input.length() )
			{
				return i - 1;
			}
			if( input.charAt( i ) == matchOpenChar )
			{
				openCnt++;
			}
			else if( input.charAt( i ) == matchCloseChar )
			{
				openCnt--;
				if( openCnt == 0 )
				{
					return i;
				}
			}
		}

		// no parens for you!
		return -1;
	}

	/**
	 * Looks up a function string and returns a funcPtg if it is found
	 *
	 * @param func function string without parents i.e. SUM or DB
	 * @returns Ptg        valid funcPtg or null if not found
	 */
	// 20090210 KSC: add form so can set parent record for PtgFunc and PtgFuncVar - nec for self-referential formulas such as COLUMN
	private static Ptg getFuncPtg( String func, XLSRecord form )
	{
		Ptg funcPtg = null;
		//    if (true) {
		if( Locale.JAPAN.equals( Locale.getDefault() ) )
		{
			for( int y = 0; y < FunctionConstants.jRecArr.length; y++ )
			{
				if( func.equalsIgnoreCase( FunctionConstants.jRecArr[y][0] ) )
				{
					int FID = Integer.parseInt( FunctionConstants.jRecArr[y][1] );
					int Ftype = Integer.parseInt( FunctionConstants.jRecArr[y][2] );
					if( Ftype == FunctionConstants.FTYPE_PTGFUNC )
					{
						funcPtg = new PtgFunc( FID, form );
					}
					else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR )
					{
						funcPtg = new PtgFuncVar( FID, 0, form );
					}
					else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN )
					{
						funcPtg = new PtgFuncVar( FunctionConstants.xlfADDIN, 0, form );
					}
					return funcPtg;
				}
			}
		}
		for( int y = 0; y < FunctionConstants.recArr.length; y++ )
		{
			if( func.equalsIgnoreCase( FunctionConstants.recArr[y][0] ) )
			{
				int FID = Integer.parseInt( FunctionConstants.recArr[y][1] );
				int Ftype = Integer.parseInt( FunctionConstants.recArr[y][2] );
				if( Ftype == FunctionConstants.FTYPE_PTGFUNC )
				{
					funcPtg = new PtgFunc( FID, form );
				}
				else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR )
				{
					funcPtg = new PtgFuncVar( FID, 0, form );
				}
				else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN )
				{
					funcPtg = new PtgFuncVar( FunctionConstants.xlfADDIN, 0, form );
				}
				return funcPtg;
			}
		}
		for( int y = 0; y < FunctionConstants.unimplRecArr.length; y++ )
		{
			if( func.equalsIgnoreCase( FunctionConstants.unimplRecArr[y][0] ) )
			{
				int FID = Integer.parseInt( FunctionConstants.unimplRecArr[y][1] );
				int Ftype = Integer.parseInt( FunctionConstants.unimplRecArr[y][2] );
				if( Ftype == FunctionConstants.FTYPE_PTGFUNC )
				{
					funcPtg = new PtgFunc( FID, form );
				}
				else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR )
				{
					funcPtg = new PtgFuncVar( FID, 0, form );
				}
				else if( Ftype == FunctionConstants.FTYPE_PTGFUNCVAR_ADDIN )
				{
					funcPtg = new PtgFuncVar( FunctionConstants.xlfADDIN, 0, form );
				}
				return funcPtg;
			}
		}
		return funcPtg;

	}

	/**
	 * take a string guaranteed to be a single Ptg (operator, reference, string, etc) and convert to correct Ptg
	 *
	 * @param form
	 * @param fmla
	 * @param bIsUnary -- operator is a unary version
	 * @return
	 */
	private static Ptg parseSinglePtg( XLSRecord form, String fmla, boolean bIsUnary )
	{
		WorkBook bk = form.getWorkBook();    // nec. to determine if parsed element is a valid name handle name

		String val = fmla;
		String name = convertString( val, bk );
		if( name.equals( "+" ) && bIsUnary )
		{
			name = "u+";
		}
		else if( name.equals( "-" ) && bIsUnary )
		{
			name = "u-";
		}

		Ptg pthing = null;
		try
		{
			pthing = XLSRecordFactory.getPtgRecord( name );
			if( (pthing == null) && name.equals( "PtgName" ) )
			// TODO: MUST evaluate which type of PtgName is correct: understand usage!
			{
				if( (form.getOpcode() == XLSConstants.FORMULA) || (form.getOpcode() == XLSConstants.ARRAY) )
				{
					pthing = new PtgName( 0x43 );    // assume this token to be of type Value (i.e PtgNameV) instead of Reference (PtgNameR)
				}
				else    // DV needs ref-type name
				{
					pthing = new PtgName( 0x23 );    // PtgNameR
				}
			}
		}
		catch( InvalidRecordException e )
		{
			log.warn( "parsing formula string.  Invalid Ptg: " + name + " error: " + e, e );
		}
		// if it is an operator we don't need to do anything with it!
		if( pthing != null )
		{
			pthing.setParentRec( form );

			if( !pthing.getIsOperator() )
			{
				if( pthing.getIsReference() )
				{
					// createPtgRefFromString will handle any type of string reference
					// will return a PtgRefErr if cannot parse location
					pthing = PtgRef.createPtgRefFromString( val, form );
				}
				else if( pthing instanceof PtgStr )
				{
					PtgStr pstr = (PtgStr) pthing;
					val = StringTool.strip( val, '\"' );
					pstr.setVal( val );
				}
				else if( pthing instanceof PtgNumber )
				{
					PtgNumber pnum = (PtgNumber) pthing;
					if( val.indexOf( "%" ) == -1 )
					{
						pnum.setVal( (new Double( val )) );
					}
					else
					{
						pnum.setVal( val );
					}
				}
				else if( pthing instanceof PtgInt )
				{
					PtgInt pint = (PtgInt) pthing;
					pint.setVal( (Integer.valueOf( val )) );
				}
				else if( pthing instanceof PtgBool )
				{
					PtgBool pbool = (PtgBool) pthing;
					pbool.setVal( Boolean.valueOf( val ) );
				}
				else if( pthing instanceof PtgArray )
				{
					PtgArray parr = (PtgArray) pthing;
					parr.setVal( val );
				}
				else if( pthing instanceof PtgName )
				{ // SHOULD really return PtgErr("#NAME!") as it's a missing Name instead of adding a new name
					PtgName pname = (PtgName) pthing;
					pname.setName( val );
					pname.addToRefTracker();
				}
				else if( pthing instanceof PtgNameX )
				{
					PtgNameX pnameX = (PtgNameX) pthing;
					pnameX.setName( val );
				}
				else if( pthing instanceof PtgMissArg )
				{
					((PtgMissArg) pthing).init( new byte[]{ 22 } );
				}
				else if( pthing instanceof PtgErr )
				{
					pthing = new PtgErr( PtgErr.convertStringToLookupByte( val ) );
				}
				else if( pthing instanceof PtgAtr )
				{
					pthing = new PtgAtr( (byte) 0x40 );    // assume space
				}
			}
		}
		else
		{
			PtgMissArg pname = new PtgMissArg();
		}
		return pthing;
	}

	private static String findPtg( String fmla, boolean bUnary )
	{
		String s = StringTool.allTrim( fmla );

		if( s.startsWith( "\"" ) || s.startsWith( "'" ) )
		{
			//return s;	// it's a string
			return s;
		}

		for( int i = 0; i < XLSRecordFactory.ptgOps.length; i++ )
		{
			String ptgOpStr = XLSRecordFactory.ptgOps[i][0];

			int x = s.indexOf( ptgOpStr );
			if( x == 0 )
			{ // found instance of an operator
				// if encounter a parenthesis, must determine if it is an expression limit OR
				// if it is part of a complex range, in which case the expression must be kept together
				if( ptgOpStr.equals( "(" ) )
				{
					return s;    // parens means a whole complex range==>PtgMemFunc
				}
				if( bUnary )
				{
					// unary ops +, - ... have a diff't Ptg than regular vers of the operator
					if( ptgOpStr.equals( "-" ) && (x == 1) && (ptgOpStr.length() > 1) )
					{
						break; //negative number, NOT a unary -
					}
					for( int j = 0; j < XLSRecordFactory.ptgPrefixOperators.length; j++ )
					{
						if( ptgOpStr.startsWith( XLSRecordFactory.ptgPrefixOperators[j][0].toString() ) )
						{
							ptgOpStr = XLSRecordFactory.ptgPrefixOperators[j][1].toString();
						}
					}
				}
				return XLSRecordFactory.ptgOps[i][1];
			}
		}
		return s;
	}

	/**
	 * parseFinalLevel is where strings get converted into ptg's
	 * This method can handle multiple ptg's within a string, but cannot handle
	 * recursion.  If you are having recursion problems look into getPtgsFromFormulaString above.
	 * <p/>
	 * This method should be called from the final level of parsing.
	 * There should not be additional sub-expressions at this point.
	 * Example (1+2) or (3,4,5)
	 * NOT ((1<2),3,4) or TAN(23);
	 * <p/>
	 * <p/>
	 * TODO: HANDLE these references:
	 * <p/>
	 * =SUM(table[[#This Row];['#Head3]:[Calced]])
	 * <p/>
	 * table[['#Head3]:[Calced]]
	 * <p/>
	 * I assume this means a table of data, the Head3 table? and the calced column?
	 */
	private static Stack parseFinalLevel( XLSRecord form, String fmla, boolean bIsComplete )
	{
		Stack returnStack = new Stack();
		CompatibleVector parseThings = new CompatibleVector();

		// break it up into components first
		Vector elements = new Vector();
		elements = splitString( fmla, bIsComplete );

		WorkBook bk = form.getWorkBook();    // nec. to determine if parsed element is a valid name handle name

		// convert each element into Ptg's
		// each element at this point should be a named operand, or an unidentified operator
		for( int x = 0; x < elements.size(); x++ )
		{
			String val = (String) elements.elementAt( x );
			String name = convertString( val, bk );

			Ptg pthing = null;
			try
			{
				pthing = XLSRecordFactory.getPtgRecord( name );
				if( (pthing == null) && name.equals( "PtgName" ) )
				// TODO: MUST evaluate which type of PtgName is correct: understand usage!
				{
					if( (form.getOpcode() == XLSConstants.FORMULA) || (form.getOpcode() == XLSConstants.ARRAY) )
					{
						pthing = new PtgName( 0x43 );    // assume this token to be of type Value (i.e PtgNameV) instead of Reference (PtgNameR)
					}
					else    // DV needs ref-type name
					{
						pthing = new PtgName( 0x23 );    // PtgNameR
					}
				}
			}
			catch( InvalidRecordException e )
			{
				log.warn( "parsing formula string.  Invalid Ptg: " + name + " error: " + e, e );
			}
			// if it is an operator we don't need to do anything with it!
			if( pthing != null )
			{
				pthing.setParentRec( form );

				if( !pthing.getIsOperator() )
				{
					if( pthing.getIsReference() )
					{
						// createPtgRefFromString will handle any type of string reference
						// will return a PtgRefErr if cannot parse location
						pthing = PtgRef.createPtgRefFromString( val, form );
					}
					else if( pthing instanceof PtgStr )
					{
						PtgStr pstr = (PtgStr) pthing;
						val = StringTool.strip( val, '\"' );
						pstr.setVal( val );
					}
					else if( pthing instanceof PtgNumber )
					{
						PtgNumber pnum = (PtgNumber) pthing;
						if( val.indexOf( "%" ) == -1 )
						{
							pnum.setVal( (new Double( val )) );
						}
						else
						{
							pnum.setVal( val );
						}
					}
					else if( pthing instanceof PtgInt )
					{
						PtgInt pint = (PtgInt) pthing;
						pint.setVal( (Integer.valueOf( val )) );
					}
					else if( pthing instanceof PtgBool )
					{
						PtgBool pbool = (PtgBool) pthing;
						pbool.setVal( Boolean.valueOf( val ) );
					}
					else if( pthing instanceof PtgArray )
					{
						PtgArray parr = (PtgArray) pthing;
						parr.setVal( val );
					}
					else if( pthing instanceof PtgName )
					{ // SHOULD really return PtgErr("#NAME!") as it's a missing Name instead of adding a new name
						PtgName pname = (PtgName) pthing;
						pname.setName( val );
					}
					else if( pthing instanceof PtgNameX )
					{
						PtgNameX pnameX = (PtgNameX) pthing;
						pnameX.setName( val );
					}
					else if( pthing instanceof PtgMissArg )
					{
						((PtgMissArg) pthing).init( new byte[]{ 22 } );
					}
					else if( pthing instanceof PtgErr )
					{
						pthing = new PtgErr( PtgErr.convertStringToLookupByte( val ) );
					}
				}
				parseThings.add( pthing );
			}
			else
			{
				PtgMissArg pname = new PtgMissArg();
			}

		}
		//reorder in polish notation and add to stack.
		// 20081128 KSC: Do later as reordering will depend upon position of this segment in formula returnStack = reorderStack(parseThings);  see getPtgsFromFormulaString
		returnStack = convertToStack( parseThings );
		return returnStack;
	}

	/*
    private static Stack reorderStack(Stack sourceStack, boolean bIsComplete) {
    	Stack returnStack = new Stack();
    	Stack pOperators= new Stack();
		for (int x = 0;x<sourceStack.size(); x++){
			Ptg pthing  = (Ptg)sourceStack.get(x);
			if (pthing.getIsOperator()){	// Must account for precedence of operators
				if (bIsComplete && !returnStack.isEmpty()) {
					// handle infrequent cases of two operators in a row i.e. one op and one unary op
					// e.g. 1--1
					while (pOperators.size() > 0) {	 
						returnStack.push(pOperators.pop());
					}				
  					if (((Ptg)returnStack.peek()).getIsOperator()) { 
						int precedence=  rankPrecedence(pthing);
						Ptg p= (Ptg) returnStack.pop();
						int prevprecedence= rankPrecedence(p);
				  		if (precedence>prevprecedence) {
							pOperators.push(p);	 // switch less precedence operator with greater
				  		}
				  		else
							returnStack.push(p); // back to normal
					}
				} else if (!pOperators.isEmpty()) {	// added for instances such as (2)^-(2); 
					//code below prevents the switching of the operators ^-
					Ptg p= (Ptg) pOperators.pop();
					pOperators.push(pthing);	// save operators
					pthing= p;
				} 
				pOperators.push(pthing);	// save operators
			}else{	// it's not an operand; put on stack and pop all operators thus far
				returnStack.push(pthing);
				while (pOperators.size() > 0) {
					returnStack.push(pOperators.pop());
				}				
			}
		}
		while (pOperators.size() > 0) {
			returnStack.push(pOperators.pop());
		}
		return returnStack;
    }
    */

	/**
	 * convert list of parseThings without reordering
	 *
	 * @param parseThings
	 * @return
	 */
	private static Stack convertToStack( CompatibleVector parseThings )
	{
		Stack returnStack = new Stack();
		Stack pOperators = new Stack();
		for( int x = 0; x < parseThings.size(); x++ )
		{
			Ptg pthing = (Ptg) parseThings.elementAt( x );
			returnStack.push( pthing );
		}
		return returnStack;
	}

	/*
	   * Parses an internal string for a function, splitting out elements.
	   * for instance,(1<2), 3, tan(5); should return
	   * [(1<2)][3][tan(5)].  Currently just working off of commas, but this may change...
	   * 
	   * One of the keys here is to not split on a comma from an internal function, for instance,
	   * "IF((1<2),MOD(45,6),0) should not split between 45 & 6!  Note the badLocs vector that handles this.
	   */
	private static CompatibleVector splitFunctionOperands( String formStr )
	{
		CompatibleVector locs = new CompatibleVector();
		// if there are no commas then we don't have to do all of this...
		if( formStr.equals( "" ) )
		{
			return locs;    // KSC: Handle no parameters by returning a null vector
		}

		// first handle quoted strings 20081111 KSC
		boolean loop = true;
		int pos = 0;
		CompatibleVector badLocs = new CompatibleVector();
		while( loop )
		{
			char c = '"';
			int start = formStr.indexOf( c, pos );
			if( start == -1 )
			{        // process single quotes as well
				c = '\'';
				start = formStr.indexOf( c, pos );
			}
			if( start != -1 )
			{
				int end = formStr.indexOf( c, start + 1 );
				end += 1;  //include trailing quote
				// check for being a part of a reference ...	
				if( (end < formStr.length()) && (formStr.charAt( end ) == '!') )
				{// then it's part of a reference
					end++;
					while( (end < formStr.length()) && loop )
					{
						c = formStr.charAt( end );
						if( !(Character.isLetterOrDigit( c ) || (c == ':') || (c == '$')) || (c == '-') || (c == '+') )
						{
							loop = false;
						}
						else
						{
							end++;
						}
					}
				}
				for( int y = start; y < end; y++ )
				{
					//make sure it is not a segment of a previous operand, like <> and >;
					badLocs.add( y );
				}
				if( end == 0 )
				{ // means it didn't find an end quote
					end = formStr.length() - 1;
					loop = false;
				}
				else
				{
					pos = end;
					loop = true;
				}
			}
			else
			{
				loop = false;
			}
		}

		if( formStr.indexOf( "," ) == -1 )
		{
			locs.add( formStr );
		}
		else
		{
			// Handle each parameter (delimited by ,) 	 
			// fill the badLocs vector with string locations we should disregard for comma proccesing
			for( int i = 0; i < formStr.length(); i++ )
			{
				int openparen = formStr.indexOf( "(", i );
				if( openparen != -1 )
				{
					if( !badLocs.contains( Integer.valueOf( openparen ) ) )
					{
						int closeparen = getMatchOperator( formStr, openparen, '(', ')' );
						if( closeparen == -1 )
						{
							closeparen = formStr.length();
						}
						for( i = openparen; i < closeparen; i++ )
						{
							Integer in = i;
							badLocs.add( in );
						}
					}
					else
					{    // open paren nested in quoted string
						i = openparen + 1;
					}
				}
				else    // 20081112 KSC
				{
					break;
				}
			}
			// lets do the same for the array items
			for( int i = 0; i < formStr.length(); i++ )
			{
				int openparen = formStr.indexOf( "{", i );
				if( openparen != -1 )
				{
					if( !badLocs.contains( Integer.valueOf( openparen ) ) )
					{
						int closeparen = getMatchOperator( formStr, openparen, '{', '}' );
						if( closeparen == -1 )
						{
							closeparen = formStr.length();
						}
						for( i = openparen; i < closeparen; i++ )
						{
							Integer in = i;
							badLocs.add( in );
						}
					}
					else
					{    // open paren nested in quoted string 
						i = openparen + 1;
					}
				}
				else    // 20081112 KSC
				{
					break;
				}
			}
			// now check bad locations:
			int placeholder = 0;
			int holder = 0;
			while( holder != -1 )
			{
				int i = formStr.indexOf( ",", holder );
				if( i != -1 )
				{
					Integer ing = i;
					if( !badLocs.contains( ing ) )
					{
						String s = formStr.substring( placeholder, i );
						locs.add( s );
						placeholder = i + 1;
					}
					holder = i + 1;
				}
				else
				{
					String s = formStr.substring( placeholder, formStr.length() );
					locs.add( s );
					return locs;
				}
			}
		}
		return locs;
	}

	/**
	 * parse a given string into known Ptg operators
	 *
	 * @param s
	 * @return
	 */
	private static CompatibleVector parsePtgOperators( String s, boolean bUnary )
	{
		CompatibleVector ret = new CompatibleVector();
		s = StringTool.allTrim( s );

		for( int i = 0; i < XLSRecordFactory.ptgOps.length; i++ )
		{
			String ptgOpStr = XLSRecordFactory.ptgOps[i][0];
			if( s.startsWith( "\"" ) || s.startsWith( "'" ) )
			{
				int end = s.substring( 1 ).indexOf( s.charAt( 0 ) );
				end += 1;  //include trailing quote
				// TEST IF The quoted item is a sheet name
				if( (end < s.length()) && (s.charAt( end ) == '!') )
				{// then it's part of a reference
					end++;
					boolean loop = true;
					while( (end < s.length()) && loop )
					{    // if the quoted string is a sheet ref, get rest of reference
						char c = s.charAt( end );
						if( (c == '#') && s.endsWith( "#REF!" ) )
						{
							end += 5;
							loop = false;
						}
						else if( !(Character.isLetterOrDigit( c ) || (c == ':') || (c == '$')) || (c == '-') || (c == '+') )
						{
							loop = false;
						}
						else
						{
							end++;
						}
					}
				}
				ret.add( s.substring( 0, end + 1 ) );
				s = s.substring( end + 1 );
				bUnary = false;
				if( !s.equals( "" ) )
				{
					ret.addAll( parsePtgOperators( s, bUnary ) );
				}
				break;
			}

			int x = s.indexOf( ptgOpStr );
			if( x > -1 )
			{    // found instance of an operator
				// if encounter a parenthesis, must determine if it is an expression limit OR
				// if it is part of a complex range, in which case the expression must be kept together
				if( ptgOpStr.equals( "(" ) )
				{
					int end = getMatchOperator( s, x, '(', ')' );
					ret.add( s );    // add entire
					break;
/*					String ss= s.substring(x, end+1);					
					ret.add(ss.substring(x));	// add entire 	
					if (FormulaParser.isComplexRange(ss)) {
//						ret.add(ss.substring(x+1));	// skip beginning paren as screws up later parsing
						ret.add(ss.substring(x));	
						s= s.substring(end+1);
						bUnary= false;
						if (!s.isEmpty())								
							ret.addAll(parsePtgOperators(s, bUnary));
						break;
					}
				} else if (ptgOpStr.equals(")")) { 
					try {
						String ss= s.substring(x+1);
						char nextChar= ss.charAt(0);
						if (nextChar==' ') {	// see if there is another operand after the space
							ss= ss.trim();
							if (ss.length()>0 && ss.matches("[^(a-zA-Z].*")) {
								nextChar= ss.charAt(0);
							}
						}
						// complex ranges can contain parentheses in combo with these operators: :, (
						if (nextChar==' ' || nextChar==',' || nextChar==':' || nextChar==')')
							continue;	// keep complex range expression together 						
					} catch (Exception e) { ; }
*/
				}
				if( ptgOpStr.equals( ")" ) )    // parens are there to keep expression together 
				{
					continue;
				}
				if( x > 0 )
				{// process prefix, if any - unary since it's the first operand
					// exception here-- error range in the form of "Sheet!#REF! (eg) needs to be kept whole
					if( !(XLSRecordFactory.ptgLookup[i][1].equals( "PtgErr" ) && (s.charAt( x - 1 ) == '!')) )
					{
						ret.addAll( parsePtgOperators( s.substring( 0, x ), bUnary ) );
						bUnary = false;
					}
					else
					{    // keep entire error reference together
						ptgOpStr = s;
					}
				}
				x = x + ptgOpStr.length();
				if( bUnary )
				{
					// unary ops +, - ... have a diff't Ptg than regular vers of the operator
					if( ptgOpStr.equals( "-" ) && (x == 1) && (ptgOpStr.length() > 1) )
					{
						break; //negative number, NOT a unary -
					}
					for( int j = 0; j < XLSRecordFactory.ptgPrefixOperators.length; j++ )
					{
						if( ptgOpStr.startsWith( XLSRecordFactory.ptgPrefixOperators[j][0].toString() ) )
						{
							ptgOpStr = XLSRecordFactory.ptgPrefixOperators[j][1].toString();
						}
					}
				}
				ret.add( ptgOpStr );
				if( x < s.length() ) // process suffix, if any
				{
					ret.addAll( parsePtgOperators( s.substring( x ), true ) );
				}
				break;
			}
		}
		if( ret.isEmpty() )
		{
			ret.add( s );
		}
		return ret;
	}

	/*
     * Parses a string and returns an array based on contents
     * Assumed to be 1 "final-level" operand i.e a range, complex range, a op b[...]   
     */
	private static CompatibleVector splitString( String formStr, boolean bIsComplete )
	{
		// Use a vector, and the collections methods to sort in natural order
		CompatibleVector locs = new CompatibleVector();
		CompatibleVector retVect = new CompatibleVector();

		// check for escaped string literals & add positions to vector if needed
		formStr = StringTool.allTrim( formStr );
		if( formStr.equals( "" ) )
		{
			retVect.add( formStr );
			return retVect;
		}
		if( true )
		{
			retVect.addAll( parsePtgOperators( formStr,
			                                   bIsComplete ) );        // cleanString if not an array formula????  s= cleanString(s);
			bIsComplete = false;
			return retVect;
		}

		// 20081207 KSC: redo completely to handle complex formula strings e.g. strings containing quoted commas, parens ... 
		// first, pre-process to parse quoted strings, parentheses and array formulas
		boolean isArray = false;
		boolean loop = true;
		String s = "";
		boolean inRange = false;
		char prevc = 0;
		for( int i = 0; i < formStr.length(); i++ )
		{
			char c = formStr.charAt( i );
			if( (c == '"') || (c == '\'') )
			{				
/*				if (!s.equals("")) { 
					locs.add(s);
					s= "";
				}
*/
				int end = formStr.indexOf( c, i + 1 );
				end += 1;  //include trailing quote
				// TEST IF The quoted item is a sheet name
				if( (end < formStr.length()) && (formStr.charAt( end ) == '!') )
				{// then it's part of a reference
					end++;
					loop = true;
					while( (end < formStr.length()) && loop )
					{    // if the quoted string is a sheet ref, get rest of reference
						c = formStr.charAt( end );
						if( (c == '#') && formStr.endsWith( "#REF!" ) )
						{
							end += 5;
							loop = false;
						}
						else if( !(Character.isLetterOrDigit( c ) || (c == ':') || (c == '$')) || (c == '-') || (c == '+') )
						{
							loop = false;
						}
						else
						{
							end++;
						}
					}
				}
				locs.add( s + formStr.substring( i, end ) );
				s += formStr.substring( i, end );
				i = end - 1;
			}
			else if( c == '(' )
			{    // may be a complex range if s==""
				if( !s.equals( "" ) && !inRange )
				{
//					char prevc= s.charAt(s.length()-1);
					if( !((prevc == ' ') || (prevc == ':') || (prevc == ',') || (prevc == '(')) )
					{
						locs.add( s );
						s = "";
					}
					else
					{    // DO NOT split apart complex ranges - they parse to PtgMemFuncs
						//Logger.logInfo("FormulaParser.splitString:  PtgMemFunc" + formStr);
						s += c;
						inRange = true;
					}
				}
			}
			else if( c == ':' )
			{
				if( (prevc == ')') && (locs.size() > 0) )    // complex range in style of: F(x):Y(x)
				{
					s = (String) locs.get( locs.size() - 1 ) + '(' + s;
				}
				inRange = true;
				s += c;
			}
			else if( c == '{' )
			{
				if( !s.equals( "" ) )
				{
					locs.add( s );
					s = "";
				}
				int end = formStr.indexOf( "}", i + 1 );
				end += 1;  //include trailing }
				locs.add( formStr.substring( i, end ) );
				i = end - 1;
			}
			else
			{
				s += c;
			}
			if( c != ' ' )
			{
				prevc = c;
			}
		}
		if( !s.equals( "" ) )
		{
			locs.add( s );
			s = "";
		}

		// loop through the possible operator ptg's and get locations & length of them		
		for( Object loc : locs )
		{
			s = (String) loc;
			if( s.startsWith( "\"" ) || s.startsWith( "'" ) )
			{
				retVect.add( s );    // quoted strings
			}
			else
			{
				if( s.startsWith( "{" ) )    // it's an array formula
				{
					isArray = true; // Do what?? else, cleanString??
				}
				retVect.addAll( parsePtgOperators( s, bIsComplete ) );        // cleanString if not an array formula????  s= cleanString(s);
			}
			bIsComplete = false;    // already parsed part of the formula string so cannot be unary :)
		}
		return retVect;
	}

	/**
	 * helper method that turns operator & operand strings into the Ptg\ equivalent
	 * if there is no equivalant it leaves the string alone.
	 */
	private static String convertString( String ptg, WorkBook bk )
	{
		// first check for operators
		for( int i = 0; i < XLSRecordFactory.ptgLookup.length; i++ )
		{
			if( ptg.equalsIgnoreCase( XLSRecordFactory.ptgLookup[i][0] ) )
			{
				return ptg;
			}
		}

		// KSC: Added for missing arguments ("")
		if( StringTool.allTrim( ptg ).equals( "" ) )
		//return "PtgMissArg";
		//if (ptg.equals(""))
		{
			return "PtgAtr";    // a space
		}

		// Now we need to figure out what type of operand it is
		// see if it is a string, should be encased by ""
		if( ptg.substring( 0, 1 ).equalsIgnoreCase( "\"" ) )
		{
			return "PtgStr";
		}
		// is it an array?
		if( ptg.substring( 0, 1 ).equalsIgnoreCase( "{" ) )
		{
			return "PtgArray";
		}
		// see if it is an integer
		if( ptg.indexOf( "." ) == -1 )
		{
			try
			{
				Integer i = Integer.valueOf( ptg );
				if( (i >= 0) && (i <= 65535) ) // PtgInts are UNSIGNED + <=65535
				{
					return "PtgInt";
				}
				return "PtgNumber";
			}
			catch( NumberFormatException e )
			{
			}
		}
		if( ptg.indexOf( "%" ) == (ptg.length() - 1) )
		{ // see if it's a percentage
			try
			{
				Double d = new Double( ptg.substring( 0, ptg.indexOf( "%" ) ) );
				return "PtgNumber";
			}
			catch( NumberFormatException e )
			{
			}
		}
		// see if it is a Number
		try
		{
			Double d = new Double( ptg );
			return "PtgNumber";
		}
		catch( NumberFormatException e )
		{
		}

		// at this point it is probably some sort of ptgref
		if( (ptg.indexOf( ":" ) != -1) || (ptg.indexOf( ',' ) != -1) || (ptg.indexOf( "!" ) != -1) )
		{
			// ptgarea or ptgarea3d or ptgmemfunc
			return "PtgArea"; // in ParseFinalLevel, PtgRef.createPtgRefFromString will handle all types of string refs
		}

		// maybe it is a garbage string, or a reference to a name (unsupported right now....)
		// check if the last character is a number. If not, it sure isn't a reference, no?
		// NO!  Can have named ranges with numbers at the end -- better to try to parse it
		try
		{
			if( bk.getName( ptg ) == null )
			{// it's not a named range
				ExcelTools.getRowColFromString( ptg );    // if passes it's a PtgRef
				return "PtgRef";
			}
			return "PtgName";
		}
		catch( IllegalArgumentException e )
		{
			return "PtgName";
		}
	}

	/*
     * helper method that cleans out unneccesary parts of the formula string.
     */
	private static String cleanString( String dirtystring )
	{
		String cleanstring = StringTool.allTrim( dirtystring );
		cleanstring = StringTool.strip( cleanstring, "(" );
		cleanstring = StringTool.strip( cleanstring, "," );
		return cleanstring;
	}

	protected static Stack getPtgsFromFormulaString( String fmla )
	{
		return null;

	}

	/**
	 * parse a formula in string form and create a formula record from it
	 * caluclate the new formula based on boolean setting calculate
	 *
	 * @param form      Formula rec
	 * @param fmla      String formula either =EXPRESSION or {=EXPRESSION} for array formulas
	 * @param calculate boolean truth of "calculate formula after setting"
	 * @return Formula rec
	 */
	public static Formula setFormula( Formula form, String fmla, int[] rc )
	{
		if( fmla.charAt( 0 ) != '{' )
		{
			try
			{
				Stack newptgs = FormulaParser.getPtgsFromFormulaString( form, fmla );
				FormulaParser.adjustParameterIds( newptgs );     // 20100614 KSC: adjust function parameter id's, if necessary, for Value, Array or Reference type
				form.setExpression( newptgs );
			}
			catch( FunctionNotSupportedException e )
			{  // 200902 KSC: still add record if function is not found (using N/A in place of said function)
				log.error( "Adding new Formula at " + form.getSheet() + "!" + ExcelTools.formatLocation( rc ) + " failed: " + e.toString() + "." );
				Stack newptgs = new Stack();
				newptgs.push( new PtgErr( PtgErr.ERROR_NA ) );
				form.setExpression( newptgs );
			}
		}
		else
		{ // Handle Array Formulas
			PtgExp pe = new PtgExp();
			pe.setParentRec( form );
			// rowcol reference is from PARENT PtgExp not (necessarily) this formula's cell address
			// [BugTracker 2683 + OOXML Array Formulas]
			Object o = form.getSheet().getArrayFormulaParent( rc );
			if( o != null )    // there is a parent array formula; use it's rowcol
			{
				rc = (int[]) o;
			}
			else
			{            // no parent yet- add
				String addr = ExcelTools.formatLocation( rc );
				form.getSheet().addParentArrayRef( addr, addr );
			}
			pe.init( rc[0], rc[1] );
			Stack e = new Stack();
			e.push( pe );
			FormulaParser.adjustParameterIds( e );     // adjust function parameter id's, if necessary, for Value, Array or Reference type
			form.setExpression( e );    // add PtgExp to Formula Stack
			Array a = new Array();    // Create new Array Record
			a.setSheet( form.getSheet() );
			a.setWorkBook( form.getWorkBook() );
			a.init( fmla, rc[0], rc[1] );    // init Array record from Formula String
			form.addInternalRecord( a );    // link array record to parent formula
		}
		
    	/* is this calc necessary?
    	Object val = null;
    	try{
    	    if (form.getWorkBook().getCalcMode() == WorkBook.CALCULATE_ALWAYS)
    	    	val = form.calculateFormula();			
    	}catch (Exception e){
      		Logger.logWarn("Unsupported Function: " + e + ".  ExtenXLS calculation will be unavailable: " + fmla);	//20081118 KSC: display a little more info ... 
    	}
        if(DEBUGLEVEL > 0)Logger.logInfo("FormulaParser.setFormula() string:" +fmla + " value: " + val);
        */

		return form;
	}

	public static String getFormulaString( Formula form )
	{
		Stack expression = form.getExpression();
		return FormulaParser.getExpressionString( expression );
	}

	public static String getExpressionString( Stack expression )
	{
		StringBuffer retval = new StringBuffer( "" );
		int sz = expression.size();
		Ptg[] stck = new Ptg[sz];
		stck = (Ptg[]) expression.toArray( stck );
		Stack newstck = new Stack();
		for( int t = 0; t < sz; t++ )
		{ // flip the stack 
			newstck.add( 0, stck[t] );
		}
		Stack vals = new Stack();
		while( !newstck.isEmpty() )
		{
			handlePtg( newstck, vals );
		}
		String s = "";
		while( !vals.isEmpty() )
		{
			Ptg topP = (Ptg) vals.pop();
			s = topP.getTextString() + s;
		}
		retval = new StringBuffer( "=" + s );
		return retval.toString();
	}

	/**
	 * adjustParameterIds pre-processes the function expression stack,
	 * analyzing function parameters to ensure that function operands
	 * contain the correct PtgID (which indicates Value, Reference or Array)
	 * <br>
	 * Reference-type ptg's contain an ID which indicates the type
	 * required by the calling function: Value, Reference or Array.
	 * <br>
	 * When formula strings are parsed into formula expressions,
	 * ptg reference-type parameters are assigned a default id,
	 * but this id may not be correct for all functions.
	 */
	public static void adjustParameterIds( Stack expression )
	{
		StringBuffer retval = new StringBuffer( "" );
		int sz = expression.size();
		Ptg[] stck = new Ptg[sz];
		stck = (Ptg[]) expression.toArray( stck );
		Stack newstck = new Stack();
		for( int t = 0; t < sz; t++ )
		{ // flip the stack 
			newstck.add( 0, stck[t] );
		}
		Stack params = new Stack();
		// we only care about PtgFuncVar and PtgFunc's but need to process the expression
		// stack thoroughly to get the correct parameters
		// Process function stack, gathering parameters.  When we have all the parameters
		// for a PtgFunc or a PtgFuncVar, adjust any PtgRef types, if necessary
		while( !newstck.isEmpty() )
		{
			Ptg p = (Ptg) newstck.pop();
			int x = 0;// cargs = p.getNumParams();
			int t = 0;
			if( p.getIsControl() )
			{
				// do the parens thing here...
				if( p.getOpcode() == 0x15 )
				{ // its a parens... and there is a val 
					if( t > 0 )
					{
						// 20060128 - KSC: handle parens 
						Ptg[] vx = new Ptg[1];    // parens are unary ops so only 1 var allowed
						vx[0] = (Ptg) params.pop();
						p.setVars( vx );
						params.push( p );    // put paren (with var) back on stack
					}
					else
					{ // this paren wraps other parens...
						params.push( p );
					}
				}
			}
			else if( p.getIsOperator() || p.getIsFunction() )
			{
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
				}// it's a ptgfuncvar
				if( (p.getOpcode() == 0x21) || (p.getOpcode() == 0x41) || (p.getOpcode() == 0x61) )
				{
					t = p.getNumParams();
				}// it's a ptgFunc

				if( t > params.size() )
				{
					t = params.size();
				}
				Ptg[] vx = new Ptg[t];
				while( t > 0 )
				{
					vx[--t] = (Ptg) params.pop();// get'em
				}
				p.setVars( vx );  // set'em
				// here is where we adjust the ptg's of the func or funcvar parameters
				if( (p.getOpcode() == 0x22) || (p.getOpcode() == 0x42) || (p.getOpcode() == 0x62) ) /* it's a ptgfuncvar*/
				{
					((PtgFuncVar) p).adjustParameterIds();
				}
				else if( (p.getOpcode() == 0x21) || (p.getOpcode() == 0x41) || (p.getOpcode() == 0x61) )/* it's a ptgFunc */
				{
					((PtgFunc) p).adjustParameterIds();
				}
				params.push( p );// push it back on the stack
			}
			else if( p.getIsOperand() )
			{
				params.push( p );
			}
		}
	}

	/**
	 * set up the Formula chain
	 * <p/>
	 * A5
	 * 10
	 * <p/>
	 * 3
	 * EXP(
	 * +
	 * SUM(
	 * =SUM(A5*10+EXP(3))
	 * =SUM(EXP(A5*103))
	 * <p/>
	 * A3      add to vals
	 * E5      add to vals
	 * +       pop last2 vals add to vals
	 * (       check last -- if oper add oper, else add last, add to vals
	 * <p/>
	 * A1:A2   at this point we should see: (A3+E5) push to vals
	 * SUM(    check last -- if oper add oper, else add last, add to vals
	 * +       pop last2 add to vals
	 * SUM(    check last -- if oper add oper, else add last, add to vals
	 * =SUM((A3+E5)+SUM(A1:A2))
	 * =SUM((A3+E5)+SUM(A1:A2))
	 * <p/>
	 * A1      add to vals
	 * A2      add to vals
	 * +       pop last2 vals add to vals?
	 * (       check last -- if oper add oper, else add last, add to vals
	 * A3      add to vals
	 * A4      add to vals
	 * +       check last -- if paren pop last2 vals add to vals?
	 * (       check last -- if oper add oper, else add last, add to vals
	 * /       pop last2 vals add to vals
	 * <p/>
	 * =(A1+A2)/(A3+A4)
	 * =(A1+A2)/(A3+A4)
	 * <p/>
	 * <p/>
	 * PtgFunc taking 3 vals again can only have 1
	 * We know IF has 3 vals
	 * Do we need logic which knows how many vars a Ptg takes? (we sure do... -nick) Might help.
	 * <p/>
	 * ---> WRITE CODE TO SWITCH ON NUMBER OF PARAMS.  Should Fix.
	 * <p/>
	 * <p/>
	 * =IF(SUM(CONCATENATE(SUM((EXP(C2,D4,4))*2SUM(A5,C7,A2)A1:A5))=1),SUM(3),SUM(22))
	 * =IF(CONCATENATE(C2,(D4+EXP(4))*2,SUM(A5,C7,A2),SUM(A1:A5))=1,3,22)
	 */
	static void handlePtg( Stack newstck, Stack vals )
	{
		Ptg p = (Ptg) newstck.pop();
		int x = 0;// cargs = p.getNumParams();
		int t = 0;
		if( p.getIsOperator() || p.getIsControl() || p.getIsFunction() )
		{
			t = vals.size();   //this is faulty logic.  We don't care what is there, the operator should tell us.
			// do the parens thing here...
			if( p.getIsControl() /* !vals.isEmpty()*/ )
			{
				if( p.getOpcode() == 0x15 )
				{ // its a parens... and there is a val 
					if( t > 0 )
					{
						// 20060128 - KSC: handle parens 
						Ptg[] vx = new Ptg[1];    // parens are unary ops so only 1 var allowed
						vx[0] = (Ptg) vals.pop();
						p.setVars( vx );
						vals.push( p );    // put paren (with var) back on stack
					}
					else
					{ // this paren wraps other parens...
						vals.push( p );
					}
					return;
				}
			}
			if( t > 0 )
			{
				// make sure we have the correct amount popped back in..
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
				}// it's a ptgfuncvar!
				if( (p.getOpcode() == 0x21) || (p.getOpcode() == 0x41) || (p.getOpcode() == 0x61) )
				{
					t = p.getNumParams();
				}// it's a ptgFunc

				if( t > vals.size() )
				{
					// is this a real error? throw an exception?
						log.warn( "FormulaParser.handlePtg: number of parameters " + t + " is greater than available " + vals.size() );
					t = vals.size();
				}
				Ptg[] vx = new Ptg[t];
				while( t > 0 )
				{
					vx[--t] = (Ptg) vals.pop();// get'em
				}
				p.setVars( vx );  // set'em
			}
			vals.push( p );// push it back on the stack
		}
		else if( p.getIsOperand() )
		{
			vals.push( p );
		}
		else if( p instanceof PtgAtr )
		{
			// this is probably just a space at this point, don't output error message
		}
		else
		{
				log.debug( "FormulaParser Error - Ptg Type: " + p.getOpcode() + " " + p.getString() );
		}
	}

	/**
	 * create a new formula record at row column rc using formula string formStr
	 *
	 * @param formStr String
	 * @param st
	 * @param rc      int[]
	 * @return new Formula record
	 * @throws Exception
	 */
	public static Formula getFormulaFromString( String formStr, Boundsheet st, int[] rc ) throws Exception
	{
		Formula f = new Formula();
		if( st != null )
		{
			f.setSheet( st );
			f.setWorkBook( st.getWorkBook() );
		}
		f.setData( new byte[6] );    // necessary for setRowCol		
		f.setRowCol( rc );        // do before calculateFormula as array formulas use rowcol 20090817 KSC: [BugTracker 2683 + OOXML Array Formulas]
		f = FormulaParser.setFormula( f, formStr, rc );

		return f;
	}

	/**
	 * create a new formula record at row column rc using formula string formStr
	 *
	 * @param formStr String
	 * @param rc      int[]
	 * @return new Formula record
	 * @throws Exception
	 */
	public static Formula setFormulaString( String formStr, int[] rc ) throws /* 20070212 KSC: FunctionNotSupported*/Exception
	{
		Formula f = new Formula();
		f = FormulaParser.setFormula( f, formStr, rc );
		return f;
	}

	/**
	 * returns true of string s is in the form of a basic reference e.g. A1
	 *
	 * @param s
	 * @return
	 */
	public static boolean isRef( String s )
	{
		if( s == null )
		{
			return false;
		}
		String simpleOne = "(([ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})";
		return s.matches( simpleOne );
	}

	/**
	 * returns true if the stirng in question is in the form of a range
	 *
	 * @param s
	 * @return
	 */
	public static boolean isRange( String s )
	{
		if( s == null )
		{
			return false;
		}
		// TODO PRECOMPILE THESE
		String one = "(([(]*[ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+[)]*){1})";
		String aRange = one + "(:" + one + ")?";
		String rangeop = "([ ]*[: ,][ ]*)";
		String rangeMatchString = aRange + rangeop + aRange + "(" + rangeop + aRange + ")*";
		String simpleOne = "(([ ]*[']?([a-zA-Z0-9 ]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})";
		String simpleRangeMatchString = "(" + simpleOne + "[ ]*[:][ ]*" + simpleOne + ")";
		return (s.matches( rangeMatchString ));
	}

	/**
	 * returns true if the string represents a complex range (i.e. one containing multiple range values separated by one or more of: , : or space
	 *
	 * @param s
	 * @return
	 */
	public static boolean isComplexRange( String s )
	{
		// TODO PRECOMPILE THESE
		String one = "(([(]*[ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+[)]*){1})";
		String aRange = one + "(:" + one + ")?";
		String rangeop = "([ ]*[: ,][ ]*)";
		String rangeMatchString = aRange + rangeop + aRange + "(" + rangeop + aRange + ")*";
		String simpleOne = "(([ ]*[']?([a-zA-Z0-9 _]*[']*[!])?[$]*[a-zA-Z]{1,2}[$]*[0-9]+){1})";
		String simpleRangeMatchString = "(" + simpleOne + "[ ]*[:][ ]*" + simpleOne + ")";
		return (isRange( s )) && !(s.matches( simpleRangeMatchString ));
	}

}