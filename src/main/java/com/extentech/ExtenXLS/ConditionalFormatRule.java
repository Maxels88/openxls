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
package com.extentech.ExtenXLS;

import com.extentech.formats.XLS.Cf;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgRef;

import java.awt.*;

/**
 * ConditionalFormatRule defines a single rule for
 * manipulation of the ConditionalFormat cells in Excel
 * <p/>
 * Each ConditionalFormatRule contains one rule and corresponding format data.
 *
 * @see ConditionalFormatHandle
 * @see Handle
 */
public class ConditionalFormatRule implements Handle
{

	Cf currentCf = null;
	// static shorts for setting ConditionalFormat type
	public static final byte VALUE_ANY = 0x0;
	public static final byte VALUE_INTEGER = 0x1;
	public static final byte VALUE_DECIMAL = 0x2;
	public static final byte VALUE_USER_DEFINED_LIST = 0x3;
	public static final byte VALUE_DATE = 0x4;
	public static final byte VALUE_TIME = 0x5;
	public static final byte VALUE_TEXT_LENGTH = 0x6;
	public static final byte VALUE_FORMULA = 0x7;

	// static shorts for setting action on error
	public static byte ERROR_STOP = 0x0;
	public static byte ERROR_WARN = 0x1;
	public static byte ERROR_INFO = 0x2;

	// static shorts for setting conditions on ConditionalFormat
	public static final byte CONDITION_BETWEEN = 0x0;
	public static final byte CONDITION_NOT_BETWEEN = 0x1;
	public static final byte CONDITION_EQUAL = 0x2;
	public static final byte CONDITION_NOT_EQUAL = 0x3;
	public static final byte CONDITION_GREATER_THAN = 0x4;
	public static final byte CONDITION_LESS_THAN = 0x5;
	public static final byte CONDITION_GREATER_OR_EQUAL = 0x6;
	public static final byte CONDITION_LESS_OR_EQUAL = 0x7;

	public static String[] OPERATORS = {
			"nocomparison",
			"between",
			"notBetween",
			"equal",
			"notEqual",
			"greaterThan",
			"lessThan",
			"greaterOrEqual",
			"lessOrEqual",
			"beginsWith",
			"endsWith",
			"containsText",
			"notContains"
	};

	/**
	 * Get the byte representing the condition type string passed in.
	 * Options are'
	 * "between",
	 * "notBetween",
	 * "equal",
	 * "notEqual",
	 * "greaterThan",
	 * "lessThan",
	 * "greaterOrEqual",
	 * "lessOrEqual"
	 *
	 * @return
	 */
	public static byte getConditionNumber( String conditionType )
	{
		for( int i = 0; i < OPERATORS.length; i++ )
		{
			if( conditionType.equalsIgnoreCase( OPERATORS[i] ) )
			{
				return (byte) i;
			}
		}
		return -1;
	}

	public static String[] VALUE_TYPE = {
			"any", "integer", "decimal", "userDefinedList", "date", "time", "textLength", "formula"
	};

	/**
	 * Get the byte representing the value type string passed in.
	 * Options are'
	 * "any",
	 * "integer",
	 * "decimal",
	 * "userDefinedList",
	 * "date",
	 * "time",
	 * "textLength",
	 * "formula"
	 *
	 * @return
	 */
	public static byte getValueNumber( String valueType )
	{
		for( int i = 0; i < VALUE_TYPE.length; i++ )
		{
			if( valueType.equalsIgnoreCase( VALUE_TYPE[i] ) )
			{
				return (byte) i;
			}
		}
		return -1;
	}

	/**
	 * Create a conditional format rule from a Cf record.
	 *
	 * @param theCf
	 */
	protected ConditionalFormatRule( Cf theCf )
	{
		currentCf = theCf;
	}

	/**
	 * evaluates the criteria for this Conditional Format Rule
	 * <br>if the criteria involves a comparison i.e. equals, less than, etc., it uses the
	 * value from the passed in referenced cell to compare with
	 *
	 * @param Ptg refcell - the Ptg location to obtain cell value from
	 * @return boolean true if evaluation of criteria passes
	 * @see com.extentech.formats.XLS.Cf#evaluate(com.extentech.formats.XLS.formulas.Ptg)
	 */
	public boolean evaluate( CellHandle refcell )
	{
		Ptg/*Ref*/ pr = PtgRef.createPtgRefFromString( refcell.getCellAddress(), null );
		return currentCf.evaluate( pr );
	}

	/**
	 * Get the type operator of this ConditionalFormat as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in
	 * ConditionalFormatHandle
	 *
	 * @return
	 */
	public byte getTypeOperator()
	{
		return (byte) currentCf.getOperator();
	}

	/**
	 * set the type operator of this ConditionalFormat as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in
	 * ConditionalFormatHandle
	 */
	public void setTypeOperator( byte typOperator )
	{
		currentCf.setOperator( typOperator );
	}

	/**
	 * Get the second condition of the ConditionalFormat as
	 * a string representation
	 *
	 * @return
	 */
	public String getSecondCondition()
	{
		if( (currentCf != null) && (currentCf.getFormula2() != null) )
		{
			return currentCf.getFormula2().getFormulaString();
		}
		return null;
	}

	/**
	 * Set the second condition of the ConditionalFormat utilizing
	 * a string.  This value must conform to the Value Type of this
	 * ConditionalFormat or unexpected results may occur.  For example,
	 * entering a string representation of a date here will not work
	 * if your ConditionalFormat is an integer...
	 * <p/>
	 * String passed in should be a vaild XLS formula.  Does not need to include the "="
	 * <p/>
	 * Types of conditions
	 * Integer values
	 * Decimal values
	 * User defined list
	 * Date
	 * Time
	 * Text length
	 * Formula
	 * <p/>
	 * Be sure that your ConditionalFormat type (getConditionalFormatType()) matches the type of data.
	 *
	 * @return
	 */
	public void setSecondCondition( Object secondCond )
	{
		String setval = secondCond.toString();
		if( secondCond instanceof java.util.Date )
		{
			double d = DateConverter.getXLSDateVal( (java.util.Date) secondCond );
			setval = d + "";
		}
		currentCf.setCondition2( setval );
	}

	/**
	 * retrieves the border colors for the current Conditional Format
	 *
	 * @return java.awt.Color array of Color objects for each border side (Top, Left, Bottom, Right)
	 * @see com.extentech.formats.XLS.Cf#getBorderColors()
	 */
	public Color[] getBorderColors()
	{
		return currentCf.getBorderColors();
	}

	/**
	 * returns the bottom border line color for the current Conditional Format
	 *
	 * @return int bottom border line color constant
	 * @see com.extentech.formats.XLS.Cf#getBorderLineColorBottom()
	 * @see FormatHandle.COLOR_* constants
	 */
	public int getBorderLineColorBottom()
	{
		return currentCf.getBorderLineColorBottom();
	}

	/**
	 * returns the left border line color for the current Conditional Format
	 *
	 * @return int left border line color constant
	 * @see com.extentech.formats.XLS.Cf#getBorderLineColorLeft()
	 * @see FormatHandle.COLOR_* constants
	 */
	public int getBorderLineColorLeft()
	{
		return currentCf.getBorderLineColorLeft();
	}

	/**
	 * returns the right border line color for the current Conditional Format
	 *
	 * @return int right border line color constant
	 * @see com.extentech.formats.XLS.Cf#getBorderLineColorRight()
	 * @see FormatHandle.COLOR_* constants
	 */
	public int getBorderLineColorRight()
	{
		return currentCf.getBorderLineColorRight();
	}

	/**
	 * returns the top border line color for the current Conditional Format
	 *
	 * @return int top border line color constant
	 * @see com.extentech.formats.XLS.Cf#getBorderLineColorTop()
	 * @see FormatHandle.COLOR_* constants
	 */
	public int getBorderLineColorTop()
	{
		return currentCf.getBorderLineColorTop();
	}

	/**
	 * returns the bottom border line style for the current Conditional Format
	 *
	 * @return int bottom border line style constant
	 * @see com.extentech.formats.XLS.Cf#getBorderLineStylesBottom()
	 * @see FormatHandle.BORDER* line style constants
	 */
	public int getBorderLineStylesBottom()
	{
		return currentCf.getBorderLineStylesBottom();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getBorderLineStylesLeft()
	 */
	public int getBorderLineStylesLeft()
	{
		return currentCf.getBorderLineStylesLeft();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getBorderLineStylesRight()
	 */
	public int getBorderLineStylesRight()
	{
		return currentCf.getBorderLineStylesRight();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getBorderLineStylesTop()
	 */
	public int getBorderLineStylesTop()
	{
		return currentCf.getBorderLineStylesTop();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getBorderSizes()
	 */
	public int[] getBorderSizes()
	{
		return currentCf.getBorderSizes();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getBorderStyles()
	 */
	public int[] getBorderStyles()
	{
		return currentCf.getBorderStyles();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFont()
	 */
	public Font getFont()
	{
		return currentCf.getFont();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontColorIndex()
	 */
	public int getFontColorIndex()
	{
		return currentCf.getFontColorIndex();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontEscapement()
	 */
	public int getFontEscapement()
	{
		return currentCf.getFontEscapement();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontHeight()
	 */
	public int getFontHeight()
	{
		return currentCf.getFontHeight();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontOptsCancellation()
	 */
	public int getFontOptsCancellation()
	{
		return currentCf.getFontOptsCancellation();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontOptsPosture()
	 */
	public int getFontOptsPosture()
	{
		return currentCf.getFontOptsPosture();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontUnderlineStyle()
	 */
	public int getFontUnderlineStyle()
	{
		return currentCf.getFontUnderlineStyle();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFontWeight()
	 */
	public int getFontWeight()
	{
		return currentCf.getFontWeight();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getForegroundColor()
	 */
	public int getForegroundColor()
	{
		return currentCf.getForegroundColor();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.XLSRecord#getFormatPattern()
	 */
	public String getFormatPattern()
	{
		return currentCf.getFormatPattern();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFormula1()
	 */
	public Formula getFormula1()
	{
		return currentCf.getFormula1();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getFormula2()
	 */
	public Formula getFormula2()
	{
		return currentCf.getFormula2();
	}

	/**
	 * returns the operator for this Conditional Format Rule
	 * <br>e.g. "bewteen", "greater than" ...
	 *
	 * @return
	 */
	public String getOperator()
	{
		int op = currentCf.getOperator();
		if( (op >= 0) && (op < OPERATORS.length) )
		{
			return OPERATORS[op];
		}
		return "unknown operator: " + op;
	}

	/**
	 * returns the type of this Conditional Format
	 * <br>e.g. "Cell value is" or "Formula value is"
	 *
	 * @return String Conditional Format Type
	 */
	public String getType()
	{
		return currentCf.getTypeString();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getPatternFillColor()
	 */
	public int getPatternFillColor()
	{
		return currentCf.getPatternFillColor();
	}

	/**
	 * returns the pattern color, if any, as an HTML color String.  Includes custom OOXML colors.
	 *
	 * @return String HTML Color String
	 */
	public String getPatternFgColor()
	{
		return currentCf.getPatternFgColor();
	}

	/**
	 * returns the pattern color, if any, as an HTML color String.  Includes custom OOXML colors.
	 *
	 * @return String HTML Color String
	 */
	public String getPatternBgColor()
	{
		return currentCf.getPatternBgColor();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getPatternFillColorBack()
	 */
	public int getPatternFillColorBack()
	{
		return currentCf.getPatternFillColorBack();
	}

	/**
	 * @return
	 * @see com.extentech.formats.XLS.Cf#getPatternFillStyle()
	 */
	public int getPatternFillStyle()
	{
		return currentCf.getPatternFillStyle();
	}

	/**
	 * @param borderLineColorBottom
	 * @see com.extentech.formats.XLS.Cf#setBorderLineColorBottom(int)
	 */
	public void setBorderLineColorBottom( int borderLineColorBottom )
	{
		currentCf.setBorderLineColorBottom( borderLineColorBottom );
	}

	/**
	 * @param borderLineColorLeft
	 * @see com.extentech.formats.XLS.Cf#setBorderLineColorLeft(int)
	 */
	public void setBorderLineColorLeft( int borderLineColorLeft )
	{
		currentCf.setBorderLineColorLeft( borderLineColorLeft );
	}

	/**
	 * @param borderLineColorTop
	 * @see com.extentech.formats.XLS.Cf#setBorderLineColorTop(int)
	 */
	public void setBorderLineColorTop( int borderLineColorTop )
	{
		currentCf.setBorderLineColorTop( borderLineColorTop );
	}

	/**
	 * @param borderLineStylesBottom
	 * @see com.extentech.formats.XLS.Cf#setBorderLineStylesBottom(int)
	 */
	public void setBorderLineStylesBottom( int borderLineStylesBottom )
	{
		currentCf.setBorderLineStylesBottom( borderLineStylesBottom );
	}

	/**
	 * @param borderLineStylesLeft
	 * @see com.extentech.formats.XLS.Cf#setBorderLineStylesLeft(int)
	 */
	public void setBorderLineStylesLeft( int borderLineStylesLeft )
	{
		currentCf.setBorderLineStylesLeft( borderLineStylesLeft );
	}

	/**
	 * @param borderLineStylesRight
	 * @see com.extentech.formats.XLS.Cf#setBorderLineStylesRight(int)
	 */
	public void setBorderLineStylesRight( int borderLineStylesRight )
	{
		currentCf.setBorderLineStylesRight( borderLineStylesRight );
	}

	/**
	 * @param borderLineStylesTop
	 * @see com.extentech.formats.XLS.Cf#setBorderLineStylesTop(int)
	 */
	public void setBorderLineStylesTop( int borderLineStylesTop )
	{
		currentCf.setBorderLineStylesTop( borderLineStylesTop );
	}

	/**
	 * @param fontColorIndex
	 * @see com.extentech.formats.XLS.Cf#setFontColorIndex(int)
	 */
	public void setFontColorIndex( int fontColorIndex )
	{
		currentCf.setFontColorIndex( fontColorIndex );
	}

	/**
	 * @param fontEscapementFlag
	 * @see com.extentech.formats.XLS.Cf#setFontEscapement(int)
	 */
	public void setFontEscapement( int fontEscapementFlag )
	{
		currentCf.setFontEscapement( fontEscapementFlag );
	}

	/**
	 * @param fontHeight
	 * @see com.extentech.formats.XLS.Cf#setFontHeight(int)
	 */
	public void setFontHeight( int fontHeight )
	{
		currentCf.setFontHeight( fontHeight );
	}

	/**
	 * @param fontOptsCancellation
	 * @see com.extentech.formats.XLS.Cf#setFontOptsCancellation(int)
	 */
	public void setFontOptsCancellation( int fontOptsCancellation )
	{
		currentCf.setFontStriken( (fontOptsCancellation == Cf.FONT_OPTIONS_CANCELLATION_ON) );
	}

	/**
	 * @param fontOptsPosture
	 * @see com.extentech.formats.XLS.Cf#setFontOptsPosture(int)
	 */
	public void setFontOptsPosture( int fontOptsPosture )
	{
		currentCf.setFontOptsPosture( fontOptsPosture );
	}

	/**
	 * @param fontUnderlineStyle
	 * @see com.extentech.formats.XLS.Cf#setFontUnderlineStyle(int)
	 */
	public void setFontUnderlineStyle( int fontUnderlineStyle )
	{
		currentCf.setFontUnderlineStyle( fontUnderlineStyle );
	}

	/**
	 * @param fontWeight
	 * @see com.extentech.formats.XLS.Cf#setFontWeight(int)
	 */
	public void setFontWeight( int fontWeight )
	{
		currentCf.setFontWeight( fontWeight );
	}

	/**
	 * @param patternFillColor
	 * @see com.extentech.formats.XLS.Cf#setPatternFillColor(int)
	 */
	public void setPatternFillColor( int patternFillColor )
	{
		currentCf.setPatternFillColor( patternFillColor, null );
	}

	/**
	 * @param patternFillColorBack
	 * @see com.extentech.formats.XLS.Cf#setPatternFillColorBack(int)
	 */
	public void setPatternFillColorBack( int patternFillColorBack )
	{
		currentCf.setPatternFillColorBack( patternFillColorBack );
	}

	/**
	 * @param patternFillStyle
	 * @see com.extentech.formats.XLS.Cf#setPatternFillStyle(int)
	 */
	public void setPatternFillStyle( int patternFillStyle )
	{
		currentCf.setPatternFillStyle( patternFillStyle );
	}

	/**
	 * Get the first condition of the ConditionalFormat as
	 * a string representation
	 *
	 * @return
	 */
	public String getFirstCondition()
	{
		return currentCf.getFormula1().getFormulaString();
	}

	/**
	 * Set the first condition of the ConditionalFormat
	 * <p/>
	 * This value must conform to the Value Type of this
	 * ConditionalFormat or unexpected results may occur.  For example,
	 * entering a string representation of a date here will not work
	 * if your ConditionalFormat is an integer...
	 * <p/>
	 * A java.util.Date object can also be passed in.  This value will be translated
	 * into an integer as excel stores dates.  If you need to manipulate/retrieve this value
	 * later utilize the DateConverter tool to transform the value
	 * <p/>
	 * String passed in should be a vaild XLS formula.  Does not need to include the "="
	 * <p/>
	 * Types of conditions
	 * Integer values
	 * Decimal values
	 * User defined list
	 * Date
	 * Time
	 * Text length
	 * Formula
	 * <p/>
	 * Be sure that your ConditionalFormat type (getConditionalFormatType()) matches the type of data.
	 *
	 * @param firstCond = the first condition for the ConditionalFormat
	 */
	public void setFirstCondition( Object firstCond )
	{
		String setval = firstCond.toString();
		if( firstCond instanceof java.util.Date )
		{
			double d = DateConverter.getXLSDateVal( (java.util.Date) firstCond );
			setval = d + "";
		}
		currentCf.setCondition1( setval );
	}

	public static int getConditionalFormatType()
	{
		// TODO Auto-generated method stub
		return 0;
	}
}
