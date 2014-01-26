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
package org.openxls.ExtenXLS;

import org.openxls.formats.XLS.Dv;
import org.openxls.formats.XLS.ValidationException;
import org.openxls.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * ValidationHandle allows for manipulation of the validation cells in Excel
 * <p>
 * Using the ValidationHandle, the affected range of validations can be
 * modified, along with many of the error messages and actual validation upon
 * the cells. Many of these calls are very self-explanatory and can be found in
 * the api.
 * </p>
 * Some common use cases:
 * <ol>
 * <li>Setting/changing the validation data type.
 * <pre>
 * // Change validation to only allow a formula ValidationHandle validator =
 * theSheet.getValidationHandle("A1");
 * validator.setValidationType(ValidationHandle.VALUE_FORMULA); // can use any VALUE static byte here.
 * </pre>
 * </li>
 * </li><li>
 * Setting/changing the validation condition. This requires setting the
 * condition type, and a first and second condition. Also make the page error on
 * invalid entry and set the error text.
 * <pre>
 * // Validate cell is an int between the current values of cell D1 and cell D2.
 * ValidationHandle validator = theSheet.getValidationHandle("A1");
 * validator.setValidationType(ValidationHandle.VALUE_INTEGER);
 * validator.setTypeOperator(ValidationHandle.CONDITION_BETWEEN)// any CONDITION
 * static byte validator.setFirstCondition("D1"); // any valid excel formula,
 * omitting the '=' validator.setSecondCondition("D2");
 * validator.setErrorBoxText
 * ("The value is not between the values of D1 and D2");
 * validator.setShowErrorMessage(true);
 * </pre>
 * </li><li>
 * Change the range the validation is applied to, for instance if one is
 * inseting a number of new rows and wants to grow the range A1:D1 to A1:Z1.
 * <pre>
 * ValidationHandle validator = theSheet.getValidationHandle("A1");
 * validator.setRange("A1:Z1");
 * </pre>
 * </li>
 * </ol>
 * "http://www.extentech.com">Extentech Inc.</a>
 */

public class ValidationHandle implements Handle
{
	private static final Logger log = LoggerFactory.getLogger( ValidationHandle.class );
	private Dv myDv;

	// static shorts for setting validation type
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

	// static shorts for setting conditions on validation
	public static final byte CONDITION_BETWEEN = 0x0;
	public static final byte CONDITION_NOT_BETWEEN = 0x1;
	public static final byte CONDITION_EQUAL = 0x2;
	public static final byte CONDITION_NOT_EQUAL = 0x3;
	public static final byte CONDITION_GREATER_THAN = 0x4;
	public static final byte CONDITION_LESS_THAN = 0x5;
	public static final byte CONDITION_GREATER_OR_EQUAL = 0x6;
	public static final byte CONDITION_LESS_OR_EQUAL = 0x7;

	// static shorts for setting IME modes
	public static short IME_MODE_NO_CONTROL = 0x0; // No control for IME.
	// (default)
	public static short IME_MODE_ON = 0x1; // IME is on.
	public static short IME_MODE_OFF = 0x2; // IME is off.
	public static short IME_MODE_DISABLE = 0x3;// IME is disabled.
	public static short IME_MODE_HIRAGANA = 0x4;// IME is in hiragana input
	// mode.
	public static short IME_MODE_KATAKANA = 0x5;// IME is in full-width katakana
	// input mode.
	public static short IME_MODE_KATALANA_HALF = 0x6;// IME is in half-width
	// katakana input mode.
	public static short IME_MODE_FULL_WIDTH_ALPHA = 0x7;// IME is in full-width
	// alphanumeric input
	// mode
	public static short IME_MODE_HALF_WIDTH_ALPHA = 0x8;// IME is in half-width
	// alphanumeric input
	// mode.
	public static short IME_MODE_FULL_WIDTH_HANKUL = 0x9;// IME is in full-width
	// Hankul input mode
	public static short IME_MODE_HALF_WIDTH_HANKUL = 0x10;// IME is in
	// half-width Hankul
	// input mode.

	public static String[] CONDITIONS = {
			"between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual"
	};

	/**
	 * Get the byte representing the condition type string passed in. Options
	 * are' "between", "notBetween", "equal", "notEqual", "greaterThan",
	 * "lessThan", "greaterOrEqual", "lessOrEqual"
	 *
	 * @return
	 */
	public static byte getConditionNumber( String conditionType )
	{
		for( int i = 0; i < CONDITIONS.length; i++ )
		{
			if( conditionType.equalsIgnoreCase( CONDITIONS[i] ) )
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
	 * Get the byte representing the value type string passed in. Options are'
	 * "any", "integer", "decimal", "userDefinedList", "date", "time",
	 * "textLength", "formula"
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
	 * Create a new Validation for the input cell range.
	 *
	 * Validations are specific to worksheets.
	 *
	 *
	 * @param cellRange
	 *            = cell or range of cells, example "A1", "A1:A10"
	 *
	 *            protected ValidationHandle(String cellRange, WorkSheetHandle
	 *            wsh) {
	 *
	 *            }
	 */
	/**
	 * For internal use only. Creates a Validation Handle based of the Dv passed
	 * in.
	 *
	 * @param dv
	 */
	public ValidationHandle( Dv dv )
	{
		myDv = dv;
	}

	/**
	 * Return the range of data this ValidationHandle refers to as a string Will
	 * not contain worksheet identifier, as ValidationHandles are specific to a
	 * worksheet. If the Validation effects multiple ranges they are separated
	 * by a space
	 *
	 * @return ptgRef.toString()
	 */
	public String getRange()
	{
		String[] s = myDv.getRanges();
		String out = "";
		for( int i = 0; i < s.length; i++ )
		{
			if( i > 0 )
			{
				out += " ";
			}
			out += s[i];
		}
		return out;
	}

	/**
	 * Determine if the value passed in is valid for this validation
	 *
	 * @param value
	 * @return
	 */
	public boolean isValid( Object value ) throws ValidationException
	{
		return myDv.isValid( value );
	}

	/**
	 * Determine if the value passed in is valid for this validation
	 *
	 * @param value
	 * @return
	 */
	public boolean isValid( Object value, boolean throwException ) throws RuntimeException
	{
		if( throwException )
		{
			try
			{
				return myDv.isValid( value );
			}
			catch( Exception e )
			{
				try
				{
					throw e;
				}
				catch( Exception ex )
				{
					log.error( "Error getting isValid " + ex.toString() );
				}
			}
			;
		}
		try
		{
			return myDv.isValid( value );
		}
		catch( Exception e )
		{
			return false;
		}
	}

	/**
	 * Set the range this ValidationHandle refers to. Pass in a range string,
	 * sans worksheet.
	 * <p/>
	 * This range will overwrite all other ranges this ValidationHandle refers
	 * to.
	 * <p/>
	 * In order to handle multiple ranges, use the addRange(String range) method
	 *
	 * @param range = standard excel range without worksheet information ("A1" or
	 *              "A1:A10")
	 */
	public void setRange( String range )
	{
		myDv.setRange( range );
	}

	/**
	 * Adds an additional range to the existing ranges in this validationhandle
	 *
	 * @param range
	 */
	public void addRange( String range )
	{
		myDv.addRange( range );
	}

	/**
	 * Get the text from the error box.
	 *
	 * @return
	 */
	public String getErrorBoxText()
	{
		return myDv.getErrorBoxText();
	}

	/**
	 * Set the text for the error box
	 *
	 * @param textError
	 */
	public void setErrorBoxText( String textError )
	{
		myDv.setErrorBoxText( textError );
	}

	/**
	 * Return the text in the prompt box
	 *
	 * @return
	 */
	public String getPromptBoxText()
	{
		return myDv.getPromptBoxText();
	}

	/**
	 * Set the text for the prompt box
	 *
	 * @param text
	 */
	public void setPromptBoxText( String text )
	{
		myDv.setPromptBoxText( text );
	}

	/**
	 * Set the title for the error box
	 *
	 * @param textError
	 */
	public void setErrorBoxTitle( String textError )
	{
		myDv.setErrorBoxTitle( textError );
	}

	/**
	 * Get the title from the error box
	 *
	 * @return
	 */
	public String getErrorBoxTitle()
	{
		return myDv.getErrorBoxTitle();
	}

	/**
	 * Return the title in the prompt box
	 *
	 * @return
	 */
	public String getPromptBoxTitle()
	{
		return myDv.getPromptBoxTitle();
	}

	/**
	 * Set the title for the prompt box
	 *
	 * @param text
	 */
	public void setPromptBoxTitle( String text )
	{
		myDv.setPromptBoxTitle( text );
	}

	/**
	 * Return a byte representing the error style for this ValidationHandle
	 * <p/>
	 * These map to the static final ints ERROR_* in ValidationHandle
	 *
	 * @return
	 */
	public byte getErrorStyle()
	{
		return myDv.getErrorStyle();
	}

	/**
	 * Set the error style for this ValidationHandle record
	 * <p/>
	 * These map to the static final ints ERROR_* from ValidationHandle
	 *
	 * @return
	 */
	public void setErrorStyle( byte errstyle )
	{
		myDv.setErrorStyle( errstyle );
	}

	/**
	 * Get the IME mode for this validation
	 *
	 * @return
	 */
	public short getIMEMode()
	{
		return myDv.getIMEMode();
	}

	/**
	 * set the IME mode for this validation
	 *
	 * @return
	 */
	public void setIMEMode( short mode )
	{
		myDv.setIMEMode( mode );
	}

	/**
	 * Allow blank cells in the validation area?
	 *
	 * @return
	 */
	public boolean isAllowBlank()
	{
		return myDv.isAllowBlank();
	}

	/**
	 * Allow blank cells in the validation area?
	 *
	 * @return
	 */
	public void setAllowBlank( boolean allowBlank )
	{
		myDv.setAllowBlank( allowBlank );
	}

	/**
	 * Get the first condition of the validation as a string representation
	 *
	 * @return
	 */
	public String getFirstCondition()
	{
		return myDv.getFirstCond();
	}

	/**
	 * Set the first condition of the validation
	 * <p/>
	 * This value must conform to the Value Type of this validation or
	 * unexpected results may occur. For example, entering a string
	 * representation of a date here will not work if your validation is an
	 * integer...
	 * <p/>
	 * A java.util.Date object can also be passed in. This value will be
	 * translated into an integer as excel stores dates. If you need to
	 * manipulate/retrieve this value later utilize the DateConverter tool to
	 * transform the value
	 * <p/>
	 * String passed in should be a vaild XLS formula. Does not need to include
	 * the "="
	 * <p/>
	 * Types of conditions Integer values Decimal values User defined list Date
	 * Time Text length Formula
	 * <p/>
	 * Be sure that your validation type (getValidationType()) matches the type
	 * of data.
	 *
	 * @param firstCond = the first condition for the validation
	 */
	public void setFirstCondition( Object firstCond )
	{
		String setval = firstCond.toString();
		if( firstCond instanceof java.util.Date )
		{
			double d = DateConverter.getXLSDateVal( (java.util.Date) firstCond );
			setval = d + "";
		}
		myDv.setFirstCond( setval );
	}

	/**
	 * Get the second condition of the validation as a string representation
	 *
	 * @return
	 */
	public String getSecondCondition()
	{
		return myDv.getSecondCond();
	}

	/**
	 * Set the first condition of the validation utilizing a string. This value
	 * must conform to the Value Type of this validation or unexpected results
	 * may occur. For example, entering a string representation of a date here
	 * will not work if your validation is an integer...
	 * <p/>
	 * String passed in should be a vaild XLS formula. Does not need to include
	 * the "="
	 * <p/>
	 * Types of conditions Integer values Decimal values User defined list Date
	 * Time Text length Formula
	 * <p/>
	 * Be sure that your validation type (getValidationType()) matches the type
	 * of data.
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
		myDv.setSecondCond( setval );
	}

	/**
	 * Show error box if invalid values entered?
	 *
	 * @return
	 */
	public boolean isShowErrorMsg()
	{
		return myDv.isShowErrorMsg();
	}

	/**
	 * Set show error box if invalid values entered?
	 *
	 * @return
	 */
	public void setShowErrorMsg( boolean showErrMsg )
	{
		myDv.setShowErrMsg( showErrMsg );
	}

	/**
	 * Show prompt box if cell selected?
	 *
	 * @return
	 */
	public boolean getShowInputMsg()
	{
		return myDv.getShowInputMsg();
	}

	/**
	 * Set show prompt box if cell selected?
	 *
	 * @param showInputMsg
	 */
	public void setShowInputMsg( boolean showInputMsg )
	{
		myDv.setShowInputMsg( showInputMsg );
	}

	/**
	 * In list type validity the string list is explicitly given in the formula
	 *
	 * @return boolean
	 */
	public boolean isStrLookup()
	{
		return myDv.isStrLookup();
	}

	/**
	 * In list type validity the string list is explicitly given in the formula
	 *
	 * @param strLookup
	 */
	public void setStrLookup( boolean strLookup )
	{
		myDv.setStrLookup( strLookup );
	}

	/**
	 * Suppress the drop down arrow in list type validity
	 *
	 * @return boolean
	 */
	public boolean isSuppressCombo()
	{
		return myDv.isSuppressCombo();
	}

	/**
	 * Suppress the drop down arrow in list type validity
	 */
	public void setSuppressCombo( boolean suppressCombo )
	{
		myDv.setSuppressCombo( suppressCombo );
	}

	/**
	 * Get the type operator of this validation as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in ValidationHandle
	 *
	 * @return
	 */
	public byte getTypeOperator()
	{
		return myDv.getTypeOperator();
	}

	/**
	 * set the type operator of this validation as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in ValidationHandle
	 */
	public void setTypeOperator( byte typOperator )
	{
		myDv.setTypeOperator( typOperator );
	}

	/**
	 * Get the validation type of this ValidationHandle as a byte
	 * <p/>
	 * These bytes map to the VALUE_* static values in ValidationHandle
	 *
	 * @return
	 */
	public byte getValidationType()
	{
		return myDv.getValType();
	}

	/**
	 * Set the validation type of this ValidationHandle as a byte
	 * <p/>
	 * These bytes map to the VALUE_* static values in ValidationHandle
	 */
	public void setValidationType( byte valtype )
	{
		myDv.setValType( valtype );
	}

	/**
	 * Determines if the ValidationHandle contains the cell address passed in
	 *
	 * @param range
	 * @return
	 */
	public boolean isInRange( String celladdy )
	{
		return myDv.isInRange( celladdy );
	}

	/**
	 * Return an xml representation of the ValidationHandle
	 *
	 * @return
	 */
	public String getXML()
	{
		StringBuffer xml = new StringBuffer();
		xml.append( "<datavalidation" );
		xml.append( " type=\"" + VALUE_TYPE[getValidationType()] + "\"" );
		xml.append( " operator=\"" + CONDITIONS[getTypeOperator()] + "\"" );
		xml.append( " allowBlank=\"" + ((isAllowBlank()) ? "1" : "0") + "\"" );
		xml.append( " showInputMessage=\"" + ((getShowInputMsg()) ? "1" : "0") + "\"" );
		xml.append( " showErrorMessage=\"" + ((isShowErrorMsg()) ? "1" : "0") + "\"" );
		xml.append( " errorTitle=\"" + getErrorBoxTitle() + "\"" );
		xml.append( " error=\"" + getErrorBoxText() + "\"" );
		xml.append( " promptTitle=\"" + getPromptBoxTitle() + "\"" );
		xml.append( " prompt=\"" + getPromptBoxText() + "\"" );
		try
		{
			xml.append( " sqref=\"" + getRange() + "\"" );
		}
		catch( Exception e )
		{
			log.error( "Problem getting range for ValidationHandle.getXML().", e );
		}
		xml.append( ">" );
		if( getFirstCondition() != null )
		{
			xml.append( "<formula1>" );
			xml.append( StringTool.convertXMLChars( getFirstCondition() ) );
			xml.append( "</formula1>" );
		}
		if( getSecondCondition() != null )
		{
			xml.append( "<formula2>" );
			xml.append( StringTool.convertXMLChars( getSecondCondition() ) );
			xml.append( "</formula2>" );
		}
		xml.append( "</datavalidation>" );
		return xml.toString();
	}

}
