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

import org.openxls.formats.XLS.BiffRec;
import org.openxls.formats.XLS.FormulaNotFoundException;
import org.openxls.formats.XLS.FunctionNotSupportedException;
import org.openxls.formats.XLS.XLSConstants;
import org.openxls.formats.XLS.XLSRecord;

import java.io.Serializable;

/**
 * Ptg is the interface all ptgs implement in order to be handled equally under the
 * eyes of the all seeing one, "he that shall not be named"  A ptg is a unique segment
 * of a formula stack that indicates a value, a reference to a value, or an operation.
 * See the docs under Formula for more information.
 *
 * @see Ptg
 * @see Formula
 */

public interface Ptg extends XLSConstants, Serializable
{

	/**
	 * VALUE type Reference (Id=0x44)
	 */
	public static short VALUE = 0;
	/**
	 * REFERENCE type Reference (Id=0x24)
	 */
	public static short REFERENCE = 1;
	/**
	 * ARRAY type Reference (Id=0x64)
	 */
	public static short ARRAY = 2;

	public static int PTG_LOCATION_POLICY_UNLOCKED = 0;
	public static int PTG_LOCATION_POLICY_LOCKED = 1;
	public static int PTG_LOCATION_POLICY_TRACK = 2;

	public static int PTG_TYPE_SINGLE = 1; // single-byte record
	public static int PTG_TYPE_ARRAY = 2; // array of bytes record

	//ptg lengths
	public static int PTG_NUM_LENGTH = 9;
	public static int PTG_ADD_LENGTH = 1;
	public static int PTG_AREA_LENGTH = 9;
	public static int PTG_AREA3D_LENGTH = 11;
	public static int PTG_AREAERR3D_LENGTH = 11;
	public static int PTG_ATR_LENGTH = 4;
	public static int PTG_CONCAT_LENGTH = 1;
	public static int PTG_DIV_LENGTH = 1;
	public static int PTG_EQ_LENGTH = 1;
	public static int PTG_EXP_LENGTH = 5;
	public static int PTG_FUNC_LENGTH = 3;
	public static int PTG_FUNCVAR_LENGTH = 4;
	public static int PTG_GE_LENGTH = 1;
	public static int PTG_GT_LENGTH = 1;
	public static int PTG_INT_LENGTH = 3;
	public static int PTG_ISECT_LENGTH = 1;
	public static int PTG_LE_LENGTH = 1;
	public static int PTG_LT_LENGTH = 1;
	public static int PTG_MEMERR_LENGTH = 7;
	public static int PTG_MEM_AREA_N_LENGTH = 7;
	public static int PTG_MEM_AREA_NV_LENGTH = 7;
	public static int PTG_MLT_LENGTH = 1;
	public static int PTG_MYSTERY_LENGTH = 1;
	public static int PTG_NE_LENGTH = 1;
	public static int PTG_NAME_LENGTH = 5;
	public static int PTG_NAMEX_LENGTH = 7;
	public static int PTG_PAREN_LENGTH = 1;
	public static int PTG_POWER_LENGTH = 1;
	public static int PTG_RANGE_LENGTH = 1;
	public static int PTG_REF_LENGTH = 5;
	public static int PTG_REF3D_LENGTH = 7;
	public static int PTG_REFERR_LENGTH = 5;
	public static int PTG_REFERR3D_LENGTH = 7;
	public static int PTG_ENDSHEET_LENGTH = 1;
	public static int PTG_SUB_LENGTH = 1;
	public static int PTG_UNION_LENGTH = 1;
	public static int PTG_BOOL_LENGTH = 2;
	public static int PTG_UPLUS_LENGTH = 1;
	public static int PTG_UMINUS_LENGTH = 1;
	public static int PTG_PERCENT_LENGTH = 1;

	//TODO:  add all the opcodes here
	public static byte PTG_INT = 0x1e;

	/**
	 * Creates a deep clone of this Ptg.
	 */
	public Object clone();

	/**
	 * constructor must pass in 'parent' XLSRecord so that there
	 * is a handle for updating...
	 *
	 * @return
	 */
	void setParentRec( XLSRecord x );

	/**
	 * update the values of the Ptg
	 */
	void updateRecord();

	/**
	 * return the length of the Ptg
	 */
	int getLength();

	/**
	 * return the number of parameters to this Ptg
	 */
	int getNumParams();

	public static final int CALCULATED = 0;
	public static final int UNCALCULATED = -1;

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 * tells the Ptg to get it on its own...
	 */
	public void updateAddressFromTrackerCell();

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 * tells the Ptg to get it on its own...
	 */
	public void initTrackerCell();

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 *
	 * @return trackercell The trackercell to set.
	 */
	public BiffRec getTrackercell();

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 *
	 * @param trackercell The trackercell to set.
	 */
	public void setTrackercell( BiffRec trackercell );

	/**
	 * a locking mechanism so that Ptgs are not endlessly
	 * re-calculated
	 *
	 * @return
	 */
	int getLock();

	/**
	 * a locking mechanism so that Ptgs are not endlessly
	 * re-calculated
	 *
	 * @return
	 */
	void setLock( int x );

	/**
	 * determine the general Ptg type
	 */
	boolean getIsOperator();

	boolean getIsBinaryOperator();

	boolean getIsUnaryOperator();

	boolean getIsStandAloneOperator();

	boolean getIsOperand();

	boolean getIsControl();

	boolean getIsFunction();

	boolean getIsReference();

	/**
	 * Operator Ptgs take other Ptgs as arguments
	 * so we need to pass them in to get a meaningful
	 * value.
	 */
	void setVars( Ptg[] parr );

	/**
	 * determines whether this operator is a 'primitive' such as +,-,=,<,>,!=,==,etc.
	 * the upshot is that primitives go BETWEEN operands, and non-primitives
	 * encapsulate
	 * <p/>
	 * ie:
	 * <p/>
	 * SUM(A1:A4)  non-primitive
	 * A1+A4       primitive
	 */
	boolean getIsPrimitiveOperator();

	/*
	   Determines whether the ptg represents multiple ptg's in reality.
	   ie ptgArea ia actually a collection of ptgRef's, so ptgArea.getIsArray returns 'true'
	*/
	boolean getIsArray();

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 */
	String getTextString();

	/**
	 * pass  in arbitrary number of values (probably other Ptgs)
	 * and return the resultant value.
	 * <p/>
	 * This effectively calculates the Expression.
	 */
	Object evaluate( Object[] obj );

	/**
	 * If a record consists of multiple sub records (ie PtgArea) return those
	 * records, else return null;
	 */
	Ptg[] getComponents();

	/**
	 * @return byte[] containing the whole ptg, including identifying opcode
	 */
	byte[] getRecord();

	/*
		@return XLSRecord containing the whole ptg       
	*/
	XLSRecord getParentRec();

	/**
	 * returns whether the Location of the Ptg is locked
	 * used during automated BiffRec movement updates
	 *
	 * @return location policy
	 */
	int getLocationPolicy();

	/**
	 * lock the Location of the Ptg so that it will not
	 * be updated during automated BiffRec movement updates
	 *
	 * @param b whether to lock the location of this Ptg
	 */
	void setLocationPolicy( int b );

	/**
	 * setLocation moves a ptg that is a reference to a location, such as
	 * a ptg range being modified
	 *
	 * @param String location, such as A1:D4
	 */
	void setLocation( String s );

	/**
	 * When the ptg is a reference to a location this returns that location
	 *
	 * @return Location
	 */
	String getLocation() throws FormulaNotFoundException;

	int[] getIntLocation() throws FormulaNotFoundException;

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 */
	public String getString();

	/**
	 * return the byte header for the Ptg
	 */
	public byte getOpcode();

	/**
	 * return the human-readable String representation of
	 * the "closing" portion of this Ptg
	 * such as a closing parenthesis.
	 */
	public String getString2();

	/**
	 * return a Ptg  consisting of the calculated values
	 * of the ptg's passed in.  Returns null for any non-operator
	 * ptg.
	 *
	 * @throws CalculationException
	 */
	public Ptg calculatePtg( Ptg[] parsething ) throws FunctionNotSupportedException, CalculationException;

	/**
	 * Gets the (return) value of this Ptg as an operand Ptg.
	 */
	public Ptg getPtgVal();

	/**
	 * returns the value of an operand ptg.
	 *
	 * @return null for non-operand Ptg.
	 */
	public Object getValue();

	/**
	 * Gets the value of the ptg represented as an int.
	 * <p/>
	 * This can result in loss of precision for floating point values.
	 * <p/>
	 * -1 will be returned for values that are not translatable to an integer
	 *
	 * @return integer representing the ptg, or NAN
	 */
	public int getIntVal();

	public double getDoubleVal();

	public boolean isBlank();

	public void close();

}