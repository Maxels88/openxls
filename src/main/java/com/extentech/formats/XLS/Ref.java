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

import com.extentech.formats.XLS.formulas.Ptg;

/**
 * Ref defines an ExtenXLS record or object that represents a range of cell locations.
 *
 * @see Ptg
 * @see Formula
 */
public interface Ref
{

	public static int PTG_LOCATION_POLICY_UNLOCKED = 0;
	public static int PTG_LOCATION_POLICY_LOCKED = 1;
	public static int PTG_LOCATION_POLICY_TRACK = 2;

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
	 * @return String Location
	 */
	String getLocation() throws FormulaNotFoundException;

	int[] getIntLocation() throws FormulaNotFoundException;

	/**
	 * returns the row/col ints for the ref
	 *
	 * @return the row col int array
	 */
	public int[] getRowCol();

	/**
	 * returns the String address of this ptg including sheet reference
	 *
	 * @return the String location of the reference including sheetname
	 */
	public String getLocationWithSheet();

	/**
	 * gets the sheetname for this ref
	 *
	 * @param sheetname
	 */
	public String getSheetName() throws WorkSheetNotFoundException;

}