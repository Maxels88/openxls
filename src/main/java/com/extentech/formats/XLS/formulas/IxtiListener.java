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

/**
 * An IxtiListener is aware of changes to the Externsheet Ixti references.
 * <p/>
 * Notably, the PtgRef3D and PtgArea recs...
 */
public interface IxtiListener
{

	/**
	 * @return Returns the ixti.
	 */
	public short getIxti();

	/**
	 * @param ixti The ixti to set.
	 */
	public void setIxti( short ixti );

	/**
	 * Add this to the ixti listeners
	 */
	public void addListener();

}