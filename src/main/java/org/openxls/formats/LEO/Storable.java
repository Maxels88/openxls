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
package org.openxls.formats.LEO;

import java.util.Vector;

/**
 * Create a Storage interface/Generic implementation
 * to represent a file or index of files within a LEO filesystem.
 */
public interface Storable
{

	/** causes the raw file data to be read in as a byte array
	 */

	/**
	 * associate this storage with its Block
	 * table entries and data
	 */
	public void init( Vector dta, int[] tab );

	/**
	 * sets whether this Storage's Data blocks are contained
	 * in the Small or Big Block arrays.
	 */
	public void setBlockType( int type );

}