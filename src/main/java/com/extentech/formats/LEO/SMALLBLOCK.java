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
package com.extentech.formats.LEO;

/**
 * LEO File SMALLBLOCK Information Record
 * <p/>
 * These blocks of data contain information related to the
 * LEO file format data blocks.
 * <p/>
 * depending on the size and complexity of the file,
 * these records may all be contained in the 'header'
 * block of the LEO Stream File.
 * <p/>
 * In files over that size, one or more SMALLBLOCK records
 * are inserted directly in the midst of substream records.
 */
public class SMALLBLOCK extends BlockImpl
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7432771150988897281L;
	public final static int SIZE = 64;

	/**
	 * returns the int representing the block type
	 */
	public final int getBlockType()
	{
		return SMALL;
	}

}