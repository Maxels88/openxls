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

/**
 * Classes implementing this interface need to be able to
 * stream the bytes of their contained records to the calling
 * class.
 *
 * @see WorkBook
 * @see WorkSheet
 */

public interface XLSSubstream
{
	String getName();

	String getSubstreamTypeName();

	short getSubstreamType();

	byte[] stream();

	short WK_GLOBALS = 0x5;
	short VB_MODULE = 0x6;
	short WK_WORKSHEET = 0x10;
	short WK_CHART = 0x20;
	short WK_MACROSHEET = 0x40;
	short WK_FILE = 0x100;

}