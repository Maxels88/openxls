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
package org.openxls.formats.XLS;

/**
 * <b> Phonetic
 * <p/>
 * This record contains default settings for the Asian Phonetic Settings dialog and the addresses of all cells which show
 * Asian phonetic text.
 * Record PHONETIC, BIFF8:
 * Offset Size Contents
 * 0 	2 	Index to FONT record used for Asian phonetic text of new cells
 * 2 	2 	Additional settings used for Asian phonetic text of new cells:
 * Bit 	Mask 	Contents
 * 1-0 	0003H 	Type of Japanese phonetic text:
 * 002 = Katakana (narrow) 102 = Hiragana
 * 012 = Katakana (wide)
 * 3-2 	000CH 	Alignment of all portions of the Asian phonetic text:
 * 002 = Not specified (Japanese only) 102 = Centered
 * 012 = Left (Top for vertical text) 112 = Distributed
 * 5-4 	0030H 	112 (always set)
 * 4 		var. 	Cell range address list with all cells with visible Asian phonetic text
 * <p/>
 * <b>Unknown Record possibly MSO related: MS Office (EFh)</b><br>
 * <p/>
 * These records contain only data.<p><pre>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       rgMSODrGr    6    	who the hell knows?
 * <p/>
 * </p></pre>
 *
 * @see MSODrawing
 */
public final class Phonetic extends XLSRecord
{
	/**
	 *
	 */
	private static final long serialVersionUID = -3464806666220926523L;
	public byte[] PROTOTYPE_BYTES = { 0, 0, 55, 0, 0, 0 };

	/**

	 */

	@Override
	public void init()
	{
		super.init();
	}

}