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

import java.util.List;

/**
 * Muls are used to represent multiple compressed values in a row.
 *
 * @see Mulrk
 * @see Mulblank
 * @see Rk
 * @see Blank
 */
public interface Mul
{

	/**
	 * whether this mul was removed from the SheetRecs already
	 *
	 * @return
	 */
	public boolean removed();

	/**
	 * return the list of active mulleds
	 * <p/>
	 * Feb 8, 2011
	 *
	 * @return
	 */
	public List getRecs();

}
