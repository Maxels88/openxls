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
 * A BIFF8 record that contains the value of a single cell.
 * For the moment this contains very little. Eventually most of the
 * cell-specific methods from {@link XLSRecord} should be moved here.
 */
public abstract class XLSCellRecord extends XLSRecord implements CellRec
{
	private static final long serialVersionUID = 7387720078386279196L;

	@Override
	public int getColFirst()
	{
		return this.getColNumber();
	}

	@Override
	public int getColLast()
	{
		return this.getColNumber();
	}

	@Override
	public boolean isSingleCol()
	{
		return (this.getColFirst() == this.getColLast());
	}
}
