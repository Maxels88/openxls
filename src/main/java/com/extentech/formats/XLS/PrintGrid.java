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
 * Record specifying whether the grid should be printed.
 */
public class PrintGrid extends XLSRecord
{
	private static final long serialVersionUID = -3649192673573344145L;

	@Override
	public void init()
	{
		super.init();
	}

	@Override
	public void setSheet( Sheet sheet )
	{
		super.setSheet( sheet );
		((Boundsheet) sheet).addPrintRec( this );
	}

	public boolean isPrintGrid()
	{
		return (getData()[0] & 0x01) == 0x01;
	}

	public void setPrintGrid( boolean print )
	{
		if( print )
		{
			getData()[0] |= 0x01;
		}
		else
		{
			getData()[0] &= ~0x01;
		}
	}
}
