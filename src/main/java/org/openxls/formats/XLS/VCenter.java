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
 * Record specifying whether the sheet is to be centered vertically
 * when printed.
 */
public class VCenter extends XLSRecord
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

	public boolean isVCenter()
	{
		return (getData()[0] & 0x01) == 0x01;
	}

	public void setVCenter( boolean center )
	{
		if( center )
		{
			getData()[0] |= 0x01;
		}
		else
		{
			getData()[0] &= ~0x01;
		}
	}
}
