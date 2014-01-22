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
package com.extentech.formats.XLS.charts;

/**
 * <b>ShtProps: Sheet Properties (0x1044)</b>
 * <p/>
 * 4		grbit		2
 * 6		mdBlank		1		Empty cells plotted as: 0= not plotted, 1= 0, 2= interpolated
 * <p/>
 * grbit:
 * 0	0x1		fManSerAlloc		1= chart has been changed from default
 * 1	0x2		fPlotVisOnly		1= plot visible cells only
 * 2	0x4		fNotSizeWith		1= do not size chart with window
 * 3	0x8		fManPlotArea		0= default dimensions	1= use POS rec
 * 4	0x10	fAlwaysAutoPlotArea	1= user has modified chart enough that fManPlotArea should be set to 0 (!!!)
 */
public class ShtProps extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5462571460161191942L;

	@Override
	public void init()
	{
		super.init();
	}
}
