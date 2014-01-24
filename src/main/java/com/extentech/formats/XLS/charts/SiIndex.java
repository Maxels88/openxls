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

import com.extentech.toolkit.ByteTools;

/**
 * <b>SiIndex: Series Index (0x1065)</b>
 * <p/>
 * Indicates the type of data contained in the Number records following it.
 * <p/>
 * 2 bytes
 * 0x0001
 * Series values or vertical values (for scatter or bubble chart groups)
 * 0x0002
 * Category labels or horizontal values (for scatter or bubble chart groups)
 * 0x0003
 * Bubble sizes
 * <p/>
 * specifies the beginning of a sequence of records that contains a cache of the data for the sequence of records that conforms to a specific AI rule (section 2.1.7.20.1) in the series (section 2.2.3.9) and error bars (section 2.2.3.13).
 * The relationship between the series and the chart data cache is specified as follows:
 * The first SIIndex record in the chart sheet substream, which MUST contain a numIndex field
 * equal to 0x0001, corresponds to the second sequence of records that conforms to the AI rule
 * The second SIIndex record in the chart sheet substream, which MUST contain a numIndex field equal to 0x0002,
 * corresponds to the third sequence of records that conforms to the AI rule (section 2.1.7.20.1).
 * The third SIIndex record in the chart sheet substream, which MUST contain a numIndex field equal to 0x0003,
 * corresponds to the fourth sequence of records that conforms to the AI rule (section 2.1.7.20.1).
 * The Number, BoolErr , Blank , and Label  records each specify an individual value stored in the cache. Each column in the cache
 * corresponds to a series or error bar, where the zero-based index of the column, specified by the cell.col field in the Number, BoolErr, Blank, or Label records,
 * equals the zero-based index of the Series record in the collection of Series records that corresponds to the series or error bar.
 */
public class SiIndex extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6810089722566956477L;
	short type;

	@Override
	public void init()
	{
		super.init();
		type = ByteTools.readShort( getData()[0], getData()[1] );
	}
}
