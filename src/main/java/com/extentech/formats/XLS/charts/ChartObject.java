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
 * ChartObject is an interface for chart records that can be described as an object.  Primarily
 * this is just a heierarchial system.  Chart objects contain an array of records that could be value records
 * and/or other chartObjects containing their own arrays.
 *
 *
 */

import com.extentech.formats.XLS.XLSRecord;

import java.util.ArrayList;

public interface ChartObject
{

	/**
	 * Add a chart record to this chart object
	 */
	public void addChartRecord( XLSRecord b );

	/**
	 * Return a list of chart records from this chart object.  This should include begin/end records, along with
	 * those of the objects sub-object, correctly ordered.
	 *
	 * @return
	 */
	public ArrayList getChartRecords();

	/**
	 * Return the chart record that this object is associated with
	 *
	 * @return
	 */
	public Chart getParentChart();

	/**
	 * Set the parent chart for this record
	 */
	public void setParentChart( Chart c );

	/**
	 * Get the output array of records, including begin/end records and those of it's children.
	 */
	public ArrayList getRecordArray();

}
