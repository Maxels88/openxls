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
package org.openxls.formats.XLS.charts;

import org.openxls.toolkit.ByteTools;

/**
 * <b>SeriesList: Specifies the Series in an Overlay Chart (0x1016)</b>
 * <p/>
 * bytes - 2	- nseries following
 * 2 * nseries = An array of 2-byte unsigned integers,
 * each of which specifies a one-based index of a Series record
 * in the collection of Series records in the current chart sheet substream
 */
public class SeriesList extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7852050067799624402L;
	int[] seriesmap = null;

	@Override
	public void init()
	{
		super.init();
		int nseries = ByteTools.readShort( getData()[0], getData()[1] );
		seriesmap = new int[nseries];
		for( int i = 0; i < nseries; i++ )
		{
			int idx = ((i + 1) * 2);
			seriesmap[i] = ByteTools.readShort( getData()[idx], getData()[idx + 1] );
		}
	}

	/**
	 * return the series mappings for the associated overlay chart
	 * <br>series mappings links the overlay chart to the absolute series number
	 * (determined by the actual order of the series in the chart array structure)
	 *
	 * @return
	 */
	public int[] getSeriesMappings()
	{
		return seriesmap;
	}

	/**
	 * set the series mappings for the associated overlay chart
	 * <br>series mappings links the overlay chart to the absolute series number
	 * (determined by the actual order of the series in the chart array structure)
	 *
	 * @param seriesmap
	 */
	public void setSeriesMappings( int[] smap )
	{
		short nseries = (short) smap.length;
		seriesmap = new int[nseries];
		byte[] data = new byte[(nseries + 1) * 2];
		byte[] b = ByteTools.shortToLEBytes( nseries );
		data[0] = b[0];
		data[1] = b[1];
		for( int i = 0; i < nseries; i++ )
		{
			int idx = ((i + 1) * 2);
			seriesmap[i] = smap[i];
			b = ByteTools.shortToLEBytes( (short) smap[i] );
			data[idx] = b[0];
			data[idx + 1] = b[1];
		}
		setData( data );
	}
}
