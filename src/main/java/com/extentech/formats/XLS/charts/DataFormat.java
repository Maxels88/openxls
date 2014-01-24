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

import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

/**
 * <b>DataFormat: Series and Data Point Numbers(0x1006)</b>
 * <p/>
 * 4	xi		2	the zero-based index of the data point within the series specified by yi. (FFFFh means entire series)
 * 6	yi		2	the zero-based index of a Series record
 * 8	iss		2	An unsigned integer that specifies properties of the data series, trendline or error bar, depending on the type of records in sequence of records:
 * <p/>
 * If does not contain a SerAuxTrend or SerAuxErrBar record, then this field specifies the plot order of the data series.
 * If the series order was changed, this field can be different from yi. MUST be less than or equal to the number of series in the chart.
 * MUST be unique among iss values for all instances of this record contained in the SERIESFORMAT rule that does not contain a SerAuxTrend or SerAuxErrBar record.
 * <p/>
 * If the SERIESFORMAT rule contains a SerAuxTrend record on the chart group, then this field specifies the trendline number for the series.
 * <p/>
 * If the SERIESFORMAT rule contains a SerAuxErrBar record on the chart group, then this field specifies a zero-based index into a Series record in the
 * collection of Series records in the current chart sheet substream for which the error bar applies to.
 * <p/>
 * 10	grbit	2	flags (0?)		ignored
 * <p/>
 * <p/>
 * ORDER OF SUB-RECS:
 * [Chart3DBarShape]
 * [LineFormat, AreaFormat, PieFormat]		== lines, fill
 * [SerFormat]			== smoothed lines ...
 * [GelFrame]
 * [MarkerFormat]
 * [AttachedLabel]		== data labels
 * [ShapeProps]
 * [CtrlMltFrt]
 */
public class DataFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -3526272512004348462L;
	private short yi;
	private short xi;
	private short iss;

	@Override
	public void init()
	{
		super.init();
		byte[] rkdata = getData();
		xi = ByteTools.readShort( rkdata[0], rkdata[1] );
		yi = ByteTools.readShort( rkdata[2], rkdata[3] );
		iss = (short) ByteTools.readUnsignedShort( rkdata[4], rkdata[5] );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ -1, -1, 0, 0, 0, 0, 0, 0 };

	public void initNew()
	{
		setOpcode( DATAFORMAT );
		setData( PROTOTYPE_BYTES );
		init();
		Chart3DBarShape cs = new Chart3DBarShape();
		cs.setOpcode( CHART3DBARSHAPE );    // creates default bar shape==0, 0
		addChartRecord( cs );
	}

	/**
	 * Create a new dataformat
	 *
	 * @param parentChart
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		DataFormat df = new DataFormat();
		df.setOpcode( DATAFORMAT );
		df.setData( df.PROTOTYPE_BYTES );
		df.init();
		Chart3DBarShape cs = new Chart3DBarShape();
		cs.setOpcode( CHART3DBARSHAPE );    // creates default bar shape==0, 0
		df.addChartRecord( cs );
		return df;
	}

	public static XLSRecord getPrototypeWithFormatRecs( Chart parentChart )
	{
		return getPrototypeWithFormatRecs( 0, parentChart );
	}

	/**
	 * Create DataFormat Record that HOPEFULLY reflects the necessary associated recs
	 *
	 * @param type
	 * @return DataFormat Record
	 */
	public static XLSRecord getPrototypeWithFormatRecs( int seriesNumber, Chart parentChart )
	{
		DataFormat df = new DataFormat();
		df.setOpcode( DATAFORMAT );
		df.setData( df.PROTOTYPE_BYTES );
		df.init();
		Chart3DBarShape cs = new Chart3DBarShape();
		cs.setOpcode( CHART3DBARSHAPE );    // creates default bar shape==0, 0
		df.addChartRecord( cs );
		df.setSeriesNumber( seriesNumber );
		LineFormat lf = (LineFormat) LineFormat.getPrototype();
		df.addChartRecord( lf );
		AreaFormat af = (AreaFormat) AreaFormat.getPrototype();
		af.setParentChart( parentChart );
		df.addChartRecord( af );
		PieFormat pf = (PieFormat) PieFormat.getPrototype();
		pf.setParentChart( parentChart );
		df.addChartRecord( pf );
		MarkerFormat mf = (MarkerFormat) MarkerFormat.getPrototype();
		mf.setParentChart( parentChart );
		df.addChartRecord( mf );
		 /* pieformat:
	MUST not exist on chart group types other than  ---> doesn't appear true (???) 
	pie, 
	doughnut, 
	bar of pie, or 
	pie of pie. 
	MUST not exist if the chart group type is doughnut and the series is not the outermost series. 
	MUST not exist on the data points on the secondary bar/pie of a bar of pie chart group.	
          */
		return df;
	}

	public void setPointNumber( int idx )
	{
		xi = (short) idx;
		byte[] rkdata = getData();
		byte[] num = ByteTools.shortToLEBytes( (short) idx );
		rkdata[0] = num[0];
		rkdata[1] = num[1];
		setData( rkdata );
	}

	/**
	 * Set the series index
	 */
	public void setSeriesIndex( int idx )
	{
		yi = (short) idx;
		byte[] rkdata = getData();
		byte[] num = ByteTools.shortToLEBytes( (short) idx );
		rkdata[2] = num[0];
		rkdata[3] = num[1];
		setData( rkdata );
	}

	public void setSeriesNumber( int idx )
	{
		iss = (short) idx;
		byte[] rkdata = getData();
		byte[] num = ByteTools.shortToLEBytes( (short) idx );
		rkdata[4] = num[0];
		rkdata[5] = num[1];
		setData( rkdata );
	}

	public short getSeriesIndex()
	{
		return yi;
	}

	public short getSeriesNumber()
	{
		return iss;
	}

	public short getPointNumber()
	{
		return xi;
	}

	private AttachedLabel getAttachedLabelRec( boolean bCreate )
	{
		AttachedLabel al = null;
		al = (AttachedLabel) Chart.findRec( chartArr, AttachedLabel.class );
		if( (al == null) && bCreate )
		{ // basic options are handled via AttachedLabel rec
			al = (AttachedLabel) AttachedLabel.getPrototype();
			int z = Chart.findRecPosition( chartArr, MarkerFormat.class );
			if( z > 0 )
			{
				chartArr.add( z + 1, al );
			}
			else
			{
				addChartRecord( al );
			}
		}
		return al;
	}

	private AreaFormat getAreaFormatRec( boolean bCreate )
	{
		AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
		if( af == null )
		{
			af = (AreaFormat) AreaFormat.getPrototype( 0 );
			addChartRecord( LineFormat.getPrototype() );
			addChartRecord( af );
			addChartRecord( PieFormat.getPrototype() );
			addChartRecord( MarkerFormat.getPrototype() );
		}
		return af;
	}

	/**
	 * return XLSRecord data (2 bytes), which controls bar shape
	 */
	public short getShape()
	{
		Chart3DBarShape cs = (Chart3DBarShape) Chart.findRec( chartArr, Chart3DBarShape.class );
		return cs.getShape();
	}

	/**
	 * set the shape bit of the associated XLSRecord
	 *
	 * @param shape
	 */
	public void setShape( int shape )
	{
		Chart3DBarShape cs = (Chart3DBarShape) Chart.findRec( chartArr, Chart3DBarShape.class );
		cs.setShape( (short) shape );
	}

	/**
	 * set smooth lines setting (applicable for line, scatter charts)
	 *
	 * @param smooth
	 */
	public void setSmoothLines( boolean smooth )
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf == null )
		{
			if( smooth )
			{
				setHasLines();
				sf = (Serfmt) Serfmt.getPrototype();
				int i = Chart.findRecPosition( chartArr, PieFormat.class );
				chartArr.add( i + 1, sf );
				sf.setSmoothedLine( true );
			}
		}
		else
		{
			sf.setSmoothedLine( smooth );
		}
	}

	/**
	 * returns true if this parent chart has smoothed lines (Line, Scatter, Radar charts)
	 *
	 * @return
	 */
	public boolean getSmoothedLines()
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf != null )
		{
			return sf.getSmoothLine();
		}
		return false;
	}

	/**
	 * returns true if this parent chart has lines (Line, Scatter, Radar charts)
	 *
	 * @return
	 */
	public boolean getHasLines()
	{
		LineFormat l = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		if( l != null )
		{
			return l.getLineStyle() != LineFormat.NONE;
		}
		return false;
	}

	/**
	 * sets this chart to have lines (line chart, radar, scatter ...)
	 */
	public void setHasLines()
	{
		setHasLines( 0 );
	}

	/**
	 * sets this chart to have lines (line chart, radar, scatter ...) of the specific line style
	 * <br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 */
	public void setHasLines( int lineStyle )
	{
		LineFormat l = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		if( l == null )
		{    // these come as a group - assume none or only has Chart3DBarShape ...
			int z = Chart.findRecPosition( chartArr, Chart3DBarShape.class ) + 1;
			LineFormat lf = (LineFormat) LineFormat.getPrototype( lineStyle, -1 );
			chartArr.add( z++, lf );
			AreaFormat af = (AreaFormat) AreaFormat.getPrototype();
			af.setParentChart( parentChart );
			chartArr.add( z++, af );
			PieFormat pf = (PieFormat) PieFormat.getPrototype();
			pf.setParentChart( parentChart );
			chartArr.add( z++, pf );
			MarkerFormat mf = (MarkerFormat) MarkerFormat.getPrototype();
			mf.setParentChart( parentChart );
			chartArr.add( z++, mf );

		}
		else
		{
			l.setLineStyle( lineStyle );
		}
	}

	/**
	 * sets 3D Bubbles (Bubble Chart only)
	 *
	 * @param b true if has 3d Bubbles
	 */
	public void setHas3DBubbles( boolean b )
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf == null )
		{
			if( b )
			{
				sf = (Serfmt) Serfmt.getPrototype();
				int i = Chart.findRecPosition( chartArr, PieFormat.class );
				sf.setParentChart( getParentChart() );
				chartArr.add( i + 1, sf );
				sf.setHas3dBubbles( true );
			}
		}
		else
		{
			sf.setHas3dBubbles( b );
		}
	}

	/**
	 * returns true if this paernt chart has 3D bubbles (Bubble chart only)
	 *
	 * @return
	 */
	public boolean getHas3DBubbles()
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf != null )
		{
			return sf.get3DBubbles();
		}
		return false;
	}

	/**
	 * return if data markers are displayed with a shadow on bubble,
	 * scatter, radar, stock, and line chart groups.
	 */
	public boolean getHasShadow()
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf != null )
		{
			return sf.hasShadow();
		}
		return false;
	}

	/**
	 * data markers are displayed with a shadow on bubble,
	 * scatter, radar, stock, and line chart groups.
	 *
	 * @param b
	 */
	public void setHasShadow( boolean b )
	{
		Serfmt sf = (Serfmt) Chart.findRec( chartArr, Serfmt.class );
		if( sf == null )
		{
			if( b )
			{
				sf = (Serfmt) Serfmt.getPrototype();
				int i = Chart.findRecPosition( chartArr, PieFormat.class );
				sf.setParentChart( getParentChart() );
				chartArr.add( i + 1, sf );
				sf.setHasShadow( true );
			}
		}
		else
		{
			sf.setHasShadow( b );
		}
	}

	/**
	 * percentage=distance of pie slice from center of pie as %
	 *
	 * @param p
	 */
	public void setPercentage( int p )
	{
		PieFormat pf = (PieFormat) Chart.findRec( chartArr, PieFormat.class );
		if( pf == null )
		{
			setHasLines( LineFormat.NONE );
			pf = (PieFormat) Chart.findRec( chartArr, PieFormat.class );
		}
		pf.setPercentage( (short) p );
	}

	/**
	 * return percentage=distance of pie slice from center of pie as %
	 */
	public int getPercentage()
	{
		PieFormat pf = (PieFormat) Chart.findRec( chartArr, PieFormat.class );
		if( pf != null )
		{
			return pf.getPercentage();
		}
		return 0;

	}

	/**
	 * sets the data labels to the desired type:
	 * <li>"ShowValueLabel"
	 * <li>"ShowValueAsPercent"
	 * <li>"ShowLabelAsPercent"
	 * <li>"ShowLabel"
	 * <li>"ShowSeriesName"
	 * <li>"ShowBubbleLabel"
	 *
	 * @param type
	 */
	public void setDataLabels( String type )
	{
		AttachedLabel al = getAttachedLabelRec( true );
		al.setType( type, "1" );
	}

	/**
	 * returns true if has data labels
	 *
	 * @return
	 */
	public boolean getHasDataLabels()
	{
		return (getAttachedLabelRec( false ) != null);
	}

	/**
	 * return if has the specified data label
	 * <li>"ShowValueLabel"
	 * <li>"ShowValueAsPercent"
	 * <li>"ShowLabelAsPercent"
	 * <li>"ShowLabel"
	 * <li>"ShowSeriesName"
	 * <li>"ShowBubbleLabel"
	 *
	 * @param type String option
	 * @return string true or false
	 */
	public String getDataLabelType( String type )
	{
		AttachedLabel al = getAttachedLabelRec( false );
		if( al != null )
		{
			return al.getType( type );
		}
		return null;
	}

	/**
	 * return a string of ALL the label options chosen.  One or more of:
	 * <li>Value
	 * <li>ValuePerecentage
	 * <li>CategoryPercentage	// Pie only
	 * <li>CategoryLabel
	 * <li>BubbleLabel
	 * <li>SeriesLabel
	 *
	 * @return string true or false
	 */
	public String getDataLabelType()
	{
		AttachedLabel al = getAttachedLabelRec( false );
		if( al != null )
		{
			return al.getType();
		}
		return null;
	}

	/**
	 * return the data label int or 0 if no data labels chosen
	 *
	 * @return
	 */
	public int getDataLabelTypeInt()
	{
		AttachedLabel al = getAttachedLabelRec( false );
		if( al != null )
		{
			return al.getTypeInt();
		}
		return 0;
	}

	/**
	 * sets the data labels for the entire chart (as opposed to a specific series/data point).
	 * A combination of:
	 * <li>SHOWVALUE= 0x1;
	 * <li>SHOWVALUEPERCENT= 0x2;
	 * <li>SHOWCATEGORYPERCENT= 0x4;
	 * <li>SHOWCATEGORYLABEL= 0x10;
	 * <li>SHOWBUBBLELABEL= 0x20;
	 * <li>SHOWSERIESLABEL= 0x40;
	 */
	public void setHasDataLabels( int dl )
	{
		AttachedLabel al = getAttachedLabelRec( true );
		al.setType( (short) dl );
	}

	/**
	 * returns type of marker, if any <br>
	 * 0 = no marker <br>
	 * 1 = square <br>
	 * 2 = diamond <br>
	 * 3 = triangle <br>
	 * 4 = X <br>
	 * 5 = star <br>
	 * 6 = Dow-Jones <br>
	 * 7 = standard deviation <br>
	 * 8 = circle <br>
	 * 9 = plus sign
	 *
	 * @return
	 */
	public int getMarkerFormat()
	{
		// default actually looks like: 2, 1, 5, 4 ...
		MarkerFormat mf = (MarkerFormat) Chart.findRec( chartArr, MarkerFormat.class );
		if( mf != null )
		{
			return mf.getMarkerFormat();
		}

		return 0;
	}

	/**
	 * 0 = no marker <br>
	 * 1 = square <br>
	 * 2 = diamond <br>
	 * 3 = triangle <br>
	 * 4 = X <br>
	 * 5 = star <br>
	 * 6 = Dow-Jones <br>
	 * 7 = standard deviation <br>
	 * 8 = circle <br>
	 * 9 = plus sign
	 */
	public void setMarkerFormat( int marker )
	{
		MarkerFormat mf = (MarkerFormat) Chart.findRec( chartArr, MarkerFormat.class );
		if( mf == null )
		{
			if( chartArr.isEmpty() )
			{    // these records come in a set
				Chart3DBarShape cs = new Chart3DBarShape();
				cs.setOpcode( CHART3DBARSHAPE );    // creates default bar shape==0, 0
				addChartRecord( cs );
				LineFormat lf = (LineFormat) LineFormat.getPrototype();
				lf.setParentChart( parentChart );
				lf.setLineStyle( 5 );
				chartArr.add( lf );
				AreaFormat af = (AreaFormat) AreaFormat.getPrototype();
				af.setParentChart( parentChart );
				chartArr.add( af );
				PieFormat pf = (PieFormat) PieFormat.getPrototype();
				pf.setParentChart( parentChart );
				chartArr.add( pf );
				mf = (MarkerFormat) MarkerFormat.getPrototype();
				mf.setParentChart( parentChart );
				chartArr.add( mf );
			}
			else
			{    // shouldn't get here but it goes
				mf = (MarkerFormat) MarkerFormat.getPrototype();
				mf.setParentChart( parentChart );
				int z = Chart.findRecPosition( chartArr, PieFormat.class );
				if( z > -1 )
				{
					chartArr.add( z + 1, mf );
				}
				else    // dunno, add to end
				{
					chartArr.add( mf );
				}
			}
		}
		mf.setMarkerFormat( (short) marker );
	}

// 0x893 CtlMltFrt	
	// [-98, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 0, 13, 19, 0, 6, -128, 34, 0, 35, 0, -67, 64, 0, 0]

	/**
	 * sets the color identified by this DataFormat in the group of records
	 * belonging to the parent series
	 *
	 * @param clr color int
	 */
	public void setSeriesColor( int clr )
	{
		AreaFormat af = getAreaFormatRec( true );
		// Finally got the AreaFormat record that governs the color for this series
		af.seticvFore( clr );
	}

	/**
	 * sets the color identified by this DataFormat in the group of records
	 * belonging to the parent series
	 *
	 * @param clr Hex color string
	 */
	public void setSeriesColor( String clr )
	{
		AreaFormat af = getAreaFormatRec( true );
		// Finally got the AreaFormat record that governs the color for this series
		af.seticvFore( clr );
	}
	
	
/*	
	public int[] getDataLabelsPIE(int defaultdl) {
		 i++;
		      	//ArrayList dls= new ArrayList();
		int[] dls= new int[chartArr.size()-i-1];    	
		int j= 0;
		for (; i < chartArr.size(); i++) {
			if (chartArr.get(i) instanceof DataFormat) {
				DataFormat df= ((DataFormat) chartArr.get(i));
				AttachedLabel al= (AttachedLabel) Chart.findRec(df.chartArr, AttachedLabel.class);
				if (al!=null) 
					dls[j]= al.getTypeInt();
				dls[j++]|=defaultdl;
			} 
		}
		return dls;
	}
*/

	/**
	 * get the bg color identified by this DataFormat
	 * <br>Usually part of a series group of records
	 *
	 * @return
	 */
	public String getBgColor()
	{
		String bg = Frame.getBgColor( chartArr );
		return bg;
	}

	/**
	 * sets the color of the desired pie slice
	 *
	 * @param clr   color int
	 * @param slice 0-based pie slice number
	 */
	public void setPieSliceColor( String clr, int slice )
	{
		AreaFormat af = getAreaFormatPie( slice );
		af.seticvFore( clr );
	}

	/**
	 * sets the color of the desired pie slice
	 *
	 * @param clr   color int
	 * @param slice 0-based pie slice number
	 */
	public void setPieSliceColor( int clr, int slice )
	{
		AreaFormat af = getAreaFormatPie( slice );
		af.seticvFore( clr );
	}

	/**
	 * returns (creates if necessary) the area format for the desired pie slice (pie charts only)
	 *
	 * @param slice int 0-based slice nmber
	 * @return AreaFormat record
	 */
	private AreaFormat getAreaFormatPie( int slice )
	{
		// must add x number of dataformat recs
// FINISH- not 100%    	
/*		if (i==s.chartArr.size()) {
			df= (DataFormat) DataFormat.getPrototypeWithFormatRecs(0);
			df.setPointNumber(slice);
			PieFormat pf= (PieFormat) Chart.findRec(df.chartArr, PieFormat.class);
			pf.setPercentage((short)25);	// default percentage
			AttachedLabel al= (AttachedLabel) AttachedLabel.getPrototype();
			al.setType("CandP");		// default
			df.addChartRecord(al);
			s.chartArr.add(s.chartArr.size()-1, df);	// -1 to skip SERTOCRT rec
		}
*/
		AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
		return af;
	}

	/**
	 * returns (creates if necessary) the area format for this series- controls bar/series colors
	 *
	 * @return AreaFormat record
	 */
	private AreaFormat getAreaFormat()
	{
		AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
		if( af == null )
		{
			af = (AreaFormat) AreaFormat.getPrototype( 0 );
			// NOTE: below list of records is what has been observed in Excel 2003 chart files -
			// unsure if need marker format always ?
			addChartRecord( LineFormat.getPrototype() );
			addChartRecord( af );
			addChartRecord( PieFormat.getPrototype() );
			addChartRecord( MarkerFormat.getPrototype() );
		}
		return af;
	}

}
