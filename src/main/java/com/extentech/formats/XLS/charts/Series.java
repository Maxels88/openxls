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

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.DLbls;
import com.extentech.formats.OOXML.DPt;
import com.extentech.formats.OOXML.Marker;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Format;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.FormatConstantsImpl;
import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;
import java.util.Vector;

/** <b>Series: Series Definition (1003h)</b><br>

 This record defines the Series data of a chart.

 sdtX and sdtY fields determine data type (numeric and text)
 cValx and cValy fields determine number of cells in series

 Offset           Name    Size    Contents
 --
 4               sdtX    2       Type of data in cats (1=num, 3=str)
 8               sdtY    2       Type of data in values (1=num, 3=str)
 10              cValx   2       Count of categories
 12              cValy	2       Count of Values
 14              sdtBSz  2		Type of data in Bubble size series (0=dates, 1=num, 2=seq., 3=text)
 16           sdtValBSz  2		Count of Bubble series vals

 sdtX (2 bytes): An unsigned integer that specifies the type of data in categories (3), or horizontal values on bubble and scatter chart groups, in the series.
 value :0x0001		The series contains categories), or horizontal values on bubble and scatter chart groups, with numeric information.
 value:  0x0003		The series contains categories, or horizontal values on bubble and scatter chart groups, with text information.
 sdtY (2 bytes): An unsigned integer that specifies that the values, or vertical values on bubble and scatter chart groups, in the series contain numeric information.
 MUST be set to 0x0001, and MUST be ignored.
 cValx (2 bytes): An unsigned integer that specifies the count of categories (3), or horizontal values on bubble and scatter chart groups, in the series.
 This value MUST be less than or equal to 0x0F9F.
 cValy (2 bytes): An unsigned integer that specifies the count of values, or vertical values on bubble and scatter chart groups, in the series.
 This value MUST be less than or equal to 0x0F9F.
 sdtBSize (2 bytes): An unsigned integer that specifies that the bubble size values in the series contain numeric information.
 This value MUST be set to 0x0001, and MUST be ignored.
 cValBSize (2 bytes): An unsigned integer that specifies the count of bubble size values in the series.
 This value MUST be less than or equal to 0x0F9F.



 The series object contains a collection of sub objects.  Usually this will take the form of 4 ai records,
 type 0-3, and supporting records such as labels.

 </pre>

 * @see Chart
 */

/**
 * sdtX (2 bytes): An unsigned integer that specifies the type of data in categories (3), or horizontal values on bubble and scatter chart groups, in the series. MUST be a value from the following table.
 * Value 	Meaning
 * 0x0001	The series contains categories, or horizontal values on bubble and scatter chart groups, with numeric information.
 * 0x0003	The series contains categories, or horizontal values on bubble and scatter chart groups, with text information.
 * sdtY (2 bytes): An unsigned integer that specifies that the values, or vertical values on bubble and scatter chart groups, in the series contain numeric information. MUST be set to 0x0001, and MUST be ignored.
 * cValx (2 bytes): An unsigned integer that specifies the count of categories (3), or horizontal values on bubble and scatter chart groups, in the series. This value MUST be less than or equal to 0x0F9F.
 * cValy (2 bytes): An unsigned integer that specifies the count of values, or vertical values on bubble and scatter chart groups, in the series. This value MUST be less than or equal to 0x0F9F.
 * sdtBSize (2 bytes): An unsigned integer that specifies that the bubble size values in the series contain numeric information. This value MUST be set to 0x0001, and MUST be ignored.
 * cValBSize (2 bytes): An unsigned integer that specifies the count of bubble size values in the series. This value MUST be less than or equal to 0x0F9F.
 */
public final class Series extends GenericChartObject implements ChartObject
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7290108485347063887L;
	public static int SERIES_TYPE_NUMERIC = 1;
	public static int SERIES_TYPE_STRING = 3;

	protected int sdtX = -1, sdtY = -1, cValx = -1, cValy = -1, sdtBSz = -1, sdtValBSz = -1;

	private SpPr shapeProps = null;    // OOXML-specific holds the shape properties (line and fill) for this series (all charts)
	private Marker m = null;            // OOXML-specific object to hold marker properties for this series (radar, scatter and line charts only)
	private DLbls d = null;            // OOXML-specific object holds Data Labels properties for this series (all charts except surface)
	private ArrayList dPts = null;    // OOXML-specific object holds Data Labels properties for this series (all charts except surface)

	@Override
	public void init()
	{
		super.init();
		sdtX = (int) ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		sdtY = (int) ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		cValx = (int) ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		cValy = (int) ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		sdtBSz = (int) ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		sdtValBSz = (int) ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );
		if( DEBUGLEVEL > 10 )
		{
			Logger.logInfo( toString() );
		}
	}

	public Series()
	{

	}

	public void update()
	{
		byte[] rkdata = this.getData();
		byte[] b = ByteTools.shortToLEBytes( (short) sdtX );
		rkdata[0] = b[0];
		rkdata[1] = b[1];
		b = ByteTools.shortToLEBytes( (short) sdtY );
		rkdata[2] = b[0];
		rkdata[3] = b[1];
		b = ByteTools.shortToLEBytes( (short) cValx );
		rkdata[4] = b[0];
		rkdata[5] = b[1];
		b = ByteTools.shortToLEBytes( (short) cValy );
		rkdata[6] = b[0];
		rkdata[7] = b[1];
		b = ByteTools.shortToLEBytes( (short) sdtBSz );
		rkdata[8] = b[0];
		rkdata[9] = b[1];
		b = ByteTools.shortToLEBytes( (short) sdtValBSz );
		rkdata[10] = b[0];
		rkdata[11] = b[1];
		this.setData( rkdata );
	}

	/**
	 * Returns the series value AI associated with this series.
	 *
	 * @return
	 */
	public Ai getSeriesValueAi()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_VALS )
				{
					return thisAi;
				}
			}
		}
		return null;
	}

	/**
	 * returns the custom number format for this series, 0 if none
	 *
	 * @return
	 */
	private int getSeriesNumberFormat()
	{
		Ai ai = this.getSeriesValueAi();
		int i = 0;
		if( ai != null )
		{
			i = ai.getIfmt();
			if( i != 0 )
			{
				return i;
			}
			// if 0, number format is determined by the application
			// meaning it uses the number format of the source data
			try
			{
				com.extentech.formats.XLS.formulas.PtgRef p = (com.extentech.formats.XLS.formulas.PtgRef) ai.getCellRangePtgs()[0].getComponents()[0];
				i = ai.getWorkBook().getXf( p.getRefCells()[0].getIxfe() ).getIfmt();
			}
			catch( Exception e )
			{
			}
		}
		return i;
	}

	/**
	 * return the String representation of the numeric format pattern for the series (values) axis
	 *
	 * @return
	 */
	public String getSeriesFormatPattern()
	{
		int ifmt = getSeriesNumberFormat();
		String[][] fmts = FormatConstantsImpl.getBuiltinFormats();
		for( int x = 0; x < fmts.length; x++ )
		{
			if( ifmt == Integer.parseInt( fmts[x][1], 16 ) )
			{
				return fmts[x][0];
			}
		}
		// custom??
		try
		{
			Format fmt = this.getWorkBook().getFormat( ifmt );
			return fmt.getFormat();
		}
		catch( Exception e )
		{
		}
		return "General";
	}

	/**
	 * returns the custom number format for a value-type axis
	 *
	 * @return
	 */
	private int getCategoryNumberFormat()
	{
		Ai ai = this.getCategoryValueAi();
		if( ai != null )
		{
			return ai.getIfmt();
		}
		return 0;
	}

	/**
	 * return the String representation of the numeric format pattern for the Catgeory axis
	 *
	 * @return
	 */
	public String getCategoryFormatPattern()
	{
		int ifmt = getCategoryNumberFormat();
		String[][] fmts = FormatConstantsImpl.getBuiltinFormats();
		for( int x = 0; x < fmts.length; x++ )
		{
			if( ifmt == Integer.parseInt( fmts[x][1], 16 ) )
			{
				return fmts[x][0];
			}
		}
		// custom??
		try
		{
			Format fmt = this.getWorkBook().getFormat( ifmt );
			return fmt.getFormat();
		}
		catch( Exception e )
		{
		}
		return "General";
	}

	/**
	 * sets the legend for this series to a text value
	 *
	 * @param newLegend new text value for legend for the current series
	 * @param wbh       workbookhandle
	 */
	public void setLegend( String newLegend, WorkBookHandle wbh )
	{
		this.getLegendAi().setLegend( newLegend );
		Chart parent = this.getParentChart();
		parent.getChartSeries().legends = null;    // ensure cache is cleared
		parent.getLegend().adjustWidth( parent.getMetrics( wbh ), parent.getChartType(), parent.getChartSeries().getLegends() );
	}

	/**
	 * set legend to a cell ref.
	 *
	 * @param newLegendCell
	 */
	public void setLegendRef( String newLegendCell )
	{
		Ai ai = this.getLegendAi();
		ai.changeAiLocation( ai.toString(), newLegendCell );
		SeriesText st = this.getLegendSeriesText();
		ai.setRt( 2 );
		String legendText = "";
		try
		{
			//CellHandle cell= this.getWorkBook().getCell(newLegendCell);
//    		newLegendCell= newLegendCell.replace('!', ':');	// for this method it's Sheet:cell (????)	
			BiffRec r = ai.getWorkBook().getCell( newLegendCell );
			legendText = r.getStringVal();
		}
		catch( Exception e )
		{
			Logger.logErr( "Series.setLegendRef: Error setting Legend Reference to '" + newLegendCell + "': " + e.toString() );
		}
		st.setText( legendText );
	}

	/**
	 * get legend text
	 *
	 * @return
	 */
	public String getLegendText()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_TEXT )
				{
					try
					{
						com.extentech.formats.XLS.formulas.Ptg[] p = thisAi.getCellRangePtgs();
						return ((com.extentech.formats.XLS.formulas.PtgRef) p[0]).getFormattedValue();
					}
					catch( Exception e )
					{
					}
					try
					{
						if( chartArr.size() > (i + 1) )
						{
							SeriesText st = (SeriesText) chartArr.get( i + 1 );
							if( st != null )
							{
								return st.toString();
							}
						}
					}
					catch( ClassCastException e )
					{
						// couldn't find it!
					}
				}
			}
		}
		return "";
	}

	/**
	 * return the legend cell reference
	 *
	 * @return
	 */
	public String getLegendRef()
	{
		Ai ai = this.getLegendAi();
		if( ai != null )
		{
			return ai.getDefinition();
		}
		return null;
	}

	/**
	 * Return the SeriesText object related to the Legend
	 *
	 * @return
	 */
	protected SeriesText getLegendSeriesText()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_TEXT )
				{
					if( chartArr.size() > (i + 1) )
					{
						try
						{
							SeriesText st = (SeriesText) chartArr.get( i + 1 );
							return st;
						}
						catch( ClassCastException e )
						{
							// couldn't find it!
							return null;
						}
					}
				}
			}
		}
		return null;
	}

	/**
	 * Returns the legend value Ai associated with this series
	 *
	 * @return
	 */
	public Ai getLegendAi()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_TEXT )
				{
					return thisAi;
				}
			}
		}
		return null;
	}

	public Ai getBubbleValueAi()
	{
//    	if (hasBubbleSizes()) {
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_BUBBLES )
				{
					return thisAi;
				}
			}
		}
//    	}
		return null;
	}

	/**
	 * Create a series with all sub components.
	 * <p/>
	 * REcords are...
	 * AI
	 * SeriesText (optional?)
	 * AI
	 * AI
	 * AI
	 * DataFormat
	 * SerToCrt
	 *
	 * @param seriesData
	 * @return
	 */
	protected static Series getPrototype( String seriesRange,
	                                      String categoryRange,
	                                      String bubbleRange,
	                                      String legendRange,
	                                      String legendText,
	                                      ChartType chartobj )
	{
		Series series = (Series) Series.getPrototype();
		Chart parentChart = chartobj.getParentChart();
		WorkBook book = parentChart.getWorkBook();
		Ai ai;
		// create Series text with Legend
		if( legendRange != null )
		{
			ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_LEGEND );
			ai.setWorkBook( book );
			ai.setSheet( parentChart.getSheet() );    // 20080124 KSC: changeAiLocation from "A1" to ""
			try
			{
				ai.changeAiLocation( "", legendRange/*(seriesText.getWorkSheetName() + "!" + seriesText.getCellAddress())*/ );
			}
			catch( Exception e )
			{
				;
			} // it's OK to not have a valid range
			series.addChartRecord( ai );
			SeriesText st = SeriesText.getPrototype( legendText );
			ai.setSeriesText( st );
			series.addChartRecord( st );
			// 20091102 KSC: when adding series legend will not expand correctly if autopositioning is turned off
			// [BugTracker 2844]
			Legend l = chartobj.getDataLegend();
			if( l != null )
			{
				l.setAutoPosition( true );
				l.incrementHeight( parentChart.getCoords()[3] );
			}

		}
		else
		{
			ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_NULL_LEGEND );
			ai.setWorkBook( book );
			ai.setSheet( parentChart.getSheet() );
			series.addChartRecord( ai );
		}
//        parentChart.addAi(ai);        
		// create Series Value Ai
		ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_SERIES );
		ai.setParentChart( parentChart );
		ai.setWorkBook( book );
		ai.setSheet( parentChart.getSheet() );
		try
		{
			ai.changeAiLocation( ai.toString(), seriesRange );
		}
		catch( Exception e )
		{
// not necessary to report        	Logger.logErr("Error setting Series Range: '"  + seriesRange + "'-" + e.toString()); 
		} // it's OK to not have a valid range
		series.addChartRecord( ai );
//        parentChart.addAi(ai);        
		// create Category Ai
		ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_CATEGORY );
		ai.setWorkBook( book );
		ai.setSheet( parentChart.getSheet() );
		try
		{
			ai.changeAiLocation( ai.toString(), categoryRange );
		}
		catch( Exception e )
		{
			;
		} // it's OK to not have a valid range
		series.addChartRecord( ai );
//        parentChart.addAi(ai);        
		// create Bubble (undocumented) Ai
		ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_BUBBLE );
		ai.setWorkBook( book );
		ai.setSheet( parentChart.getSheet() );
		if( !bubbleRange.equals( "" ) )
		{
			try
			{
				ai.changeAiLocation( ai.toString(), bubbleRange );
			}
			catch( Exception e )
			{
				;
			} // it's OK to not have a valid range
			ai.setRt( 2 );
/* 20120123 KSC: confused vs. per-series format verses format in chartformat rec ...
 *              if (((BubbleChart) parentChart.getChartObject()).is3d()) {
            	
            }*/
		}
		series.addChartRecord( ai );
		DataFormat df = null;
		df = (DataFormat) DataFormat.getPrototype();
		// update the data format correctly
		Vector ser = parentChart.getAllSeries();    // get ALL series
		int yi = -1;        // Changed from 0
		int iss = -1;    // ""
		for( int i = 0; i < ser.size(); i++ )
		{
			Series srs = (Series) ser.get( i );
			int newYi = srs.getSeriesIndex();
			int newIss = srs.getSeriesNumber();
			if( newYi > yi )
			{
				yi = newYi;
			}
			if( newIss > iss )
			{
				iss = newIss;
			}
		}
		yi++;
		iss++;
		df.setSeriesIndex( yi );
		df.setSeriesNumber( iss );
		if( chartobj.getBarShape() != 0 )
		{ // must ensure each series contains proper shape records
			df.setShape( chartobj.getBarShape() );
		}
		series.addChartRecord( df );
		SerToCrt stc = (SerToCrt) SerToCrt.getPrototype();
		// get the correct chart index for the sertocrt
		int vCount = 0;
		int cCount = 0;
		int bCount = 0;
		if( ser.size() > 0 )
		{    // 20070709 KSC: will be 0 if adding new blank chart
			Series s = (Series) ser.get( 0 );
			ArrayList cr = s.getChartRecords();
			for( int i = 0; i < cr.size(); i++ )
			{
				BiffRec b = (BiffRec) cr.get( i );
				if( b.getOpcode() == SERTOCRT )
				{
					SerToCrt stcc = (SerToCrt) b;
					stc.setData( stcc.getData() );
				}
			}
			// set the series level variables correctly
			vCount = s.getValueCount();
			cCount = s.getCategoryCount();
			bCount = s.getBubbleCount();
		}
		series.init();
		// 20070711 KSC: vCount and cCount are via current range, no??
		try
		{
			if( seriesRange.indexOf( ":" ) != -1 )
			{
				int coords[] = com.extentech.ExtenXLS.ExcelTools.getRangeCoords( seriesRange );
				vCount = coords[4];
			}
			else
			{
				vCount = 1;
			}
			series.setValueCount( vCount );
		}
		catch( Exception e )
		{
		}
		try
		{
			cCount = com.extentech.ExtenXLS.ExcelTools.getRangeCoords( categoryRange )[4];
			series.setCategoryCount( cCount );
		}
		catch( Exception e )
		{
		}
		if( !bubbleRange.equals( "" ) )
		{
			try
			{
				bCount = com.extentech.ExtenXLS.ExcelTools.getRangeCoords( bubbleRange )[4];
				series.setBubbleCount( bCount );
			}
			catch( Exception e )
			{
			}
		}
		series.addChartRecord( stc );
		return series;
	}

	public static XLSRecord getPrototype()
	{
		Series s = new Series();
		s.setOpcode( SERIES );
		s.setData( s.PROTOTYPE_BYTES );
		s.init();
		return s;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 3, 0, 1, 0, 3, 0, 3, 0, 1, 0, 0, 0 };

	/**
	 * Returns the category value AI associated with this series.
	 *
	 * @return
	 */
	public Ai getCategoryValueAi()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_CATEGORIES )
				{
					return thisAi;
				}
			}
		}
		return null;
	}

	/**
	 * Get the series index (file relative)
	 *
	 * @return
	 */
	protected int getSeriesIndex()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return (int) df.getSeriesIndex();
		}
		return -1;
	}

	/**
	 * Get the series Number (display)
	 *
	 * @return
	 */
	protected int getSeriesNumber()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return (int) df.getSeriesNumber();
		}
		return -1;
	}

	protected int getCategoryCount()
	{
		return cValx;
	}

	protected int getValueCount()
	{
		return cValy;
	}

	public void setCategoryCount( int i )
	{ //20070711 KSC: changed from protected
		cValx = i;
		this.update();
	}

	protected int getBubbleCount()
	{
		return sdtValBSz;
	}

	public void setBubbleCount( int i )
	{ //20070711 KSC: changed from protected
		sdtValBSz = i;
		this.update();
	}

	public void setValueCount( int i )
	{    //20070711 KSC: changed from protected
		cValy = i;
		this.update();
	}

	// 20070712 KSC: get/set for data types
	public int getCategoryDataType()
	{
		return sdtX;
	}

	public int getValueDataType()
	{
		return sdtY;
	}

	public void setCategoryDataType( int i )
	{
		sdtX = i;
		this.update();
	}

	public void setValueDataType( int i )
	{
		sdtY = i;
		this.update();
	}

	public boolean hasBubbleSizes()
	{
		return sdtValBSz > 0;
	}

	/**
	 * Gets the dataformat record associated with this Series, if any.
	 * If none present, option to create a basic DataFormat set of records DataFormat controls
	 *
	 * @return DataFormat Record
	 */
	private DataFormat getDataFormatRec( boolean bCreate )
	{
		DataFormat df = (DataFormat) Chart.findRec( this.chartArr, DataFormat.class );
		if( (df == null) && bCreate )
		{ // create dataformat
			df = (DataFormat) DataFormat.getPrototypeWithFormatRecs( this.getParentChart() );
			this.addChartRecord( df );
		}
		return df;
	}

	/**
	 * retrieves the data format record which corresponds to the desired pie slice
	 *
	 * @param slice
	 * @return
	 */
	private DataFormat getDataFormatRecSlice( int slice, boolean bCreate )
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df == null )
		{
			Logger.logErr( "Series.getDataFormatRecSlice: cannot find data format record" );
			return null;
		}
		int seriesNumber = df.getSeriesNumber();
		// for PIE charts, DataFormats are stored in 1st series, 
		// just after the initial data format rec
		// must check point number to see if it's the desired df
		Series s;
		if( seriesNumber == 0 ) // we're on the first one
		{
			s = this;
		}
		else
		{    // should not have more than 1 series for a pie chart!
			s = (Series) getParentChart().getAllSeries().get( seriesNumber );
		}
		int i = Chart.findRecPosition( s.chartArr, DataFormat.class ); // get position of the first df
		i++;    // skip 1st
		int lastSlice = 0;
		while( i < s.chartArr.size() )
		{
			if( s.chartArr.get( i ) instanceof DataFormat )
			{
				df = (DataFormat) s.chartArr.get( i );
				lastSlice = df.getPointNumber();
				if( df.getPointNumber() == slice )
				{
					return df;
				}
			}
			i++;
		}
		// create
		if( bCreate )
		{
			i--;
			while( lastSlice <= slice )
			{
				df = (DataFormat) DataFormat.getPrototypeWithFormatRecs( this.getParentChart() );
				df.setPointNumber( lastSlice++ );
				df.setParentChart( this.getParentChart() );
				s.chartArr.add( i++, df );
			}
			return df;
		}
		return null;
	}

	/**
	 * returns the shape of the data point for this series
	 *
	 * @return
	 */
	public int getShape()
	{
		int ret = 0;
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			ret = df.getShape();
		}
		return ret;
	}

	public void setShape( int shape )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setShape( shape );
	}

	/**
	 * returns true if this series has smoothed lines
	 *
	 * @return
	 */
	public boolean getHasSmoothedLines()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getSmoothedLines();
		}
		return false;
	}

	/**
	 * set smooth lines setting (applicable for line, scatter charts)
	 *
	 * @param smooth
	 */
	public void setHasSmoothLines( boolean smooth )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setSmoothLines( smooth );
	}

	/**
	 * sets this series to have lines
	 * <br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 */
	public void setHasLines( int lineStyle )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setHasLines( lineStyle );
	}

	/**
	 * sets the color for this series
	 * NOTE: for PIE Charts, use setPieSliceColor
	 *
	 * @param clr String color hex String
	 * @see setPieSliceColor
	 */
	public void setColor( String clr )
	{
		int type = this.getParentChart().getChartType();
		if( type == ChartConstants.PIECHART )
		{
			setPieSliceColor( clr, this.getSeriesIndex() );
			return;
		}
		DataFormat df = this.getDataFormatRec( true );
		df.setSeriesColor( clr );
	}

	/**
	 * if the exact/correct color index is not used, the fill color comes out black
	 *
	 * @param clr
	 * @return
	 */
	private int ensureCorrectColorInt( int clr )
	{
		// "The Chart color table is a subset of the full color table"
		if( clr == FormatConstants.COLOR_RED )
		{
			clr = FormatConstants.COLOR_RED_CHART;
		}
		if( clr == FormatConstants.COLOR_BLUE )
		{
			clr = FormatConstants.COLOR_BLUE_CHART;
		}
		if( clr == FormatConstants.COLOR_YELLOW )
		{
			clr = FormatConstants.COLOR_YELLOW_CHART;
		}
		if( clr == FormatConstants.COLOR_DARK_GREEN )    // no standard chart dark green ...
		{
			clr = FormatConstants.COLOR_GREEN;
		}
		if( clr == FormatConstants.COLOR_DARK_YELLOW )    // no standard chart dark yellow ...
		{
			clr = FormatConstants.COLOR_YELLOW_CHART;
		}
		if( clr == FormatConstants.COLOR_OLIVE_GREEN )
		{
			clr = FormatConstants.COLOR_OLIVE_GREEN_CHART;
		}
		if( clr == FormatConstants.COLOR_WHITE )
		{
			clr = FormatConstants.COLOR_WHITE3;
		}
		return clr;

	}

	/**
	 * sets the color for this series
	 * NOTE: for PIE Charts, use setPieSliceColor
	 *
	 * @param clr color int
	 * @see setPieSliceColor
	 */
	public void setColor( int clr )
	{
		clr = ensureCorrectColorInt( clr );
		int type = this.getParentChart().getChartType();
		if( type == ChartConstants.PIECHART )
		{
			setPieSliceColor( clr, this.getSeriesIndex() );
			return;
		}
		DataFormat df = this.getDataFormatRec( true );
		df.setSeriesColor( clr );

	}

	/**
	 * sets the color of the desired pie slice
	 *
	 * @param clr   color int
	 * @param slice 0-based pie slice number
	 */
	public void setPieSliceColor( int clr, int slice )
	{
		clr = ensureCorrectColorInt( clr );
		int type = this.getParentChart().getChartType();
		if( type != ChartConstants.PIECHART )
		{
			return;
		}
		DataFormat df = this.getDataFormatRecSlice( slice, true );
		if( df != null )
		{
			df.setPieSliceColor( clr, slice );
		}
		else
		{
			Logger.logErr( "Series.setPieSliceColor: unable to fnd pie slice record" );
		}
	}

	/**
	 * sets the color of the desired pie slice
	 *
	 * @param clr   color int
	 * @param slice 0-based pie slice number
	 */
	public void setPieSliceColor( String clr, int slice )
	{
		int type = this.getParentChart().getChartType();
		if( type != ChartConstants.PIECHART )
		{
			return;
		}
		DataFormat df = this.getDataFormatRecSlice( slice, true );
		if( df != null )
		{
			df.setPieSliceColor( clr, slice );
		}
		else
		{
			Logger.logErr( "Series.setPieSliceColor: unable to fnd pie slice record" );
		}
	}

	// Periwinkle 	Plum+ 	Ivory 	Light Turquoise 	Dark Purple 	Coral 	Ocean Blue 	Ice Blue  {17, 25, 19, 27, 28, 22, 23, 24};
	// try these color int numbers instead:
	// alternative explanation:  chart fills 16-23, chart lines 24-31
	public static int automaticSeriesColors[] = {
			24,
			25,
			26,
			27,
			28,
			29,
			30,
			31,
			32,
			33,
			34,
			35,
			36,
			37,
			38,
			39,
			40,
			41,
			42,
			43,
			44,
			45,
			46,
			48,
			49,
			50,
			51,
			52,
			53,
			54,
			55,
			56,
			57,
			58,
			59,
			60,
			61,
			62
	};        // also used in mapping default colors in AreaFormat, MarkerFormat, Frame ...

	/**
	 * retrieve the series/bar color for this series
	 * NOTE: for Pie Charts, must use getPieSliceColor
	 *
	 * @return color int
	 * @see getPieSliceColor
	 */
	public String getSeriesColor()
	{
		DataFormat df = this.getDataFormatRec( false );
		int type = this.getParentChart().getChartType();
		int seriesNumber = df.getSeriesNumber();
		if( type == ChartConstants.PIECHART )
		{
			return FormatHandle.colorToHexString( FormatHandle.COLORTABLE[automaticSeriesColors[seriesNumber]] );
		}
		String bg = df.getBgColor();
		if( bg != null )
		{
			return bg;
		}
		// otherwise, color is automatic or default chart series color
		return FormatHandle.colorToHexString( FormatHandle.COLORTABLE[automaticSeriesColors[seriesNumber]] );
	}

	/**
	 * get the pie slice color in this pie chart
	 *
	 * @param slice
	 * @return color int
	 */
	public String getPieSliceColor( int slice )
	{
		int type = this.getParentChart().getChartType();
		if( type != ChartConstants.PIECHART )
		{
			return null;
		}
		DataFormat df = this.getDataFormatRecSlice( slice, false );
		if( df != null )
		{
			String bg = df.getBgColor();
			if( bg != null )
			{
				return bg;
			}
		}
		// otherwise, color is automatic or default chart series color
		return FormatHandle.colorToHexString( FormatHandle.COLORTABLE[automaticSeriesColors[slice]] );
	}

	/**
	 * return the type of markers for each series:
	 * <br>0 = no marker
	 * <br>1 = square
	 * <br>2 = diamond
	 * <br>3 = triangle
	 * <br>4 = X
	 * <br>5 = star
	 * <br>6 = Dow-Jones
	 * <br>7 = standard deviation
	 * <br>8 = circle
	 * <br>9 = plus sign
	 */
	public int getMarkerFormat()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getMarkerFormat();
		}
		return 0;
	}

	/**
	 * return data label options as an int
	 * <br>can be one or more of:
	 * <br>SHOWVALUE= 0x1;
	 * <br>SHOWVALUEPERCENT= 0x2;
	 * <br>SHOWCATEGORYPERCENT= 0x4;
	 * <br>SMOOTHEDLINE= 0x8;
	 * <br>SHOWCATEGORYLABEL= 0x10;
	 * <br>SHOWBUBBLELABEL= 0x20;
	 * <br>SHOWSERIESLABEL= 0x40;
	 *
	 * @return a combination of data label options above or 0 if none
	 * @see AttachedLabel
	 */
	public int getDataLabel()
	{
		int datalabels = 0;
		DataLabExtContents dl = (DataLabExtContents) Chart.findRec( this.chartArr, DataLabExtContents.class );
		if( dl != null )
		{ // Extended Label -- add to attachedlabel, if any
			datalabels = dl.getTypeInt();
		}
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			datalabels |= df.getDataLabelTypeInt();
		}
		return datalabels;
	}

	/**
	 * PIE data label information is contained within the 1st series only
	 * <br>TODO: not implemented yet
	 * return data label options as an int
	 * <br>can be one or more of:
	 * <br>SHOWVALUE= 0x1;
	 * <br>SHOWVALUEPERCENT= 0x2;
	 * <br>SHOWCATEGORYPERCENT= 0x4;
	 * <br>SMOOTHEDLINE= 0x8;
	 * <br>SHOWCATEGORYLABEL= 0x10;
	 * <br>SHOWBUBBLELABEL= 0x20;
	 * <br>SHOWSERIESLABEL= 0x40;
	 *
	 * @param defaultdl int default data label setting for overall chart
	 * @return int array of data labels for each pie slice
	 * @see AttachedLabel
	 */
	public int[] getDataLabelsPIE( int defaultdl )
	{
		int datalabels = 0;
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			datalabels |= df.getDataLabelTypeInt();
		}
		return null;
	}

	/**
	 * return the OOXML shape property for this series
	 *
	 * @return
	 */
	public SpPr getSpPr()
	{
		return shapeProps;
	}

	/**
	 * set the OOXML shape properties for this series
	 *
	 * @param sp
	 */
	public void setSpPr( SpPr sp )
	{
		shapeProps = sp;
	}

	/**
	 * return the OOXML marker properties for this series
	 *
	 * @return
	 */
	public Marker getMarker()
	{
		return m;
	}

	/**
	 * set the OOXML marker properties for this series
	 *
	 * @param Sp
	 */
	public void setMarker( Marker m )
	{
		this.m = m;
	}

	/**
	 * return the OOXML dLbls (data labels) properties for this series
	 *
	 * @return
	 */
	public DLbls getDLbls()
	{
		return d;
	}

	/**
	 * set the OOXML dLbls (data labels) properties for this series
	 *
	 * @param Sp
	 */
	public void setDLbls( DLbls d )
	{
		this.d = d;
	}

	/**
	 * return OOXML dPt (data points) for this series
	 *
	 * @return
	 */
	public DPt[] getDPt()
	{
		if( dPts != null )
		{
			return (DPt[]) dPts.toArray( new DPt[]{ } );
		}
		return null;
	}

	/**
	 * add a dPt element (data point) for this series
	 *
	 * @param Sp
	 */
	public void addDpt( DPt d )
	{
		if( dPts == null )
		{
			dPts = new ArrayList();
		}
		dPts.add( d );
	}

	/**
	 * returns the val OOXML element that defines the values for the series values
	 *
	 * @param valstr either "val" or "yval" for scatter/bubble charts
	 * @return
	 */
	public StringBuffer getValOOXML( String valstr )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:" + valstr + ">" );
		ooxml.append( "\r\n" );

		Ai seriesAi = null;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_VALS )
				{
					seriesAi = thisAi;
					break;
				}

			}
		}

		ooxml.append( "<c:numRef>" );
		ooxml.append( "\r\n" );        // number reference
		ooxml.append( "<c:f>" + OOXMLAdapter.stripNonAscii( seriesAi.toString() ) + "</c:f>" );
		ooxml.append( "\r\n" );    // string range
		// Need numCache for chart lines apparently 
		ooxml.append( "<c:numCache>" );
		ooxml.append( "\r\n" );        // specifies the last data shown on the chart for a series
		// formatCode	== format pattern
		ooxml.append( "<c:formatCode>" + this.getSeriesFormatPattern() + "</c:formatCode>" );
		CellRange cr = new CellRange( seriesAi.toString(), parentChart.wbh, false );
		CellHandle[] ch = cr.getCells();
		// ptCount	== point count
		ooxml.append( getValueRangeOOXML( ch ) );
		// pt * n  == a Numeric Point each has a <v> child, an idx attribute and an optional formatcode attribute
		ooxml.append( "</c:numCache>" );
		ooxml.append( "\r\n" );        //

		ooxml.append( "</c:numRef>" );
		ooxml.append( "\r\n" );
		ooxml.append( "</c:" + valstr + ">" );
		ooxml.append( "\r\n" );
		return ooxml;

	}

	/**
	 * return the cat OOXML element used to define a series category
	 *
	 * @param cat    string category cell range for the given series --
	 *               almost always the same for each series (except for scatter/bubble charts)
	 * @param catstr either "cat" or "xval" for scatter/bubble charts
	 *               cat elements must contain string references
	 *               xval contain numeric references
	 * @return
	 */
	public StringBuffer getCatOOXML( String cat, String catstr )
	{
		StringBuffer ooxml = new StringBuffer();
		if( !"".equals( cat ) )
		{  // causes 1004 vb error upon Excel SAVE - BAXTER SAVE BUG
			ooxml.append( "<c:" + catstr + ">" );
			ooxml.append( "\r\n" );            // categories contain a string "formula" ref + string caches
			if( catstr.equals( "cat" ) )
			{
				ooxml.append( "<c:strRef>" );    // string reference
			}
			else
			{
				ooxml.append( "<c:numRef>" );    // number reference
			}
			ooxml.append( "\r\n" );
			ooxml.append( "<c:f>" + OOXMLAdapter.stripNonAscii( cat ) + "</c:f>" );
			ooxml.append( "\r\n" );
			/* 20090211 KSC: if errors in referenced cells whole chart will error; best to avoid caching at all */
			if( catstr.equals( "cat" ) )
			{
				ooxml.append( "</c:strRef>" );    // string reference
			}
			else
			{
				ooxml.append( "</c:numRef>" );    // number reference
			}
			ooxml.append( "\r\n" );
			ooxml.append( "</c:" + catstr + ">" );
			ooxml.append( "\r\n" );
		}
		else
		{
			// TESTING-- remove when done
			//Logger.logWarn("ChartHandle.getOOXML: null category found- skipping");
		}
		return ooxml;
	}

	/**
	 * returns the bubbleSize OOXML element that defines the values for the series values
	 *
	 * @param isBubble3d true if it's a 3d bubble chart
	 * @return
	 */
	public StringBuffer getBubbleOOXML( boolean isBubble3d )
	{
		StringBuffer ooxml = new StringBuffer();
		Ai bubbleAi = null;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			if( br.getOpcode() == AI )
			{
				Ai thisAi = (Ai) br;
				if( thisAi.getType() == Ai.TYPE_BUBBLES )
				{
					bubbleAi = thisAi;
					break;
				}
			}
		}

		ooxml.append( "<c:bubbleSize>" );
		ooxml.append( "\r\n" );
		ooxml.append( "<c:numRef>" );
		ooxml.append( "\r\n" );        // number reference
		ooxml.append( "<c:f>" + bubbleAi.toString() + "</c:f>" );
		ooxml.append( "\r\n" );
		ooxml.append( "<c:numCache>" );
		ooxml.append( "\r\n" );
		try
		{
			CellHandle[] cells = CellRange.getCells( bubbleAi.toString(), parentChart.wbh );
			ooxml.append( getValueRangeOOXML( cells ) );
			ooxml.append( "\r\n" );
		}
		catch( NumberFormatException e )
		{
			Logger.logErr( "geteriesOOXML: Number format exception for Bubble Range: " + bubbleAi.toString() );
		}
		ooxml.append( "</c:numCache>" );
		ooxml.append( "\r\n" );
		ooxml.append( "</c:numRef>" );
		ooxml.append( "\r\n" );
		ooxml.append( "</c:bubbleSize>" );
		ooxml.append( "\r\n" );
		if( isBubble3d )
		{
			ooxml.append( "<c:bubble3D val=\"1\"/>" );
		}
		ooxml.append( "\r\n" );
		return ooxml;
	}

	/**
	 * returns the bubbleSize OOXML element that defines the values for the series values
	 *
	 * @param isBubble3d true if it's a 3d bubble chart
	 * @return
	 */
	public StringBuffer getLegendOOXML( boolean from2003 )
	{
		StringBuffer ooxml = new StringBuffer();
		String txt = this.getLegendText();
		Ai ai = this.getLegendAi();
/*       String txt= null;
       try {
       	com.extentech.formats.XLS.formulas.Ptg[] p= ai.getCellRangePtgs();
       	txt= ((com.extentech.formats.XLS.formulas.PtgRef)p[0]).getFormattedValue();
       } catch (Exception e) {}
       try {
           if (chartArr.size()>i+1) {
              SeriesText st = (SeriesText)chartArr.get(i+1);
              if (st!=null)
           	   txt= st.toString();
           }
       }catch(ClassCastException e) {
           // couldn't find it!
       }
*/

		ooxml.append( "<c:tx>" );
		ooxml.append( "\r\n" );
		ooxml.append( "<c:strRef>" );
		ooxml.append( "\r\n" );        // string reference
		if( ai != null )
		{
			ooxml.append( "<c:f>" + OOXMLAdapter.stripNonAscii( ai.getDefinition() ) + "</c:f>" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:strCache>" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:ptCount val=\"1\"/>" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:pt idx=\"0\">" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:v>" + OOXMLAdapter.stripNonAscii( txt ) + "</c:v>" );
			ooxml.append( "</c:pt>" );
			ooxml.append( "\r\n" );
			ooxml.append( "</c:strCache>" );
			ooxml.append( "\r\n" );
			ooxml.append( "</c:strRef>" );
			ooxml.append( "\r\n" );
			ooxml.append( "</c:tx>" );
			ooxml.append( "\r\n" );
			if( this.getSpPr() != null )
			{
				ooxml.append( this.getSpPr().getOOXML() );
			}
			else if( from2003 )
			{
				SpPr ss;
				if( parentChart.getChartType() != RADARCHART )
				{
					ss = new SpPr( "c", this.getSeriesColor().substring( 1 ), 12700, "000000" );
				}
				else
				{
					ss = new SpPr( "c", null, 25400, this.getSeriesColor().substring( 1 ) );
				}
				ooxml.append( ss.getOOXML() );
			}
		}
		return ooxml;
	}

	/**
	 * generate the OOXML used to represent this set of value cells (element numRef)
	 *
	 * @param cells
	 * @return
	 */
	private static String getValueRangeOOXML( CellHandle[] cells )
	{
		StringBuffer ooxml = new StringBuffer();
		if( cells == null )
		{
			return "<c:ptCount val=\"0\"/>";
		}
		ooxml.append( "<c:ptCount val=\"" + cells.length + "\"/>" );
		ooxml.append( "\r\n" );
		for( int j = 0; j < cells.length; j++ )
		{
			ooxml.append( "<c:pt idx=\"" + j + "\">" );
			ooxml.append( "\r\n" );
			if( !cells[j].getStringVal().equals( "NaN" ) )
			{
				ooxml.append( "<c:v>" + cells[j].getStringVal() + "</c:v>" );
				ooxml.append( "\r\n" );
			}
			else
			{    // appears that NaN is an invalid entry
				ooxml.append( "<c:v>0</c:v>" );
				ooxml.append( "\r\n" );
			}
			ooxml.append( "</c:pt>" );
			ooxml.append( "\r\n" );
		}
		return ooxml.toString();
	}

}