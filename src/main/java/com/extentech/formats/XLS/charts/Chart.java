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

import com.extentech.ExtenXLS.ChartHandle;
import com.extentech.ExtenXLS.ChartHandle.ChartOptions;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.ByteStreamer;
import com.extentech.formats.XLS.Dimensions;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.MSODrawing;
import com.extentech.formats.XLS.Obj;
import com.extentech.formats.XLS.Sheet;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import org.json.JSONArray;
import org.json.JSONException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

/**
 * <b>Chart: Chart Location and Dimensions 0x1002h</b><br>
 * <p/>
 * The Chart record determines the chart dimensions and
 * marks the beginning of the Chart records.
 * <p/>
 * <p><pre>
 * * Note that all these values are split up 2 bytes integer and 2 bytes fractional
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       x           4       x pos of upper left corner
 * 8       y           4       y pos of upper left corner
 * 12      dx          4       x-size
 * 16      dy          4       y-size
 * <p/>
 * </p></pre>
 * <p/>
 * <p/>
 * <p/>
 * notes on implementation:
 * <p/>
 * <br>A chart may have up to 9 chart types (specific instances of a chart e.g. bar, line, etc.) known as chart groups.  These are identified by the chartgroup array list.
 * <br>The parent of these chart groups are the axis group.
 * <br>A chart's series and trendlines contains indexes into the chart group.
 *
 * @see Ai
 */
public class Chart extends GenericChartObject implements ChartObject
{
	private static final Logger log = LoggerFactory.getLogger( Chart.class );
	static final long serialVersionUID = 6702247464633674375l;

	// Objects which define a chart -- REMEMBER TO UPDATE OOXMLChart copy constructor if modify below member vars
	protected ArrayList<ChartType> chartgroup = new ArrayList();    // can have up to 4 chart objects (Bar, Area, Line, etc.) per a given chart
	protected int nCharts = 0;    // total number of charts (0= the default chart + any overlay charts (up to 4 charts (9 in OOXML) see chartgroup list)
	protected ChartAxes chartaxes = null;
	protected ChartSeries chartseries = new ChartSeries();
	protected TextDisp charttitle = null;
	protected Dimensions dimensions;        // datarange dimensions ...
	// The following two objects are external objects in the XLS stream that are associated with the chart.
	protected Obj obj = null;
	protected MSODrawing msodrawobj = null;    // stores coordinates and shape info ...
	HashMap<String, Double> chartMetrics = new HashMap();        // hold chartMetrics: x, y, w, h, canvasw, canvash 
	public transient WorkBookHandle wbh;

	// TODO: MANAGE FONTS ---?  

	// internal chart records
	protected ArrayList chartRecs = new ArrayList();
	protected AbstractList preRecs;
	protected AbstractList postRecs = new ArrayList();
	protected boolean dirtyflag = false;    // if anything has changed in the chart (except series, which is handled via another var) 
	protected boolean metricsDirty = true;    // initially true so creates min, max and other metrics, true if should be recalculated   
	// below vars used to save state in addInitialChartRecord recursion
	protected Ai currentAi;                // used in init only
	protected int hierarchyDepth = 0;    //	""
	protected ArrayList initobs = new ArrayList();

    
    /*
    public Chart copy()  throws InstantiationException, IllegalAccessException  {
        Chart copy = this.getClass().newInstance();
        return copy;
     }*/

	@Override
	public void init()
	{
		super.init();
		getData();
		chartseries.setParentChart( this );
		/**
		 * NOTE: the x, y, dx (w) and dy (h) are in most cases
		 * 		not used; the MSODRAWING coords governs the 
		 * 		chart canvas; see getMetrics for plot and other adjustments
		 *
		 * New Info:  have to see if this works
		 * Get chart area width in pixels
		 chart area width in pixels = (dx field of Chart record - 8) * DPI of the display device / 72
		 If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
		 chart area width in pixels -= 2 * line width of the display device in pixels

		 Get chart area height in pixels
		 chart area height in pixels = (dy field of Chart record - 8) * DPI of the display device / 72
		 If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
		 chart area height in pixels -= 2 * line height of the display device in pixels

		 NOTE:
		 * Since the 1980s, the Microsoft Windows operating system has set the default display "DPI" to 96 PPI, 
		 * while Apple/Macintosh computers have used a default of 72 PPI.[2] 
		 *
		 byte[] rkdata = this.getData();
		 try {
		 short s = ByteTools.readShort(rkdata[0],rkdata[1]);
		 short ss = ByteTools.readShort(rkdata[2],rkdata[3]);

		 if(ss<0)ss=0;
		 String parser = s + "." + ss;
		 x = (new Float(parser)).floatValue();

		 s = ByteTools.readShort(rkdata[4],rkdata[5]);
		 ss = ByteTools.readShort(rkdata[6],rkdata[7]);
		 if(ss<0)ss=0;

		 parser = s + "." + ss;
		 y = (new Float(parser)).floatValue();


		 /*Value of the real number = Integral + (Fractional / 65536.0)
		 *  Integral (2 bytes): A signed integer that specifies the integral part of the real number.
		 Fractional (2 bytes): An unsigned integer that specifies the fractional part of the real number.
		 *
		 int integral= 0;
		 integral |= rkdata[8] & 0xFF;
		 integral <<= 8;
		 integral |= rkdata[9] & 0xFF;
		 int fractional= ByteTools.readUnsignedShort(rkdata[10],rkdata[11]);
		 dx= (float)(integral + (fractional/65536.0));
		 integral= ((rkdata[12] & 0xFF) << 8) | (rkdata[13] & 0xFF);
		 if (integral < 0) integral += 256;
		 fractional= ByteTools.readUnsignedShort(rkdata[14],rkdata[15]);
		 dy= (float) integral;
		 } catch (NumberFormatException e) {	// 20080414 KSC: parsing appears to be off - TODO Check out Chart record specs
		 log.error("Chart.init: parsing dimensions failed: " + e);
		 }

		 if ((DEBUGLEVEL > 3))
		 Logger.logInfo("Chart Found @ x:" + x + " y:" + y + " dx:" + dx + " dy:" + dy);*/
	}

    
    /*
     * Note, other important chart records we are not as of yet trapping:
     * CttMlFrt 0x89E -- additional properties for chart elements
     * ShapePropsStream	0x8A4
     * TextPropsStream	0x8A5
     * CrtLayout12  	0x89D	-- layout info for attached label, legend
     * 0x89E=The CrtMlFrt record specifies additional properties for chart elements, as specified by the Chart Sheet SubstreamABNF. These properties complement the record to which they correspond, and are stored as a structure chain defined in XmlTkChain. An application can ignore this record without loss of functionality, except for the additional properties. If this record is longer than 8224 bytes, it MUST be split into several records. The first section of the data appears in this record and subsequent sections appear in one or more CrtMlFrtContinue records that follow this record. 
     * 
     * Legend --> Pos, TextDisp, Frame, CrtLayout12, TextPropsStream, CrtLayout12 
     */

	/**
	 * Adds chart records to the chartRecs array.  This array should be completed with all chart records
	 * before the initChartRecords method is called, creating the hierarchical chart structure.
	 *
	 * @param br - a chart record
	 * @return false if we reach the final "end" record, signifying the end of the chart object structure.
	 * <p/>
	 * Be aware that there can be additional chart records after this structure,  Chart only deals with it's own heirarchy.
	 */
	public boolean addInitialChartRecord( BiffRec br )
	{
		if( br.getOpcode() == 0x1033 )
		{
			++hierarchyDepth;
		}
		else if( br.getOpcode() == 0x1034 )
		{
			--hierarchyDepth;
			if( hierarchyDepth == 0 )
			{
				chartRecs.add( br );
				return false;
			}
		}
		if( hierarchyDepth != 0 )
		{
			if( br.getOpcode() == AI )
			{
				currentAi = (Ai) br;
				currentAi.setParentChart( this );  // necessary for base ops
				currentAi.setSheet( this.getSheet() ); // needed for cell change updates
			}
			else if( (currentAi != null) && (br.getOpcode() == SERIESTEXT) )
			{
				currentAi.setSeriesText( (SeriesText) br );
				currentAi = null;    // done 
			}
			else if( br.getOpcode() == SERIES )
			{
				chartseries.add( new Object[]{
						br, nCharts
				} );    // store series records - initially map all to default chart (normal case)
			}
			else if( br.getOpcode() == AXISPARENT )
			{
				chartaxes = new ChartAxes( (AxisParent) br );
			}
			else if( br.getOpcode() == AXIS )
			{
				chartaxes.add( (Axis) br );
			}
			else if( br.getOpcode() == CHARTFORMAT )
			{    // usually only 1 (default chart id= 0) but can have up to 9 overlay charts 
				nCharts++;
			}
			else if( br.getOpcode() == SERIESLIST )
			{ // maps chart # to series #            	
				chartseries.addSeriesMapping( nCharts - 1,
				                              ((SeriesList) br).getSeriesMappings() );// 1-based series #'s to "assign" to the overlay chart
			}
			try
			{
				if( ((GenericChartObject) br).chartType != -1 )
				{ // store the chart object which defines this chart (NOTE: may be up to 4 charts in one = overlay charts)
					ChartFormat cf = null;
					for( int i = chartRecs.size() - 1; i >= 0; i-- )
					{
						if( ((BiffRec) chartRecs.get( i )).getOpcode() == XLSConstants.CHARTFORMAT )
						{
							cf = (ChartFormat) chartRecs.get( i );
							break;
						}
					}
// sadly, can't do it here - have to wait until process hierarchy chartobj.add(ChartType.getChartObject((GenericChartObject)br, cf, this.getWorkBook()));
					initobs.add( new BiffRec[]{ br, cf } );
				} // Set Parent Chart as it's necessary for all basic chart ops 
				((GenericChartObject) br).setParentChart( this );  // Note: Scl is not a GenericChartObject + can have unknown (XLSRecords) - will cause ClassCastException
			}
			catch( ClassCastException c )
			{
				;
			}
			chartRecs.add( br );
		}
		else
		{
			if( br.getOpcode() == DIMENSIONS )
			{
				dimensions = (Dimensions) br;
			}
			br.getData();
			postRecs.add( br );
		}

		return true;
	}

	public void setDirtyFlag( boolean b )
	{
		dirtyflag = b;
	}

	/**
	 * Take the initial array of records for the chart, and create a
	 * hierarchial array of objects for the chart.  Also, populate initial values for easy access.
	 */
	public void initChartRecords()
	{
		try
		{    // Added to find/set obj + msodrawing records - replaces setObj 
			// see TestReadWrite.TestIOOBError
			// For Normal charts, Obj rec is the last record before the Chart record
			if( !this.getSheet().isChartOnlySheet() )
			{
				int pos = this.getSheet().getSheetRecs().size() - 1;
				BiffRec rec = (BiffRec) this.getSheet().getSheetRecs().get( pos );
				obj = (Obj) rec;
				obj.setChart( this );
				// Usually, MsoDrawing is just preceding Obj record, except in those rare carses where
				// there are Continues and Txo's ...
				while( --pos > 0 )
				{
					rec = (BiffRec) this.getSheet().getSheetRecs().get( pos );
					if( rec.getOpcode() == MSODRAWING )
					{
						this.msodrawobj = (MSODrawing) rec;
						break;
					}
				}
			} // chart-only worksheets have no obj/mso apparently
		}
		catch( Exception e )
		{
			log.error( "initChartRecords: Error in Chart Records:  " + e.toString() );
		}

		// Turn it into a static array initially to speed up random access
		BiffRec[] bArr = new BiffRec[chartRecs.size()];
		bArr = (BiffRec[]) chartRecs.toArray( bArr );
		this.initChartObject( this, bArr );
		for( Object initob : initobs )
		{
			BiffRec[] ios = (BiffRec[]) initob;
			chartgroup.add( ChartType.createChartTypeObject( (GenericChartObject) ios[0], (ChartFormat) ios[1], this.getWorkBook() ) );
			Legend l = (Legend) Chart.findRec( ((ChartFormat) ios[1]).chartArr, Legend.class );
			if( l != null )
			{
				chartgroup.get( chartgroup.size() - 1 ).addLegend( l );
			}
		}
		initobs = new ArrayList();    // clear out 
	}

	/**
	 * Handle the iteration and creation of the chart objects
	 */
	private ChartObject initChartObject( ChartObject cobj, BiffRec[] cRecs )
	{

		for( int i = 0; i < cRecs.length; i++ )
		{
			BiffRec b = cRecs[i];
			b.getData();
			if( !(b.getOpcode() == BEGIN) && !(b.getOpcode() == END) )
			{
				if( (cRecs.length > (i + 1)) && ((cRecs[i + 1]).getOpcode() == BEGIN) )
				{ // this is an object with sub-data
					try
					{
						ChartObject co = (ChartObject) b;
						int endloc = this.getMatchingEndRecordLocation( i, cRecs );
						int arrlen = endloc - i;
						BiffRec[] objArr = new BiffRec[arrlen];
						System.arraycopy( cRecs, i + 1, objArr, 0, arrlen );
						cobj.addChartRecord( (XLSRecord) this.initChartObject( co, objArr ) );
						// necessary initialization of key elements 
						if( co instanceof TextDisp )
						{
							int type = ((TextDisp) co).getType();
							// -1 = default text ...?
							if( type == ObjectLink.TYPE_TITLE )
							{
								charttitle = (TextDisp) co;
							}
							else if( type == ObjectLink.TYPE_XAXIS )
							{
								chartaxes.setTd( XAXIS, (TextDisp) co );
							}
							else if( type == ObjectLink.TYPE_YAXIS )
							{
								chartaxes.setTd( YAXIS, (TextDisp) co );
							}
							else if( type == ObjectLink.TYPE_ZAXIS )
							{
								chartaxes.setTd( ZAXIS, (TextDisp) co );
							}
							else if( type == ObjectLink.TYPE_DATAPOINTS ) // series or data points
							{
								; // do what??
							}
							else if( type == ObjectLink.TYPE_DISPLAYUNITS ) //
							// KSC: TESTING!! Take out when done
//                            	Logger.logInfo("Display Units");
							{
								;
							}
						}
						try
						{
							co.setParentChart( this );
						}
						catch( Exception e )
						{
						}
						i += arrlen;
					}
					catch( ClassCastException e )
					{
						// it's not a defined chart object.  Add it in!!!  If we are missing a chart object containing other records we
						// will not be able to write these out correctly.
						log.error( "Error in parsing chart. Please add the correct object (opcode: " + b.getOpcode() + ") to be a Chart Object" );
					}
				}
				else
				{
					cobj.addChartRecord( (XLSRecord) b );
				}
			}
		}
		return cobj;
	}

	/**
	 * Just a little helper method!  Determines where the matching end record is for a begin record
	 *
	 * @param startLoc
	 * @param cRecs
	 * @return
	 */
	private int getMatchingEndRecordLocation( int startLoc, BiffRec[] cRecs )
	{
		int offset = 0;
		for( int i = startLoc + 2; i < cRecs.length; i++ )
		{
			BiffRec b = cRecs[i];
			if( b.getOpcode() == BEGIN )
			{
				offset++;
			}
			if( b.getOpcode() == END )
			{
				offset--;
				if( offset < 0 )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * put together the chart records for output to the streamer
	 *
	 * @return arranged List of Chart Records.
	 */
	public List assembleChartRecords()
	{
		Vector outputVec = new Vector();
		if( preRecs != null )
		{
			outputVec.addAll( preRecs );
		}
		if( dirtyflag )
		{
			for( int i = 0; i < nCharts; i++ )
			{
				chartseries.updateSeriesMappings( chartgroup.get( i ).getSeriesList(),
				                                  i );    // if has multiple or overlay charts, update series mappings
			}
			outputVec.addAll( this.getRecordArray() );
		}
		else
		{
			outputVec.add( this );
			outputVec.addAll( chartRecs );
		}
		if( postRecs != null )
		{
			outputVec.addAll( postRecs );
		}
		return outputVec;
	}

	/**
	 * This handles setting the records external to, but supporting the chart object hierarchy.
	 * <p/>
	 * Think of this kind of like the recvec & cell records in the XLS parsing we do.
	 * <p/>
	 * We have a pre and post array of biffrecs that are (for now)  unmodifieable.
	 * These get appended into the whole array on ouput in append chart records.
	 *
	 * @param recs
	 */
	public void setPreRecords( AbstractList recs )
	{
		preRecs = recs;
	}

	/**
	 * Returns a map of series ==> series ptgs
	 * Ranges can either be in the format "C5" or "Sheet1!C4:D9"
	 */
	public HashMap getSeriesPtgs()
	{
		return this.chartseries.getSeriesPtgs();
	}

	/**
	 * Returns a serialized copy of this Chart
	 */
	public byte[] getChartBytes()
	{
		for( int i = 0; i < preRecs.size(); i++ )
		{
			((BiffRec) preRecs.get( i )).getData();
		}

		ObjectOutputStream obs = null;
		byte[] b = null;
		try
		{
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			obs = new ObjectOutputStream( baos );
			obs.writeObject( this );
			b = baos.toByteArray();
		}
		catch( IOException e )
		{
			log.error( "Error obtaining chart bytes", e );
		}
		return b;
	}

	public byte[] getSerialBytes()
	{
		for( int i = 0; i < preRecs.size(); i++ )
		{
			((BiffRec) preRecs.get( i )).getData();
		}
		ObjectOutputStream obs = null;
		byte[] b = null;
		try
		{
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			BufferedOutputStream bufo = new BufferedOutputStream( baos );
			obs = new ObjectOutputStream( bufo );
			obs.writeObject( this );
			bufo.flush();
			b = baos.toByteArray();
		}
		catch( IOException e )
		{
			log.error( "Error getting serial bytes", e );
		}
		return b;
	}

	/**
	 * Gets the records for the chart, including the global references, ie the MSODrawing and Obj records
	 *
	 * @return
	 */
	public List getXLSrecs()
	{
		List l = this.assembleChartRecords();
		if( obj != null )
		{
			l.add( 0, obj );
		}
		if( msodrawobj != null )
		{
			l.add( 0, msodrawobj );
		}
		return l;
	}

	/**
	 * Dimensions rec holds the datarange for the chart; must update ...
	 *
	 * @return
	 */
	public int getMinRow()
	{
		return dimensions.getRowFirst();
	}

	public int getMaxRow()
	{
		return dimensions.getRowLast();
	}

	public int getMinCol()
	{
		return dimensions.getColFirst();
	}

	public int getMaxCol()
	{
		return dimensions.getColLast();
	}

	public void setDimensionsRecord( int r0, int r1, int c0, int c1 )
	{
		dimensions.setRowFirst( r0 - 1 );
		dimensions.setRowLast( r1 - 1 );
		dimensions.setColFirst( c0 );
		dimensions.setColLast( c1 - 1 );
		/* also must remove label and numberrec cached records if altered series
			otherwise causes errors when removing or altering series 
	     */
		for( int i = postRecs.size() - 1; i > 0; i-- )
		{
			BiffRec b = (BiffRec) postRecs.get( i );
			int op = b.getOpcode();
			if( (op == NUMBER) || (op == LABEL) )
			{
				postRecs.remove( i );
			}
		}
	}

	public void setDimensions( Dimensions d )
	{
		dimensions = d;
	}

	public void setDimensionsRecord()
	{
		Vector serieslist = this.getAllSeries( -1 );
		int nSeries = serieslist.size();
		int nPoints = 0;
		for( Object aSerieslist : serieslist )
		{
			try
			{
				Series s = (Series) aSerieslist;
				int[] coords = ExcelTools.getRangeCoords( s.getSeriesValueAi().getDefinition() );
				if( coords[3] > coords[1] )
				{
					nPoints = Math.max( nPoints, (coords[3] - coords[1]) + 1 );    // c1-c0
				}
				else
				{
					nPoints = Math.max( nPoints, (coords[2] - coords[0]) + 1 );    // r1-r0
				}
			}
			catch( Exception e )
			{
			}
		}
		this.setDimensionsRecord( 0, nPoints, 0, nSeries );
	}

	/**
	 * returns the unique Obj Id # of this chart as seen in Vb macros et. al
	 *
	 * @return in id #
	 */
	public int getId()
	{
		if( this.obj != null )
		{
			return this.obj.getObjId();
		}
		return -1;
	}

	/**
	 * sets the unique id # for this chart
	 * used upon addChart
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( this.obj != null )
		{
			this.obj.setObjId( id );
		}
	}

	/**
	 * flags chart metrics should
	 * to be recalculated
	 */
	public void setMetricsDirty()
	{
		dirtyflag = true;
		metricsDirty = true;
	}

	/**
	 * Change series ranges for ALL matching series
	 *
	 * @param originalrange
	 * @param newrange
	 * @return
	 */
	public boolean changeSeriesRange( String originalrange, String newrange )
	{
		setMetricsDirty();
		return chartseries.changeSeriesRange( originalrange, newrange );
	}

	/**
	 * Return a string representing all series in this chart
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL charts
	 * @return
	 */
	public String[] getSeries( int nChart )
	{
		return chartseries.getSeries( nChart );
	}

	/**
	 * Return an array of strings, one for each category
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for All
	 * @return
	 */
	public String[] getCategories( int nChart )
	{
		return chartseries.getCategories( nChart );
	}

	/**
	 * get all the series objects for ALL charts
	 * (i.e. even in overlay charts)
	 *
	 * @return
	 */
	public Vector getAllSeries()
	{
		return chartseries.getAllSeries( -1 );
	}

	/**
	 * get all the series objects in the specified chart (-1 for ALL)
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL series
	 * @return
	 */
	public Vector getAllSeries( int nChart )
	{
		return chartseries.getAllSeries( nChart );
	}

	/**
	 * Add a series object to the array.
	 *
	 * @param seriesRange   = one row range, expressed as (Sheet1!A1:A12);
	 * @param categoryRange = category range
	 * @param bubbleRange=  bubble range, if any 20070731 KSC
	 * @param seriesText    = label for the series;
	 * @param nChart        0=default, 1-9= overlay charts
	 *                      s     * @return
	 */
	public Series addSeries( String seriesRange,
	                         String categoryRange,
	                         String bubbleRange,
	                         String legendRange,
	                         String legendText,
	                         int nChart )
	{
		Series s = chartseries.addSeries( seriesRange,
		                                  categoryRange,
		                                  bubbleRange,
		                                  legendRange,
		                                  legendText,
		                                  chartgroup.get( nChart ),
		                                  nChart );
		setMetricsDirty();
		return s;
	}

	/**
	 * return an array of legend text
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL
	 */
	public String[] getLegends( int nChart )
	{
		return chartseries.getLegends( nChart );
	}

	/**
	 * specialty method to take absoulte index of series and remove it
	 * <br>only used in WorkSheetHandle
	 *
	 * @param index
	 */
	public void removeSeries( int index )
	{
		setMetricsDirty();
		chartseries.removeSeries( index );
	}

	/**
	 * remove desired series from chart
	 *
	 * @param index
	 */
	public void removeSeries( Series seriestodelete )
	{
		// remove from cache
		setMetricsDirty();
		chartseries.removeSeries( seriestodelete );
		// remove from chartArray
		int nSeries = -1;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == SERIES )
			{
				nSeries++;
				if( seriestodelete.equals( b ) )
				{    // this is the one to delete
					chartArr.remove( i );
					// now adjust series number for all subsequent Series
					for( int j = i; j < chartArr.size(); j++ )
					{
						b = chartArr.get( j );
						if( b.getOpcode() == SERIES )
						{
							Series s = (Series) b;
							int x = Chart.findRecPosition( s.chartArr, DataFormat.class );
							if( x > 0 )
							{
								DataFormat df = (DataFormat) s.chartArr.get( x );
								df.setSeriesIndex( nSeries++ );
							}
						}
						else    // we've got 'em all
						{
							break;
						}
					}
					// now make sure referenced label and number recs are removed
					/** how does this work, it does not reference a particular series label, just all of them?
					 for (int j= 0; j < postRecs.size(); j++) {
					 if ((postRecs.get(j) instanceof NumberRec) ||
					 (postRecs.get(j) instanceof Label)) {
					 postRecs.remove(j);
					 j--;
					 }

					 }
					 */
					break;
				}
			}
		}
	}

	/**
	 * get a chart series handle based off the series name
	 *
	 * @param seriesName
	 * @param nChart     0=default, 1-9= overlay charts
	 * @return
	 */
	public Series getSeries( String seriesName, int nChart )
	{
		return chartseries.getSeries( seriesName, nChart );
	}

	/**
	 * @param originalrange
	 * @param newrange
	 * @return
	 */
	public boolean changeCategoryRange( String originalrange, String newrange )
	{
		setMetricsDirty();
		return chartseries.changeCategoryRange( originalrange, newrange );
	}

	public boolean changeTextValue( String originalval, String newval )
	{
		setMetricsDirty();
		return chartseries.changeTextValue( originalval, newval );
	}

	/**
	 * returns the plot area background color string (hex)
	 *
	 * @return color string
	 */
	public String getPlotAreaBgColor()
	{
		String bg = null;
		try
		{
			Frame f = (Frame) Chart.findRec( this.chartArr, Frame.class );
			bg = f.getBgColor();
		}
		catch( Exception e )
		{
		}
		int ct = this.getChartType();
		if( (bg == null) || ((ct != PIECHART) && (ct != RADARCHART)) )
		{
			bg = chartaxes.getPlotAreaBgColor();
		}
		if( bg == null )
		{
			bg = "#FFFFFF";
		}
		return bg;
	}

	/**
	 * sets the plot area background color
	 *
	 * @param bg color int
	 */
	public void setPlotAreaBgColor( int bg )
	{
		this.chartaxes.setPlotAreaBgColor( bg );
		setMetricsDirty();
	}

	/**
	 * returns the plot area line color string (hex)
	 *
	 * @return color string
	 */
	public String getPlotAreaLnColor()
	{
		String clr = "#000000";
		try
		{
			Frame f = (Frame) Chart.findRec( this.chartArr, Frame.class );
			clr = f.getLineColor();
		}
		catch( Exception e )
		{
		}
		return clr;
	}

	/**
	 * Get the title of the Chart
	 *
	 * @return title of the Chart
	 */
	public String getTitle()
	{
		if( charttitle != null )
		{
			return charttitle.toString();
		}
		return "";
		// getTitleTD();
	}

	/**
	 * return the title TextDisp element
	 *
	 * @return
	 */
	public TextDisp getTitleTd()
	{
		return charttitle;
	}

	/**
	 * return the textdisp object that defines the chart title
	 *
	 * @return public TextDisp getTitleTD(){
	 * if(title != null){
	 * return title;
	 * }
	 * for (int i=0;i<chartArr.size();i++) {
	 * BiffRec b = (BiffRec)chartArr.get(i);
	 * if (b.getOpcode()==TEXTDISP) {
	 * TextDisp td = (TextDisp)b;
	 * if (td.isChartTitle()){
	 * title = td;
	 * }
	 * }
	 * }
	 * return title;
	 * }
	 */

	public String toString()
	{
		String t = getTitle();
		if( !t.equals( "" ) )
		{
			return t;
		}
		return "Untitled Chart";
	}

	/**
	 * return the default font for the chart
	 *
	 * @return
	 */
	public Font getDefaultFont()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			if( chartArr.get( i ).getOpcode() == DEFAULTTEXT )
			{
				if( ((DefaultText) chartArr.get( i )).getType() == 2 )
				{// correct code???? 2 or 3 ?? 
					TextDisp td = (TextDisp) chartArr.get( i + 1 );
					int idx = td.getFontId();
					if( idx > -1 )
					{
						return wkbook.getFont( idx );
					}
					return null;
				}
			}
		}
		return null;
	}

	/**
	 * Set the default font for the specific DefaultText rec
	 *
	 * @param type
	 * @param fontId
	 */
	public void setDefaultFont( int type, int fontId )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b.getOpcode() == TEXTDISP )
			{
				//xxxxx            
			}
		}
		dirtyflag = true;
	}

	/**
	 * Get all of the Fontx records in this chart.  Utilized for
	 * updating font references when moving between workbooks
	 */
	public ArrayList getFontxRecs()
	{
		ArrayList ret = new ArrayList();
		for( Object chartRec : chartRecs )
		{
			BiffRec b = (BiffRec) chartRec;
			if( b.getOpcode() == FONTX )
			{
				ret.add( b );
			}
		}
		return ret;
	}

	/**
	 * resets all fonts in the chart to the default font of the workbook Useful
	 * for OOXML charts which are based upon default chart
	 */
	public void resetFonts()
	{
		for( Object chartRec : chartRecs )
		{
			BiffRec b = (BiffRec) chartRec;
			if( b.getOpcode() == FONTX )
			{
				((Fontx) b).setIfnt( 0 );
			}
		}
	}

	/**
	 * return the font associated with the chart title
	 *
	 * @return Font
	 */
	public Font getTitleFont()
	{
		if( charttitle == null )
		{
			return null;
		}
		return charttitle.getFont( wkbook );
	}

	/**
	 * Utilized by external chart copy, this
	 * sets references internally in this chart which allows cross-workbook
	 * identification of worksheets
	 * to occur.
	 */
	public void populateForTransfer()
	{
		Iterator i = this.getAllSeries( -1 ).iterator();
		while( i.hasNext() )
		{
			Series s = (Series) i.next();
			try
			{
				for( int j = 0; j < s.chartArr.size(); j++ )
				{
					if( s.chartArr.get( j ).getOpcode() == AI )
					{
						((Ai) s.chartArr.get( j )).populateForTransfer( this.getSheet().getSheetName() );
					}
				}
			}
			catch( Exception e )
			{
				log.error( "Chart.populateForTransfer: " + e.toString() );
			}
		}

	}

	/**
	 * update Sheet References contained in associated Series/Category Ai's
	 * using saved Original Sheet References
	 */
	public void updateSheetRefs( String newSheetName, String origWorkBook )
	{
		Iterator i = this.getAllSeries( -1 ).iterator();
		while( i.hasNext() )
		{
			Series s = (Series) i.next();
			try
			{
				for( int j = 0; j < s.chartArr.size(); j++ )
				{
					if( s.chartArr.get( j ).getOpcode() == AI )
					{
						((Ai) s.chartArr.get( j )).updateSheetRef( newSheetName,
						                                           origWorkBook );    // 20080630 KSC: add sheet name for lookup
					}
				}
			}
			catch( Exception e )
			{
				log.error( "Chart.updateSheetRefs: " + e.toString() );
			}
		}
	}

	/**
	 * Set the title of the Chart
	 *
	 * @param str title of the Chart
	 */
	public void setTitle( String str )
	{
		// No need to create a TD record if setting title to ""
		if( ((str == null) || str.equals( "" )) && (charttitle == null) )    // if no title and setting title to "", just leave it
		{
			return;
		}
		if( charttitle == null )
		{ // 20070709 KSC: Adding a new chart, add Title recs ...        	
			try
			{
				TextDisp td = (TextDisp) TextDisp.getPrototype( ObjectLink.TYPE_TITLE, str, this.getWorkBook() );
				this.addChartRecord( td );    // add TextDisp title to end of chart recs ...
				charttitle = td;

			}
			catch( Exception e )
			{
				log.error( "Unable to set title of chart to: " + str + " This chart object does not contain a title record" );
			}
		}
		else
		{ // just set title text
			charttitle.setText( str );
		}
		setMetricsDirty();
	}

	/**
	 * returns the chart type for the default chart
	 */
	public int getChartType()
	{
		return chartgroup.get( 0 ).getChartType();
	}

	/**
	 * returns the chart type for the desired chart
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return
	 */
	public int getChartType( int nChart )
	{
		return chartgroup.get( nChart ).getChartType();
	}

	/**
	 * returns an array of all chart types found in this chart
	 * 0= default, 1-9 (max) are overlay charts
	 *
	 * @return
	 */
	public int[] getAllChartTypes()
	{
		int[] charttypes = new int[nCharts];    // max        	
		for( int i = 0; i < nCharts; i++ )
		{
			charttypes[i] = chartgroup.get( i ).getChartType();
		}
		return charttypes;
	}

	/**
	 * returns the number of charts in this chart -
	 * 1= default, can have up to 9 overlay charts as well
	 *
	 * @return
	 */
	public int getNumberOfCharts()
	{
		return nCharts - 1;
	}

	/**
	 * return the default Chart Object associated with this chart
	 *
	 * @return
	 */
	public ChartType getChartObject()
	{
		return chartgroup.get( 0 );
	}

	/**
	 * return the nth Chart Object associated with this chart (default=0, overlay=1 thru 9)
	 *
	 * @return
	 */
	public ChartType getChartObject( int nChart )
	{
		return chartgroup.get( nChart );
	}

	/**
	 * gets the integral order of the specific chart obect in question
	 *
	 * @param ct Chart Type rwobject
	 * @return
	 */
	public int getChartOrder( ChartType ct )
	{
		for( int i = 0; i < chartgroup.size(); i++ )
		{
			if( chartgroup.get( i ).equals( ct ) )
			{
				return i;
			}
		}
		return 0;
	}

	/**
	 * create internal records necessary for an overlay chart
	 */
	private ChartFormat createNewChart( int nChart )
	{
		AxisParent ap = (AxisParent) findRec( chartArr, AxisParent.class );
		return ap.createNewChart( nChart );
	}

	/**
	 * adds a new Chart Type to the group of Chart Type Objects, or replaces one at existing index
	 *
	 * @param ct     Chart Type Object
	 * @param nChart index
	 */
	protected void addChartType( ChartType ct, int nChart )
	{
		if( nChart < chartgroup.size() )
		{
			chartgroup.remove( nChart );
		}
		else
		{
			nCharts++;
		}
		chartgroup.add( nChart, ct );
		dirtyflag = true;
	}

	/**
	 * retrieves the parent of the chart type object (==ChartFormat) at index nChart, or creates a new Chart Type Parent
	 *
	 * @param nChart
	 * @return ChartFormat
	 */
	protected ChartFormat getChartOjectParent( int nChart )
	{
		ChartFormat cf = null;
		if( nChart >= chartgroup.size() )  // create new
		{
			cf = this.createNewChart( nChart );
		}
		else
		{
			cf = chartgroup.get( nChart ).cf;
		}
		return cf;
	}

	/**
	 * changes this chart to be a specific chart type with specific display options
	 * <br>for multiple charts, specify nChart 1-9, for the default chart,
	 * nChart= 0
	 *
	 * @param chartType             chart type one of: BARCHART, LINECHART, AREACHART, COLCHART, PIECHART,
	 *                              DONUGHTCHART, RADARCHART, RADARAREACHART, PYRAMIDCHART, CONECHART, CYLINDERCHART,
	 *                              SURFACTCHART
	 * @param nChart                chart # 0 is default
	 * @param EnumSet<ChartOptions> 0 or more chart options (Such as Stacked, Exploded ...)
	 * @see ChartOptions
	 */
	public void setChartType( int chartType, int nChart, EnumSet<ChartOptions> options )
	{
		GenericChartObject c;

		ChartFormat cf = getChartOjectParent( nChart );
		ChartType ct = ChartType.create( chartType, this, cf );
		ct.setOptions( options );

		// save and reset legend:
		Legend l = this.getLegend();
		ct.addLegend( l );

		if( (ct instanceof BubbleChart) && options.contains( ChartOptions.THREED ) )
		{
			((BubbleChart) ct).setIs3d( true );        // when set, every series created will
		}

		addChartType( ct, nChart );
    	

/* TODO: axes, other chart options ... Legend???
			// The axis group MUST contain two value axes if and only if all chart groups are of type bubble or scatter.	
    		} else if (chartType==BUBBLECHART || chartType==SCATTERCHART) {
// TODO: FINISH!!! IS NOT CORRECT UNLESS PROPER AXES ARE CREATED     			
    	    //The axis group MUST contain a category or date axis if the axis group contains an area, bar, column, filled radar, line, radar, or surface chart group    		
    		} else { // ensure has proper axes   
// TODO: FINISH!!! IS NOT CORRECT UNLESS PROPER AXES ARE CREATED     			
    		}
*/
		/**
		 The axis group MUST contain a series axis if and only if the chart group attached to the axis group is one of the following:
		 An area chart group with the fStacked field of the Area record equal to 0.
		 A column chart group with the fStacked field of the Bar record equal to 0 and the fClustered field of the Chart3d record equal to 0.
		 A line chart group with field fStacked of the Line record equal to 0.
		 A surface chart group
		 The chart group on the axis group MUST contain a Chart3d record if the axis group contains a series axis.
		 */
		/**
		 * for overlay charts i.e. multiple charts, restrictions:
		 * Because there are many different ways to represent data visually, each representation has specific requirements about the layout of the data and the way it is plotted. 
		 * This results in restrictions on the combinations of chart group types that can be plotted on the same axis group, and the combinations of chart group types that can 
		 be plotted in the same chart.
		 A chart MUST contain one of the following:			
		 A single axis group that contains a single chart group that contains a Chart3d record.		
		 One or two axis groups that each contain a single bubble chart group.			
		 One or two axis groups that each conform to one of the following restrictions on chart group type combinations:			
		 Zero or one of each of the following chart group types: area, column, line, and scatter.			
		 Zero or one of each of the following chart group types: bar of pie, doughnut, pie, and pie of pie.			
		 A single bar chart group.			
		 A single filled radar chart group.			
		 A single radar chart group.

		 */
	}

	/**
	 * @return truth of "Chart is Three-D"
	 */
	public boolean isThreeD()
	{
		return isThreeD( 0 );
	}

	/**
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return truth of "Chart is Three-D"
	 */
	public boolean isThreeD( int nChart )
	{
		return chartgroup.get( nChart ).isThreeD();
	}

	/**
	 * return the 3d rec for this chart or null if it doesn't exist
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return
	 */
	public ThreeD getThreeDRec( int nChart )
	{
		return chartgroup.get( nChart ).getThreeDRec( false );
	}

	/**
	 * @return truth of "Chart is Stacked"
	 */
	@Override
	public boolean isStacked()
	{
		return isStacked( 0 );
	}

	/**
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return truth of "Chart is Stacked"
	 */
	public boolean isStacked( int nChart )
	{
		return chartgroup.get( nChart ).isStacked();
	}

	/**
	 * return truth of "Chart is 100% Stacked"
	 *
	 * @return
	 */
	public boolean is100PercentStacked()
	{
		return is100PercentStacked( 0 );
	}

	/**
	 * return truth of "Chart is 100% Stacked"
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return
	 */
	public boolean is100PercentStacked( int nChart )
	{
		return chartgroup.get( nChart ).is100PercentStacked();
	}

	/**
	 * @return truth of "Chart is Clustered"  (Bar/Col only)
	 */
	public boolean isClustered()
	{
		return isClustered( 0 );
	}

	/**
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return truth of "Chart is Clustered"  (Bar/Col only)
	 */
	public boolean isClustered( int nChart )
	{
		return chartgroup.get( nChart ).isClustered();
	}

	/**
	 * return chart-type-specific options in XML form
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return String XML
	 */
	public String getChartOptionsXML( int nChart )
	{
		return chartgroup.get( nChart ).getChartOptionsXML();
	}

	/**
	 * interface for setting chart-type-specific options
	 * in a generic fashion
	 *
	 * @param op  option
	 * @param val value
	 * @see ExtenXLS.handleChartElement
	 * @see ChartHandle.getXML
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		dirtyflag = true;
		return this.setChartOption( op, val, 0 );
	}

	/**
	 * interface for setting chart-type-specific options
	 * in a generic fashion
	 *
	 * @param op     option
	 * @param val    value
	 * @param nChart 0=default, 1-9= overlay charts
	 * @see ExtenXLS.handleChartElement
	 * @see ChartHandle.getXML
	 */
	public boolean setChartOption( String op, String val, int nChart )
	{
		dirtyflag = true;
		return chartgroup.get( nChart ).setChartOption( op, val );
	}

	/**
	 * get the value of *almost* any chart option (axis options are in Axis)
	 * for the default chart
	 *
	 * @param op String option e.g. Shadow or Percentage
	 * @return String value of option
	 */
	@Override
	public String getChartOption( String op )
	{
		return chartgroup.get( 0 ).getChartOption( op );
	}

	/**
	 * return ThreeD settings for this chart in XML form
	 *
	 * @return String XML
	 */
	public String getThreeDXML()
	{
		return getThreeDXML( 0 );
	}

	/**
	 * return ThreeD settings for this chart in XML form
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return String XML
	 */
	public String getThreeDXML( int nChart )
	{
		return chartgroup.get( nChart ).getThreeDXML();
	}

	/**
	 * returns the 3d record of the desired chart
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return
	 */
	public ThreeD initThreeD()
	{
		return initThreeD( 0, this.getChartType( 0 ) );
	}

	/**
	 * returns the 3d record of the desired chart
	 *
	 * @param nChart    0=default, 1-9= overlay charts
	 * @param chartType one of the chart type constants
	 * @return
	 */
	public ThreeD initThreeD( int nChart, int chartType )
	{
		return chartgroup.get( nChart ).initThreeD( chartType );
	}

	/**
	 * Return the Chart Axes object
	 *
	 * @return ChartAxis
	 */
	public ChartAxes getAxes()
	{
		return this.chartaxes;
	}

	/**
	 * return the Chart Series object
	 *
	 * @return
	 */
	public ChartSeries getChartSeries()
	{
		return chartseries;
	}

	/**
	 * Returns true if the Y Axis (Value axis) is set to automatic scale
	 * <p>The default setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale (minimum, maximum values
	 * plus major and minor tick units) as necessary
	 *
	 * @return boolean true if Automatic Scaling is turned on
	 * @see getAxisAutomaticScale(int axisType)
	 */
	public boolean getAxisAutomaticScale()
	{
		return this.chartaxes.getAxisAutomaticScale();
	}

	/**
	 * returns the minimum and maximum values by examining all series on the chart
	 *
	 * @param bk
	 * @return double[] min, max
	 */
	public double[] getMinMax( WorkBookHandle wbh )
	{
		if( metricsDirty )
		{
			getMetrics( wbh );
		}
		if( this.wbh == null )
		{
			this.wbh = wbh;
		}
		metricsDirty = false;
		double[] minmaxcache = chartseries.getMetrics( metricsDirty );    // Ignore Overlay charts for now!
		return minmaxcache;
	}

	/**
	 * Return all the Chart-specific font recs in XML form
	 * These include the default fonts + title font + axis fonts ...
	 *
	 * @return
	 */
	public String getChartFontRecsXML()
	{
		HashMap fonts = new HashMap();
		int maxFont = 5;
		for( Object chartRec : chartRecs )
		{
			BiffRec b = (BiffRec) chartRec;
			if( b.getOpcode() == FONTX )
			{
				int fontId = ((Fontx) b).getIfnt();        //((TextDisp) b).getFontId();
				maxFont = Math.max( fontId, maxFont );
				Font f = this.getWorkBook().getFont( fontId );
				fonts.put( fontId, f.getXML() );
			}
		}
		StringBuffer sb = new StringBuffer();
		for( int i = 5; i <= maxFont; i++ )
		{
			if( fonts.get( Integer.valueOf( i ) ) != null )
			{
				sb.append( "\n\t\t<ChartFontRec id=\"" + i + "\" " );
				sb.append( fonts.get( Integer.valueOf( i ) ) );
				sb.append( "/>" );
			}
		}
		return sb.toString();
	}

	/**
	 * return the fontid for non-axis chart fonts (title, default ...)
	 *
	 * @return
	 */
	public String getChartFontsXML()
	{
		StringBuffer sb = new StringBuffer();
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == DEFAULTTEXT )
			{
				sb.append( " Default" + ((DefaultText) b).getType() + "=\"" );
				i++;
				b = chartArr.get( i );
				if( b.getOpcode() == TEXTDISP )
				{ // should!!
					TextDisp td = (TextDisp) b;
					sb.append( td.getFontId() );
				}
				sb.append( "\"" );
			}
			else if( b.getOpcode() == TEXTDISP )
			{
				TextDisp td = (TextDisp) b;
				if( td.isChartTitle() )
				{ // should!!
					sb.append( " Title=\"" );
					sb.append( td.getFontId() );
					sb.append( "\"" );
				}
			}
		}
		return sb.toString();
	}

	/**
	 * Set the specific chart font (title, default ...)
	 * NOTE: Axis fonts are handled separately
	 *
	 * @param type
	 * @param val
	 * @see Axis.setChartOption, TextDisp.setChartOption
	 */
	public void setChartFont( String type, String val )
	{
		if( type.equalsIgnoreCase( "Title" ) )
		{
			for( XLSRecord aChartArr : chartArr )
			{
				BiffRec b = aChartArr;
				if( b.getOpcode() == TEXTDISP )
				{    // should be the title
					TextDisp td = (TextDisp) b;
					if( td.isChartTitle() )
					{        // should!!
						td.setFontId( Integer.parseInt( val ) );
						break;
					}
				}
			}
		}
		else if( type.indexOf( "Default" ) > -1 )
		{
			type = type.substring( 7 );
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = chartArr.get( i );
				if( b.getOpcode() == DEFAULTTEXT )
				{
					if( ((DefaultText) b).getType() == Integer.parseInt( type ) )
					{
						i++;
						b = chartArr.get( i );
						if( b.getOpcode() == TEXTDISP )
						{    // should be!!
							TextDisp td = (TextDisp) b;
							td.setFontId( Integer.parseInt( val ) );
							break;
						}
					}
				}
			}
		}
		setMetricsDirty();
	}

	/**
	 * set the fontId for the chart title rec
	 *
	 * @param fontId
	 */
	public void setTitleFont( int fontId )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b.getOpcode() == TEXTDISP )
			{
				TextDisp td = (TextDisp) b;
				if( td.isChartTitle() )
				{
					td.setFontId( fontId );
				}
			}
		}
		setMetricsDirty();
	}

	// 20070802 KSC: debugging utility to write out chart recs
	public void writeChartRecs( String fName )
	{
		try
		{
			java.io.File f = new java.io.File( fName );
			BufferedWriter writer = new BufferedWriter( new FileWriter( f, true ) );
			int ctr = 0;
			if( preRecs != null )
			{
				ctr = ByteStreamer.writeRecs( new ArrayList( preRecs ), writer, ctr, 0 );
			}
			ctr = ByteStreamer.writeRecs( chartArr, writer, ctr, 0 );    // will recurse
			writer.flush();
			writer.close();
		}
		catch( Exception e )
		{
		}
	}

	/**
	 * set the sheet for this chart plus its subrecords as well
	 */
	@Override
	public void setSheet( Sheet b )
	{
		super.setSheet( b );
		if( this.msodrawobj != null )
		{
			this.msodrawobj.setSheet( b );
		}
		for( XLSRecord aChartArr : chartArr )
		{
			aChartArr.setSheet( b );
		}
	}

	/**
	 * return int[4] representing the coordinates (L, T, W, H) of this chart in pixels
	 *
	 * @return int[4] bounds
	 */
	public short[] getCoords()
	{
		short[] coords = { 20, 20, 100, 100 };
		if( this.msodrawobj != null )
		{
			coords = this.msodrawobj.getCoords();
		}
		else
		{
			log.error( "Chart missing Msodrawing record. Chart has no coordinates." );
		}
		if( !metricsDirty )
		{    // use adjusted values for w and h
			coords[2] = chartMetrics.get( "canvasw" ).shortValue();
			coords[3] = chartMetrics.get( "canvash" ).shortValue();
		}
		return coords;
	}

	/**
	 * return the coordinates of the outer plot area (plot + axes + labels) in pixels
	 *
	 * @return
	 */
	public float[] getPlotAreaCoords( float w, float h )
	{
		return this.chartaxes.getPlotAreaCoords( w, h );
	}

	/**
	 * set the coordinates of ths chart in pixels
	 *
	 * @param coords int[4]: L, T, W, H
	 */
	public void setCoords( short[] coords )
	{
		if( this.msodrawobj != null )
		{
			msodrawobj.setCoords( coords );
			setMetricsDirty();    // size is an element in minmax ...
		}
		else    // should NEVER happen
		{
			log.error( "Chart missing coordinates." );
		}

	}

	/**
	 * return bounds relative to row/cols including their respective offsets
	 *
	 * @return
	 */
	public short[] getBounds()
	{
		if( this.msodrawobj != null )
		{
			return msodrawobj.getBounds();
		}
		return null;
	}

	/**
	 * returns the offset within the column in pixels
	 *
	 * @return
	 */
	public short getColOffset()
	{
		if( this.msodrawobj != null )
		{
			return msodrawobj.getColOffset();
		}
		return 0;
	}

	/**
	 * set the bounds relative to row/col and their offsets into them
	 *
	 * @param bounds short[]
	 */
	public void setBounds( short[] bounds )
	{
		if( this.msodrawobj != null )
		{
			msodrawobj.setBounds( bounds );
		}
		setMetricsDirty();
	}

	/**
	 * return the top row of the chart
	 */
	public int getRow0()
	{
		if( this.msodrawobj != null )
		{
			return msodrawobj.getRow0();
		}
		return -1;
	}

	/**
	 * set the top row for this chart
	 *
	 * @param r
	 */
	public void setRow( int r )
	{
		if( this.msodrawobj != null )
		{
			msodrawobj.setRow( r );
		}
		setMetricsDirty();
	}

	/**
	 * return the left col of the chart
	 */
	public int getCol0()
	{
		if( this.msodrawobj != null )
		{
			return msodrawobj.getCol();
		}
		return -1;
	}

	/**
	 * return the height of this chart
	 */
	public int getHeight()
	{
		if( this.msodrawobj != null )
		{
			return msodrawobj.getHeight();
		}
		return -1;
	}

	/**
	 * set the height of this chart
	 *
	 * @param h
	 */
	public void setHeight( int h )
	{
		if( this.msodrawobj != null )
		{
			msodrawobj.setHeight( h );
		}
		setMetricsDirty();
	}

	/**
	 * retrieve the current Series JSON for comparisons
	 *
	 * @return JSONArray
	 */
	public JSONArray getSeriesJSON()
	{
		return chartseries.getSeriesJSON();
	}

	/**
	 * set the current Series JSON
	 *
	 * @param s JSONArray
	 */
	public void setSeriesJSON( JSONArray s ) throws JSONException
	{
		chartseries.setSeriesJSON( s );
		dirtyflag = true;
	}

	/**
	 * Show or remove Data Table for Chart
	 * NOTE:  METHOD IS STILL EXPERIMENTAL
	 *
	 * @param bShow
	 */
	public void showDataTable( boolean bShow )
	{
// TODO: FINISH		chartobj[0].showDataTable(bShow);
	}

	/**
	 * show or hide chart legend key
	 *
	 * @param bShow    boolean show or hide
	 * @param vertical boolean show as vertical or horizontal
	 */
	public void showLegend( boolean bShow, boolean vertical )
	{
		chartgroup.get( 0 ).showLegend( bShow, vertical );
	}

	/**
	 * remove the legend from the chart
	 */
	public void removeLegend()
	{
		showLegend( false, false );
	}

	/**
	 * return the Data Legend for this chart, if any
	 *
	 * @return
	 */
	public Legend getLegend()
	{
		return chartgroup.get( 0 ).getDataLegend();
	}

	/**
	 * return the default data label setting for the chart, if any
	 * <br>NOTE: each series can override the default data label fo the chart
	 * <br>can be one or more of:
	 * <br>VALUELABEL= 0x1;
	 * <br>VALUEPERCENT= 0x2;
	 * <br>CATEGORYPERCENT= 0x4;
	 * <br>CATEGORYLABEL= 0x10;
	 * <br>BUBBLELABEL= 0x20;
	 * <br>SERIESLABEL= 0x40;
	 *
	 * @return int default data label for chart
	 */
	protected int getDataLabel()
	{
		return chartgroup.get( 0 ).getDataLabel();

	}

	/**
	 * return data label options for each series as an int array
	 * <br>each can be one or more of:
	 * <br>VALUELABEL= 0x1;
	 * <br>VALUEPERCENT= 0x2;
	 * <br>CATEGORYPERCENT= 0x4;
	 * <br>CATEGORYLABEL= 0x10;
	 * <br>BUBBLELABEL= 0x20;
	 * <br>SERIESLABEL= 0x40;
	 *
	 * @return int array
	 * @see AttachedLabel
	 */
	protected int[] getDataLabelsPerSeries( int defaultDL )
	{
			/* NOTES:		 * 
		 * A data label is a label on a chart that is associated with a data point, or associated with a series  on an area or filled radar chart group. 
		 * A data label contains information about the associated data point, such as the description of the data point, a legend key, or custom text.
		 * 
		 * Inheritance
		* For any given data point, there is an order of inheritance that determines the contents of a data label associated with the data point:
			Data labels can be specified for a chart group, specifying the default setting for the data labels associated with the data points on the chart group .
			Data labels can be specified for a series, specifying the default setting for the data labels associated with the data points of the series. 
				This type of data label overrides the data label properties specified on the chart group for the data labels associated with the data points in a given series.
			Data labels can be specified for a data point, specifying the settings for a data label associated with a particular data point.
				This type of data label overrides the data label properties specified on the chart group and series for the data labels associated with a given data point.

		 * If formatting is not specified for an individual data point, the data point inherits the formatting of the series. 
		 * If formatting is not specified for the series, the series inherits the formatting of the chart group that contains the series. 
		 * The yi field of the DataFormat record MUST specify the zero-based index of the Series record associated with this series in the 
		 * collection of all Series records in the current chart sheet substream that contains the series. 
	 	 */
		return chartseries.getDataLabelsPerSeries( defaultDL, this.getChartType() );
	}

	/**
	 * return an array of the type of markers for each series:
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
	protected int[] getMarkerFormats( int nChart )
	{
		int[] mf = chartseries.getMarkerFormats();
		for( int marker : mf )
		{
			if( marker != 0 )
			{
				return mf;
			}
		}
		// see if chart format 
		return chartgroup.get( nChart ).getMarkerFormats();
	}

	/**
	 * returns true if this chart has data markers (line, scatter and radar charts only)
	 *
	 * @return
	 */
	public boolean hasMarkers( int nChart )
	{
		int[] markers = this.getMarkerFormats( nChart );
		for( int marker : markers )
		{
			if( marker != 0 )
			{
				return true;
			}
		}
		return false;

	}

	/**
	 * return truth if Chart has a data legend key showing
	 *
	 * @return
	 */
	public boolean hasDataLegend()
	{
		return (chartgroup.get( 0 ).getDataLegend() != null);
	}

	public MSODrawing getMsodrawobj()
	{
		return msodrawobj;
	}

	public Obj getObj()
	{
		return obj;
	}

	/**
	 * sets bar colors to vary or not
	 *
	 * @param vary
	 * @param nChart
	 */
	public void setVaryColor( boolean vary, int nChart )
	{
		this.getChartObject( nChart ).cf.setVaryColor( vary );
	}

	/**
	 * generic method to find a specific record in the list of recs in chartarr
	 *
	 * @param chartArr
	 * @param c        class of record to find
	 * @return biffrec or null
	 */
	public static BiffRec findRec( ArrayList chartArr, Class c )
	{
		for( Object aChartArr : chartArr )
		{
			BiffRec b = (BiffRec) aChartArr;
			if( b.getClass() == c )
			{
				return b;
			}
		}
		return null;
	}

	/**
	 * generic method to find a specific record in the list of recs in chartArr
	 *
	 * @param c class of record to find
	 * @return position of record
	 */
	public static int findRecPosition( ArrayList chartArr, Class c )
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getClass() == c )
			{
				return i;
			}
		}
		return -1;
	}

	// TODO: LATER
	// **** INCLUDE cell selection in series bars ...
	// chart.getSeries vs chart.getSeries(int) -- rename!!!
	// chart axes need a dirty flag to rebuild metrics
	// Axis.getSVG ==> passed categories -- should be passed chartseries instead??
	// merge minmaxcache into chartMetrics
	// metricsDirty - ensure all ops are covered ...
	// clean up Axes.getSVG w.r.t. label font, etc.
	public HashMap getMetrics( WorkBookHandle wbh )
	{
		if( metricsDirty )
		{
			try
			{
				this.wbh = wbh;
				//chartseries.setWorkBook(wbh);
				double[] minmax = chartseries.getMetrics( metricsDirty );    // Ignore Overlay charts for now!
				short[] coords = this.getCoords();
				chartMetrics.put( "x", (double) coords[0] );
				chartMetrics.put( "y", (double) coords[1] );
				chartMetrics.put( "w", (double) coords[2] );
				chartMetrics.put( "h", (double) coords[3] );
				chartMetrics.put( "canvasw", (double) coords[2] );
				chartMetrics.put( "canvash", (double) coords[3] );
				chartMetrics.put( "min", minmax[0] );
				chartMetrics.put( "max", minmax[1] );
				float[] plotcoords = null;
				plotcoords = this.getPlotAreaCoords( chartMetrics.get( "w" ).floatValue(), chartMetrics.get( "h" ).floatValue() );
				if( plotcoords == null )
				{
					CrtLayout12A crt = (CrtLayout12A) Chart.findRec( this.chartArr, CrtLayout12A.class );
					if( crt != null )
					{
						plotcoords = crt.getInnerPlotCoords( chartMetrics.get( "w" ).floatValue(), chartMetrics.get( "h" ).floatValue() );
					}
				}

				chartMetrics.put( "x", (double) plotcoords[0] );
				chartMetrics.put( "y", (double) plotcoords[1] );
				chartMetrics.put( "w", (double) plotcoords[2] );
				chartMetrics.put( "h", (double) plotcoords[3] );
				// Chart title offset
				com.extentech.formats.XLS.Font titlefont = this.getTitleFont();
				if( (titlefont != null) && !this.getTitle().equals( "" ) )
				{ // apparently can still have td even when no title is present ...
					float[] tdcoords = this.charttitle.getCoords();
					double fh = titlefont.getFontHeightInPoints();
					if( tdcoords[1] == 0 )
					{
						chartMetrics.put( "TITLEOFFSET", Math.ceil( fh * 1.5 ) );    // with padding
					}
					else
					{
						chartMetrics.put( "TITLEOFFSET", fh );    // a little padding
					}
				}
				else if( chartMetrics.get( "y" ) < 5.0 )
				{
					chartMetrics.put( "TITLEOFFSET", 10.0 );    // no title offset - a little padding
				}
				else
				{
					chartMetrics.put( "TITLEOFFSET", 0.0 );    // no title offset and no need for padding
				}
				this.getAxes().getMetrics( this.getChartType(), chartMetrics, plotcoords, this.getChartSeries().getCategories() );
				int[] lcoords = null;
				double adjust = 10;
				if( this.getLegend() != null )
				{
					this.getLegend().getMetrics( chartMetrics, this.getChartType(), this.getChartSeries() );
					lcoords = this.getLegend().getCoords();
					if( lcoords != null )
					// TODO: legend adjustment may have to do with y title and label ofsets ...?			
					{
						adjust = 2 * lcoords[4];    // spacing before and after legend box  TODO this isn't correct !!
//KSC: TESTING!		
//System.out.println("Original lcoords:  " + Arrays.toString(lcoords));					
					}
					else
					{
						lcoords = new int[6];
						lcoords[0] = chartMetrics.get( "canvasw" ).intValue();
					}
				}
				else
				{
					lcoords = new int[6];
					lcoords[0] = chartMetrics.get( "canvasw" ).intValue();
				}
				double ldist = lcoords[0] - chartMetrics.get( "w" );    // save distance between legend box and w (significant if legend is on rhs)
//System.out.println("Before Adjustments:  x:" + chartMetrics.get("x") + " w:" + chartMetrics.get("w") + " cw:" + chartMetrics.get("canvasw") + " y:" + chartMetrics.get("y") + " h:" + chartMetrics.get("h") + " ch:" + chartMetrics.get("canvash"));				
				// now adjust plot area coordinates based on canvas w, h, title and label offsets, and legend box, if any
				chartMetrics.put( "x",
				                  chartMetrics.get( "x" ) + (Double) this.getAxes().axisMetrics.get( "YAXISLABELOFFSET" ) + (Double) this
						                  .getAxes().axisMetrics.get( "YAXISTITLEOFFSET" ) );
				chartMetrics.put( "y", chartMetrics.get( "y" ) + chartMetrics.get( "TITLEOFFSET" ) );
				// TODO: seems that w is different doesn't need decrementing by x?? check out ...				
				chartMetrics.put( "w", chartMetrics.get( "w" ) - (Double) this.getAxes().axisMetrics.get( "YAXISLABELOFFSET" ) );
				chartMetrics.put( "h",
				                  chartMetrics.get( "canvash" ) - chartMetrics.get( "y" ) - (Double) this.getAxes().axisMetrics
						                  .get( "XAXISLABELOFFSET" ) - (Double) this.getAxes().axisMetrics.get( "XAXISTITLEOFFSET" ) - 10 );
//System.out.println("After Adjustments:   x:" + chartMetrics.get("x") + " w:" + chartMetrics.get("w") + " cw:" + chartMetrics.get("canvasw") + " y:" + chartMetrics.get("y") + " h:" + chartMetrics.get("h") + " ch:" + chartMetrics.get("canvash"));				

				double cw = chartMetrics.get( "canvasw" );
				// rhs legend has to have some extra adjustments to w and/or canvasw ... 
				if( lcoords[5] == Legend.RIGHT )
				{
					double legendBeg = lcoords[0] - (chartMetrics.get( "w" ) + chartMetrics.get( "x" ));
					double legendEnd = cw - (lcoords[0] + lcoords[2] + adjust);

					if( (legendBeg < 0) || (legendEnd < 0) )
					{ // try to adjust
						if( legendEnd < 0 )
						{
							chartMetrics.put( "canvasw", (lcoords[0] + lcoords[2] + adjust) );
							cw = (lcoords[0] + lcoords[2] + adjust);
						}
//						if (legendBeg < 0)
//							chartMetrics.put("w",  lcoords[0]-10.0-chartMetrics.get("x"));
					}
					if( this.getAxes().hasAxis( XAXIS ) && (ldist > 0) )    // pie, donut, don't
					// ensure distance between legend box and edge of plot area remains the same
					{
						chartMetrics.put( "w", lcoords[0] - chartMetrics.get( "x" ) - ldist );
					}
//System.out.println("Adjusted LCoords:  " + Arrays.toString(lcoords)); 				

				}
				else
				{
					double w = chartMetrics.get( "w" ) + chartMetrics.get( "x" );
					if( w > cw )
					{
						chartMetrics.put( "w", chartMetrics.get( "canvasw" ) - chartMetrics.get( "x" ) - 10 );
					}
				}

				metricsDirty = false;
			}
			catch( Exception e )
			{
				log.error( "Chart.getMetrics: " + e.toString() );
			}
		}
		return chartMetrics;
	}
}

