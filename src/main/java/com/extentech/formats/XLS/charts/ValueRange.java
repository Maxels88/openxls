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
 * <b>ValueRange: Defines Value Axis Scale (0x101f)</b>
 * Offset		Name		Size	Contents
 * 4			numMin		8		Minimum value on axis. MUST be less than the value of numMax. If the value of fAutoMin is 1, this field MUST be ignored.
 * 12			numMax		8		Maximum value on axis. MUST be greater than the value of numMin. If the value of fAutoMax is 1, this field MUST be ignored.
 * 20			numMajor	8		Value of major increment. MUST be greater than or equal to the value of numMinor. If the value of fAutoMajor is 1, this field MUST be ignored.
 * 28			numMinor	8		Value of minor increment. MUST be greater than or equal to zero. If the value of fAutoMinor is 1, this field MUST be ignored.
 * 36			numCross	8		Value where category axis crosses. If the value of fAutoCross is 1, this field MUST be ignored.
 * 44			grbit		2		Format flags
 * <p/>
 * <p/>
 * grbit
 * 0		0x1		fAutoMin		Automatic Minimum Selected
 * 0	The value specified by numMin is used as the minimum value of the value axis.
 * 1	numMin is calculated such that the data point with the minimum value can be displayed in the plot area.
 * 1		0x2		fAutoMax		Automatic Maximum Selected
 * 0	The value specified by numMax is used as the maximum value of the value axis.
 * 1	numMax is calculated such that the data point with the maximum value can be displayed in the plot area.
 * 2		0x4		fAutoMajor		Automatic Major Unit Selected
 * 0	The value specified by numMax is used as the maximum value of the value axis.
 * 1	numMax is calculated such that the data point with the maximum value can be displayed in the plot area.
 * 3		0x8		fAutoMinor		Automatic Minor Unit Selected
 * 0	The value specified by numMinor is used as the interval at which minor tick marks and minor gridlines are displayed.
 * 1	numMinor is calculated automatically.
 * 4		0x10	fAutoCross		Automatic Category Crossing Point Selected
 * 0	The value specified by numCross is used as the point at which the other axes in the axis group cross this value axis.
 * 1	numCross is calculated so that the crossing point is displayed in the plot area.
 * 5		0x20	fLogScale		Log Scale
 * 0	The scale of the value axis is linear.
 * 1	The scale of the value axis is logarithmic. The default base of the logarithmic scale is 10, unless a CrtMlFrt record follows this record, specifying the base in a XmlTkLogBaseFrt structure.
 * 6		0x40	fReverse		Values in reverse order
 * 0	Values are displayed from smallest-to-largest, from left-to-right, or from bottom-to-top, respectively, depending on the orientation of the axis.
 * 1	The values are displayed in reverse order, meaning largest-to-smallest, from left-to-right, or from bottom-to-top, respectively.
 * 7		0x80	fMaxCross		Category is to cross at maximum value
 * 0	The other axes in the axis group cross this value axis at the value specified by numCross.
 * 1	The other axes in the axis group cross the value axis at the maximum value. If fMaxCross is 1, then both fAutoCross and numCross MUST be ignored.
 * <p/>
 * All 8-byte numbers in the preceding table are IEEE floating-point numbers.
 * The numMin field defines the minimum numeric value that appears along the value axis.  This field is all zeros if Auto Minimum is selected on the Scale tab of the Format Axis dialog box.
 * The numMax field defines the maximum value displayed along the value axis and is all zeros if Auto Maximum is selected.
 * The numMajor field defines the increment (unit) of the major value divisions (gridlines) along the value axis.  The numMajor field is all zeros if Auto Major Unit is selected on the Scale tab of the Format Axis dialog box.
 * The numMinor field defines the minor value divisions (gridlines) along the value axis and is all zeros if Auto Minor Unit is selected.
 * The numCross field defines the value along the value axis at which the category axis crosses. This field is all zeros if Auto Category Axis Crosses At is selected.
 */
public class ValueRange extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2989883115978826628L;
	double numMin, numMax, numMajor, numMinor, numCross;
	//	double yMin= 0.0, yMax= 0.0;
	short grbit = 0;
	boolean fAutoMin, fAutoMax, fAutoMajor, fAutoMinor, fAutoCross, fLogScale, fReverse, fMaxCross;

	public void init()
	{
		super.init();
		numMin = ByteTools.eightBytetoLEDouble( this.getBytesAt( 0, 8 ) );
		numMax = ByteTools.eightBytetoLEDouble( this.getBytesAt( 8, 16 ) );
		numMajor = ByteTools.eightBytetoLEDouble( this.getBytesAt( 16, 24 ) );
		numMinor = ByteTools.eightBytetoLEDouble( this.getBytesAt( 24, 32 ) );
		numCross = ByteTools.eightBytetoLEDouble( this.getBytesAt( 32, 40 ) );
		grbit = ByteTools.readShort( this.getByteAt( 40 ), this.getByteAt( 41 ) );
		fMaxCross = (grbit & 0x80) == 0x80;
		fAutoMin = (grbit & 0x1) == 0x1;
		fAutoMax = (grbit & 0x2) == 0x2;
		fAutoMajor = (grbit & 0x4) == 0x4;
		fAutoMinor = (grbit & 0x8) == 0x8;
		fAutoCross = (grbit & 0x10) == 0x10;
		fLogScale = (grbit & 0x20) == 0x20;
		fReverse = (grbit & 0x40) == 0x40;
	}

	// 20070723 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		ValueRange vr = new ValueRange();
		vr.setOpcode( VALUERANGE );
		vr.setData( vr.PROTOTYPE_BYTES );
		vr.init();    // important when we parse options
		return vr;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 31, 1
	};

	/**        Excel automatic calculations:
	 * 		Given:
	 *    	yMax   The maximum y value used in your chart.
	 yMin   The minimum y value used in your chart.
	 xMax   The maximum x value used in your chart. This applies only to
	 charts that use x values, such as scatter and bubble charts.
	 xMin   The minimum x value used in your chart. This applies only to
	 charts that use x values, such as scatter and bubble charts.

	 When you create a chart in Microsoft Excel, there are three possible scenarios that may apply to your data:

	 The yMax and yMin values are both non-negative (greater than or equal to zero). This is Scenario One.
	 The yMax and yMin values are both non-positive (less than or equal to zero). This is Scenario Two.
	 The yMax value is positive, and the yMin value is negative. This is Scenario Three.
	 ************************************************************************************************************************************
	 The major unit used by the y-axis is automatically determined by Microsoft Excel, based on all of the data included in the chart.
	 ************************************************************************************************************************************
	 The following scenarios use this default major unit.
	 * @return
	 */

	/**
	 * When you create a chart in Microsoft Excel,
	 * there are three possible scenarios that can apply to your data:

	 *  Scenario one: the yMax and yMin values are both positive or equal to zero.
	 *  Scenario two: the yMax and yMin values are both negative or equal to zero.
	 *  Scenario three: the yMax value is positive, and the yMin value is negative.

	 The major unit used by the y-axis is automatically determined by Microsoft Excel,
	 based on all of the data included in the chart.

	 The following scenarios use this default major unit.
	 Scenario one:
	 * If the chart is a 2-D area, column, bar, line or x-y scatter chart, the automatic maximum for the y-axis is the first major unit greater than or equal to the value returned by the following equation:
	 yMax + 0.05 * ( yMax - yMin )
	 Otherwise, the automatic maximum for the y-axis is the first major unit greater than or equal to yMax.
	 * If the difference between yMax and yMin is greater than 16.667 percent of the value of yMax, the automatic minimum for the y-axis is zero.
	 * If the difference between yMax and yMin is less than 16.667 percent of the value of yMax, the automatic minimum for the y-axis is the first major unit less than or equal to the value returned by the following equation:
	 yMin - ( ( yMax - yMin ) / 2 )
	 Exception: If the chart is an x-y scatter or bubble chart, the automatic minimum for the y-axis is the first major unit less than or equal to yMin.

	 Scenario two:
	 same as above except that
	 yMin + 0.05 * ( yMin - yMax )
	 Otherwise, the automatic minimum for the y-axis is the first major unit less than or equal to yMin.
	 and
	 yMax - ( ( yMin - yMax ) / 2 )

	 Scenario 3
	 * The automatic maximum for the y-axis is the first major unit greater than or equal to the value returned by the following equation:
	 yMax + 0.05 * ( yMax - yMin )
	 * The automatic minimum for the y-axis is the first major unit less than or equal to the value returned by the following equation:
	 yMin + 0.05 * ( yMin - yMax )

	 The above information also applies to charts that use x values,
	 such as x-y scatter charts and bubble charts.
	 For these types of charts,
	 substitute xMax and xMin for yMax and yMin in the above scenarios.
	 */

	/**
	 * return minumum scale value -- see setMinMax for calc
	 *
	 * @return double
	 */
	public double getMin()
	{
		return numMin;
	}

	/**
	 * return max scale value -- see setMinMax for calc
	 *
	 * @return double
	 */
	public double getMax()
	{
		return numMax;
	}

	/**
	 * all important major tick step -
	 *
	 * @return double major tick step value
	 * @see setMinMax
	 */
	public double getMajorTick()
	{
		return numMajor;
	}

	public double getMinorTick()
	{
		return numMinor;
	}    // fautominor --woundn't even know how to calculate!

	/**
	 * sets the max and min values for the axis and uses these
	 * to calculate the all-important major unit (and minor unit) step
	 * <br>
	 * NOTE that Excel keeps it's "automatic" calculation private so
	 * below is the best guess looking at existing chart data
	 * <br>max and min units are also based upon chart size, so this is only a rough approximation
	 *
	 * @param yMax double maximum series value
	 * @param yMin double minimum series value
	 */
	public void setMaxMin( double yMax, double yMin )
	{
		/**
		 * When you create a chart in Microsoft Excel, there are three possible scenarios that can apply to your data:
		 Scenario one: the yMax and yMin values are both positive or equal to zero.
		 Scenario two: the yMax and yMin values are both negative or equal to zero.
		 Scenario three: the yMax value is positive, and the yMin value is negative.

		 NOTE: the difficultly in these calculations is the major unit, of which the maximum scale is based.
		 There is no true information regarding the major unit, unfortunately, except that it is based on
		 "all the data of the chart"
		 */
		if( fAutoMajor && fAutoMinor )
		{
			int charttype = this.getParentChart().getChartType();
			if( yMax >= 0 && yMin >= 0 && yMax != yMin )
			{
				// Major Unit Calculation -- best guest TODO: would be great to find out Excel's algorithm!
// TODO: major unit is affected by height (or width, for bar charts) ***** develop algorithm!!!		   		
				// add a tiny pad for range ...
				double diff = (yMax * 1.1 - yMin);
				if( fAutoMax )
				{
					double logDiff = Math.floor( Math.log10( diff ) );
					double f = (diff) / Math.pow( 10, logDiff );
					if( f <= 1 )
					{
						f = 1;
					}
					else if( f <= 2 )
					{
						f = 2;
					}
					else if( f <= 5 )
					{
						f = 5;
					}
					else
					{
						f = 10;
					}
					f = f * Math.pow( 10, logDiff );                //scaled up max
					numMajor = f * .1;    // 1/10th of scaled up max

				}
				else
				{
					numMajor = numMax / 10.0;
				}
				/**
				 * If the chart is a 2-D area, column, bar, line or x-y scatter chart,
				 * the automatic maximum for the y-axis is the first major unit greater
				 * than or equal to the value returned by the following equation:
				 yMax + 0.05 * ( yMax - yMin )
				 Otherwise, the automatic maximum for the y-axis is the first major unit
				 greater than or equal to yMax.
				 */
				if( fAutoMax )
				{
					if( charttype == ChartConstants.AREACHART ||
							charttype == ChartConstants.COLCHART ||
							charttype == ChartConstants.BARCHART ||
							charttype == ChartConstants.LINECHART ||
							charttype == ChartConstants.SCATTERCHART ||
							charttype == ChartConstants.BUBBLECHART )
					{
						if( numMajor == (int) numMajor ) // int scale - usual case
						{
							numMax = Math.ceil( yMax + 0.05 * diff * 1.1 );
						}
						else
						{
							numMax = yMax + 0.05 * diff * 1.1;
						}
						if( charttype == ChartConstants.BUBBLECHART )
						{
							numMax += numMajor;    // is this true in ALL CASES????
						}
						if( (numMax % numMajor) != 0 )    // if not = to scale, scale up to next major unit
						{
							numMax = Math.floor( (numMax + numMajor) / numMajor ) * numMajor;
						}
					}
					else
					{
						numMax = Math.floor( (yMax + numMajor) / numMajor ) * numMajor;
					}
				}
				/**
				 * If the difference between yMax and yMin is greater than 16.667 percent of the value of yMax, the automatic minimum for the y-axis is zero.
				 */
				if( fAutoMin )
				{
					/**
					 * Exception: If the chart is an x-y scatter or bubble chart, the automatic minimum for the y-axis is the first major unit less than or equal to yMin.
					 */
					if( charttype == ChartConstants.SCATTERCHART || charttype == ChartConstants.BUBBLECHART )
					{
						if( yMin % numMajor != 0 )
						{
							numMin = Math.floor( (yMin - numMajor) / numMajor ) * numMajor;
							numMin = Math.round( numMin );
						}
					}
					else
					{
						if( (yMax - yMin) > (numMax * .16667) )
						{
							numMin = 0;
						}
						else
						{    // the first major unit less than or equal to the value from the below equation:
							numMin = yMin - ((numMax - yMin) / 2);
							if( (numMin % numMajor) != 0 )
							{
								numMin = Math.floor( (numMin - numMajor) / numMajor ) * numMajor;
							}
						}
					}
				}

				// 20120905 KSC: recheck major to ensure not more than 10 steps ...
				if( numMin >= 0 && numMax >= 0 )
				{
					if( ((numMax - numMin) / numMajor) > 9 )
					{
						diff = (numMax * 1.1 - numMin);
						double logDiff = Math.floor( Math.log10( diff ) );
						double f = (diff) / Math.pow( 10, logDiff );
						if( f <= 1 )
						{
							f = 1;
						}
						else if( f <= 2 )
						{
							f = 2;
						}
						else if( f <= 5 )
						{
							f = 5;
						}
						else
						{
							f = 10;
						}
						f = f * Math.pow( 10, logDiff );                //scaled up max
						numMajor = f * .1;    // 1/10th of scaled up max
					}
				}
				numMinor = numMajor / 5;    // seems to be the correct calculation ...
			}
			else if( yMax < 0 && yMin < 0 )
			{
				double diff = (yMin - yMax);
				numMin = yMin + 0.05 * diff;
				if( diff > (yMin * .16667) )
				{
					numMax = 0;
				}
				else
				{
					numMax = yMax - (diff / 2);
				}
				numMax = Math.floor( (numMax + numMajor) / numMajor ) * numMajor;
				if( fAutoMinor )
				{
					numMinor = numMajor / 5;    // just a guess, really
					// Exception: If the chart is an x-y scatter or bubble chart, the automatic maximum for the y-axis is the first major unit greater than or equal to yMax.
				}
				else
				{ // yMax > 0 && yMin < 0
					numMax = yMax + 0.05 * (yMax - yMin);
					numMin = yMin + 0.05 * (yMin - yMax);
				}
			}
			else
			{
				numMax = yMax;
				numMin = yMin;
			}
		}
	}

	/**
	 * static utlity to calculate the min and max on a scale in a given area
	 * NOTE: this is usually the Y Axis, but scatter charts can have a value axis on the X Axis
	 *
	 * @param MaxVal Actual Maximum Value on Axis
	 * @param MinVal Actual Mimumum Value on Axis
	 * @param area   Area of Axis in pixels (I think :))
	 * @return
	 */
	public static double[] calcMaxMin( double MaxVal, double MinVal, double area )
	{
		// h==235 is normal, and the below alg. appear correct for it
		// h==776 (bar chart) ??? SIGH ... 
		double ymax = MaxVal;
		double ymin = MinVal;
		double numMajor = ymax - ymin;
		if( numMajor > 0 && numMajor < 20 )
		{
			numMajor = 2;
		}
		else if( numMajor > 20 && numMajor < 100 )
		{
			numMajor = 20;
		}
		else if( numMajor > 100 && numMajor < 500 )
		{
			numMajor = 50;
		}
		if( ymax >= 0 )
		{
			ymax = Math.floor( (ymax + numMajor) / numMajor ) * numMajor;
		}
		else
		{
			ymax = Math.floor( (ymax + numMajor) / numMajor ) * numMajor;
		}

		double numMin = ymin - numMajor;
		if( (numMin % numMajor) > 0 )
		{
			numMin -= numMajor;
		}
		if( ymin >= 0 )
		{
			numMin = Math.max( ymin - numMajor, 0 );
		}
		numMin = Math.round( numMin );

		double numMax = 0;
		if( ymax >= 0 )
		{
			numMax = (int) Math.floor( (ymax + numMajor) / numMajor ) * numMajor;
		}
		else
		{
			numMax = (int) Math.floor( (ymax + numMajor) / numMajor ) * numMajor;
		}

		return new double[]{ numMin, numMajor, ymax };
	}

	/**
	 * sets a specific OOXML axis option
	 * <br>can be one of:
	 * <br>crosses			possible crossing points (autoZero, max, min)
	 * <br>crossBeteween	whether axis crosses the cat. axis between or on categories (between, midCat)
	 * <br>crossesAt		where on axis the perpendicular axis crosses (double val)
	 * <br>majorTickMark	major tick mark position (cross, in, none, out)
	 * <br>minorTickMark	minor tick mark position ("")
	 * <br>tickLblPos		tick label position (high, low, nextTo, none)
	 * <br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
	 * <br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
	 *
	 * @param op
	 * @param val
	 */
	public boolean setOption( String op, String val )
	{
		if( op.equals( "crossesAt" ) )
		// specifies where axis crosses		  -- numCross or catCross
		{
			numCross = new Double( val ).doubleValue();
		}
		else if( op.equals( "orientation" ) )
		{    // axis orientation minMax or maxMin  -- fReverse
			fReverse = (val.equals( "maxMin" ));    // means in reverse order
			ByteTools.updateGrBit( grbit, fReverse, 6 );
		}
		else if( op.equals( "crosses" ) )
		{            // specifies how axis crosses it's perpendicular axis (val= max, min, autoZero)  -- fbetween + fMaxCross?/fAutoCross + fMaxCross
			if( val.equals( "max" ) )
			{    // TODO: this is probly wrong
				fMaxCross = true;
				ByteTools.updateGrBit( grbit, fMaxCross, 7 );
			}
			else if( val.equals( "autoZero" ) )
			{
				fAutoCross = true;    // is this correct??
				ByteTools.updateGrBit( grbit, fAutoCross, 4 );
			}
			else if( val.equals( "min" ) )
			{
				fAutoCross = false;
				ByteTools.updateGrBit( grbit, fAutoCross, 4 );
			}
		}
		else if( op.equals( "crossBetween" ) )
		{    // val= between, midCat, crossBetween
			if( val.equals( "between" ) )
			{
				fAutoCross = true;
			}
			// otherwise do what???
		}
		else if( op.equals( "max" ) )
		{            // axis max - valueRange only?
			numMax = new Double( val ).doubleValue();
			// turn off automatic scaling
			grbit = (short) (grbit & 0xFD);    // turn off bit 2
		}
		else if( op.equals( "min" ) )
		{            // axis min- valueRange only?
			numMin = new Double( val ).doubleValue();
			// turn off automatic scaling
			grbit = (short) (grbit & 0xFE);    // turn off bit 1
		}
		else if( op.equals( "majorUnit" ) )
		{
			numMajor = new Double( val ).doubleValue();
		}
		else if( op.equals( "minorUnit" ) )
		{
			numMinor = new Double( val ).doubleValue();
		}
		else
		{
			return false;
		}
		this.updateRecord();
		return true;
	}

	/**
	 * retrieve generic Value axis option
	 * <br>can be one of:
	 * <br>crosses			possible crossing points (autoZero, max, min)
	 * <br>crossBeteween	whether axis crosses the cat. axis between or on categories (between, midCat)
	 * <br>crossesAt		where on axis the perpendicular axis crosses (double val)
	 * <br>majorTickMark	major tick mark position (cross, in, none, out)
	 * <br>minorTickMark	minor tick mark position ("")
	 * <br>tickLblPos		tick label position (high, low, nextTo, none)
	 * <br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
	 * <br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
	 *
	 * @param op
	 * @return
	 */
	public String getOption( String op )
	{
		if( op.equals( "crossesAt" ) )
		{
			return String.valueOf( numCross );
		}
		if( op.equals( "orientation" ) )
		{
			return (fReverse) ? "maxMin" : "minMax";
		}
		if( op.equals( "crosses" ) )
		{
			if( fMaxCross )
			{
				return "max";
			}
			if( fAutoCross )
			{
				return "autoZero";    // correct??
			}
			return "min";    // correct??
		}
		if( op.equals( "crossBetween" ) )        // val= between, midCat, crossBetween
		{
			return "between";        // TODO: figure out!
		}
		if( op.equals( "max" ) )
		{
			return String.valueOf( numMax );
		}
		if( op.equals( "min" ) )
		{
			return String.valueOf( numMin );
		}
		if( op.equals( "majorUnit" ) )
		{
			return String.valueOf( numMajor );
		}
		if( op.equals( "minorUnit" ) )
		{
			return String.valueOf( numMinor );
		}
		return null;
	}

	/**
	 * set the minimum scale value of this Y or value axis
	 * <br>Doing so turns automatic minimum off
	 *
	 * @param min
	 */
	public void setMin( double min )
	{
		numMin = min;
		// turn off automatic scaling
		grbit = (short) (grbit & 0xFE);    // turn off bit 1
		updateRecord();
	}

	/**
	 * set the max scale value of this Y or value axis
	 * <br>Doing so turns automatic maximum off
	 *
	 * @param max
	 */
	public void setMax( double max )
	{
		numMax = max;
		// turn off automatic scaling
		grbit = (short) (grbit & 0xFD);    // turn off bit 2
		updateRecord();
	}

	public boolean isAutomaticScale()
	{
		return fAutoMin || fAutoMax || fAutoMinor || fAutoMajor;
	}

	public boolean isAutomaticMax()
	{
		return fAutoMax;
	}

	public boolean isAutomaticMin()
	{
		return fAutoMin;
	}

	public void setAutomaticMin( boolean b )
	{
		fAutoMin = b;
		if( b )
		{
			grbit = (short) (grbit | 0x1);    // turn on bit 1
			numMin = 0.0;
		}
		else
		{
			grbit = (short) (grbit & 0xFE);    // turn off bit 1
		}
		updateRecord();
	}

	public void setAutomaticMax( boolean b )
	{
		fAutoMax = b;
		if( b )
		{
			grbit = (short) (grbit | 0x2);    // turn on bit 2
			numMax = 0.0;
		}
		else
		{
			grbit = (short) (grbit & 0xFD);    // turn off bit 2
		}
		updateRecord();
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.doubleToLEByteArray( numMin );
		System.arraycopy( b, 0, this.getData(), 0, 8 );
		b = ByteTools.doubleToLEByteArray( numMax );
		System.arraycopy( b, 0, this.getData(), 8, 8 );
		b = ByteTools.doubleToLEByteArray( numMajor );
		System.arraycopy( b, 0, this.getData(), 16, 8 );
		b = ByteTools.doubleToLEByteArray( numMinor );
		System.arraycopy( b, 0, this.getData(), 24, 8 );
		b = ByteTools.doubleToLEByteArray( numCross );
		System.arraycopy( b, 0, this.getData(), 32, 8 );
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[40] = b[0];
		this.getData()[41] = b[1];
	}

	/**
	 * returns true if axis should be displayed on RHS of chart
	 * false for default LHS
	 *
	 * @return
	 */
	public boolean isReversed()
	{
		return fReverse;
	}
}
