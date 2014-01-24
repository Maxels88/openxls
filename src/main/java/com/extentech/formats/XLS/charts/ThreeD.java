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

import com.extentech.formats.OOXML.OOXMLElement;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * <b>ThreeD(3D) Chart group is a 3-D Chart Group (0x103A)</b>
 * <p/>
 * anRot 2 		Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for others  -- def = 0
 * anElev 2 	Elevation Angle (-90 to 90 degrees) (15 is default) 8
 * pcDist 2		Distance from eye to chart (0 to 100) (30 is default) 10
 * pcHeight 2 	Height of plot volume relative to width and depth (100 is default) 12
 * pcDepth 2 	Depth of points relative to width (100 is default) 14
 * pcGap 2 		Space between points  (150 is default - should be 50!!!) 16 grbit 2
 * <p/>
 * grbit
 * 0 0x1 fPerspective 1= use perspective transform
 * 1 0x2 fCluster 1= 3-D columns are clustered or stacked
 * 2 0x4 f3DScaling 1= use auto-scaling
 * 3 reserved
 * 4 0x8 fNotPieChart 1= NOT a pie chart
 * 5 0x10 f2DWalls use 2D walls and gridlines (if fPerspective MUST be ignored. if not of type BAR, AREA or
 * PIE, ignore. if BAR and fCluster=0, ignore. specifies whether the walls are rendered in 2-D.
 * If fPerspective is 1 then this MUST be ignored. If the chart group type is not bar, area or pie this MUST be ignored. If the chart
 * group is of type bar and fCluster is 0, then this MUST be ignored. if PIE MUST be 0.
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 */
public class ThreeD extends GenericChartObject implements ChartObject
{
	private static final Logger log = LoggerFactory.getLogger( ThreeD.class );
	private static final long serialVersionUID = -7501630910970731901L;
	private short anRot = 0;
	private short anElev = 15;
	private short pcDist = 30;
	private short pcHeight = 100;
	private short pcDepth = 100;
	private short pcGap = 150;
	private short grbit = 0; //
	private boolean fPerspective;
	private boolean fCluster;
	private boolean f3dScaling;
	private boolean f2DWalls; // 20070905 KSC: parse grbit

	@Override
	public void init()
	{
		super.init();
		anRot = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		anElev = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		pcDist = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		pcHeight = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		pcDepth = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );
		pcGap = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
		grbit = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );
		fPerspective = (grbit & 0x1) == 0x1;
		fCluster = (grbit & 0x2) == 0x2;
		f3dScaling = (grbit & 0x4) == 0x4;
		f2DWalls = (grbit & 0x10) == 0x10;
	}

	// 20070716 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		ThreeD td = new ThreeD();
		td.setOpcode( THREED );
		td.setData( td.PROTOTYPE_BYTES );
		td.init();
		return td;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 0, 0, 30, 0, 100, 0, 100, 0, -106, 0, 0, 0
	};

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( anRot );
		getData()[0] = b[0];
		getData()[1] = b[1];
		b = ByteTools.shortToLEBytes( anElev );
		getData()[2] = b[0];
		getData()[3] = b[1];
		b = ByteTools.shortToLEBytes( pcDist );
		getData()[4] = b[0];
		getData()[5] = b[1];
		b = ByteTools.shortToLEBytes( pcHeight );
		getData()[6] = b[0];
		getData()[7] = b[1];
		b = ByteTools.shortToLEBytes( pcDepth );
		getData()[8] = b[0];
		getData()[9] = b[1];
		b = ByteTools.shortToLEBytes( pcGap );
		getData()[10] = b[0];
		getData()[11] = b[1];
		b = ByteTools.shortToLEBytes( grbit );
		getData()[12] = b[0];
		getData()[13] = b[1];
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "AnRot" ) )
		{ // specifies the clockwise rotation, in degrees, of the 3-D plot area around a vertical line through the center of the 3-D plot area.
			anRot = Short.parseShort( val ); // usually 20
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "AnElev" ) )
		{ // signed integer that specifies the rotation, in degrees, of the 3-D plot area around a horizontal line through the center of the 3-D plot area
			anElev = Short.parseShort( val ); // usually 15
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "PcDist" ) )
		{ // view angle for the 3-D plot area.
			pcDist = Short.parseShort( val ); // usually 30
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "PcHeight" ) )
		{ // specifies the height of the 3-D plot area as a percentage of its width
			pcHeight = Short.parseShort( val );
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "PcDepth" ) )
		{ // specifies the depth of the 3-D plot area as a percentage of its width.
			pcDepth = Short.parseShort( val ); // usually 100
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "PcGap" ) )
		{ // specifies the width of the gap between the series and the front and back edges of the 3-D plot area as a percentage of the data point depth divided by 2. If
			// fCluster is not set to 1 and chart group type is not a bar, this field also specifies the distance between adjacent series as a percentage of the data point depth.
			pcGap = Short.parseShort( val ); // usually 150
			bHandled = true;
		}
		/*
		 * 20070905 KSC: parse grbit options if
		 * (op.equalsIgnoreCase("FormatOptions")) { grbit=
		 * Short.parseShort(val); bHandled= true; }
		 */
		if( op.equalsIgnoreCase( "Perspective" ) )
		{ // specifies whether the 3-D plot area is rendered with a vanishing point.
			fPerspective = Boolean.valueOf( val );
			grbit = ByteTools.updateGrBit( grbit, fPerspective, 0 );
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "Cluster" ) )
		{ // specifies whether data points are clustered together in a bar chart group
			fCluster = Boolean.valueOf( val );
			grbit = ByteTools.updateGrBit( grbit, fCluster, 1 );
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "ThreeDScaling" ) )
		{ // specifies whether the height of the 3-D plot area is automatically determined
			f3dScaling = Boolean.valueOf( val );
			grbit = ByteTools.updateGrBit( grbit, f3dScaling, 2 );
			bHandled = true;
		}
		if( op.equalsIgnoreCase( "TwoDWalls" ) )
		{ // A bit that specifies whether the chart walls are rendered in 2-D
			f2DWalls = Boolean.valueOf( val );
			grbit = ByteTools.updateGrBit( grbit, f2DWalls, 4 );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

	/**
	 * return the desired option setting in string form
	 */
	@Override
	public String getChartOption( String op )
	{
		if( op.equalsIgnoreCase( "AnRot" ) )
		{
			return String.valueOf( anRot );
		}
		if( op.equalsIgnoreCase( "AnElev" ) )
		{
			return String.valueOf( anElev );
		}
		if( op.equalsIgnoreCase( "PcDist" ) )
		{
			return String.valueOf( pcDist );
		}
		if( op.equalsIgnoreCase( "PcHeight" ) )
		{
			return String.valueOf( pcHeight );
		}
		if( op.equalsIgnoreCase( "PcDepth" ) )
		{
			return String.valueOf( pcDepth );
		}
		if( op.equalsIgnoreCase( "PcGap" ) )
		{
			return String.valueOf( pcGap );
		}
		if( op.equalsIgnoreCase( "Perspective" ) )
		{
			return ((fPerspective) ? "1" : "0");
		}
		if( op.equalsIgnoreCase( "Cluster" ) )
		{
			return ((fCluster) ? "1" : "0");
		}
		if( op.equalsIgnoreCase( "ThreeDScaling" ) )
		{
			return ((f3dScaling) ? "1" : "0");
		}
		if( op.equalsIgnoreCase( "TwoDWalls" ) )
		{
			return ((f2DWalls) ? "1" : "0");
		}
		return "";
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( anRot != 0 )
		{
			sb.append( " AnRot=\"" + anRot + "\"" );
		}
		if( anElev != 15 )
		{
			sb.append( " AnElev=\"" + anElev + "\"" );
		}
		if( pcDist != 30 )
		{
			sb.append( " pcDist=\"" + pcDist + "\"" );
		}
		if( pcHeight != 100 )
		{
			sb.append( " pcHeight=\"" + pcHeight + "\"" );
		}
		if( pcDepth != 100 )
		{
			sb.append( " pcDepth=\"" + pcDepth + "\"" );
		}
		if( pcGap != 150 )
		{
			sb.append( " pcGap=\"" + pcGap + "\"" );
		}
		if( fPerspective )
		{
			sb.append( " Perspective=\"true\"" );
		}
		if( f3dScaling )
		{
			sb.append( " ThreeDScaling=\"true\"" );
		}
		if( f2DWalls )
		{
			sb.append( " TwoDWalls=\"true\"" );
		}
		// 20070913 KSC: need to track Cluster, whether on or off
		sb.append( " Cluster=\"" + fCluster + "\"" );
		// sb.append(" formatOptions=\"" + grbit+ "\"");
		return sb.toString();
	}

	/**
	 * @return truth of "Chart is Clustered"
	 */
	public boolean isClustered()
	{
		return fCluster;
	}

	/**
	 * sets if this chart has clustered bar/columns
	 *
	 * @param bIsClustered
	 */
	public void setIsClustered( boolean bIsClustered )
	{
		fCluster = bIsClustered;
		grbit = ByteTools.updateGrBit( grbit, fCluster, 1 );
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[12] = b[0];
		getData()[13] = b[1];
	}

	/**
	 * sets the proper bit for chart type (PIE or not)
	 *
	 * @param isPieChart true if pie-type chart
	 */
	public void setIsPie( boolean isPieChart )
	{
		if( isPieChart )
		{
			grbit &= 0x8; // bit is true if "not a pie"
		}
		else
		{
			grbit |= 0x17;
		}
		updateRecord();
	}

	/**
	 * sets the Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for
	 * others
	 *
	 * @param rot
	 */
	public void setAnRot( int rot )
	{
		anRot = (short) rot;
		byte[] b = ByteTools.shortToLEBytes( anRot );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * sets the Elevation Angle (-90 to 90 degrees) (15 is default)
	 *
	 * @param elev
	 */
	public void setAnElev( int elev )
	{
		anElev = (short) elev;
		byte[] b = ByteTools.shortToLEBytes( anElev );
		getData()[2] = b[0];
		getData()[3] = b[1];
	}

	/**
	 * sets the Distance from eye to chart (0 to 100) (30 is default)
	 *
	 * @param elev
	 */
	public void setPcDist( int dist )
	{
		pcDist = (short) dist;
		byte[] b = ByteTools.shortToLEBytes( pcDist );
		getData()[4] = b[0];
		getData()[5] = b[1];
	}

	/**
	 * sets the Height of plot volume relative to width and depth (100 is
	 * default)
	 *
	 * @param elev
	 */
	public void setPcHeight( int dist )
	{
		pcHeight = (short) dist;
		byte[] b = ByteTools.shortToLEBytes( pcHeight );
		getData()[6] = b[0];
		getData()[7] = b[1];
	}

	/**
	 * sets the Depth of points relative to width (100 is default)
	 *
	 * @param elev
	 */
	public void setPcDepth( int depth )
	{
		pcDepth = (short) depth;
		byte[] b = ByteTools.shortToLEBytes( pcDepth );
		getData()[8] = b[0];
		getData()[9] = b[1];
	}

	/**
	 * sets the Space between points (50 or 150 is default)
	 *
	 * @param gap
	 */
	public void setPcGap( int gap )
	{
		pcGap = (short) gap;
		byte[] b = ByteTools.shortToLEBytes( pcGap );
		getData()[10] = b[0];
		getData()[11] = b[1];
	}

	public int getPcGap()
	{
		return pcGap;
	}

	/**
	 * return view3d OOXML representation
	 *
	 * @return
	 */
	public StringBuffer getOOXML()
	{
		StringBuffer cooxml = new StringBuffer();
		cooxml.append( "<c:view3D>" );
		cooxml.append( "\r\n" );
		// rotX == anElev
		if( anElev != 0 ) // default
		{
			cooxml.append( "<c:rotX val=\"" + anElev + "\"/>" );
		}
		// hPercent -- a height percent between 5 and 500.
		// rotY == anRot
		if( (anRot != 0) || (anElev != 0) ) // default
		{
			cooxml.append( "<c:rotY val=\"" + anRot + "\"/>" );
		}
		// depthPercentage -- This element specifies the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent).
		if( pcDepth != 100 )
		{
			cooxml.append( "<c:depthPercent val=\"" + pcDepth + "\"/>" );
		}
		// rAngAx == !fPerspective
		if( fPerspective )
		{
			cooxml.append( "<c:rAngAx val=\"1\"/>" );
		}
		// perspective == pcDist
		if( pcDist != 30 ) // default
		{
			cooxml.append( "<c:perspective val=\"" + pcDist + "\"/>" );
		}
		cooxml.append( "</c:view3D>" );
		cooxml.append( "\r\n" );
		return cooxml;

	}

	/**
	 * parse shape OOXML element view3D into a ThreeD record
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spPr object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, OOXMLChart cht )
	{
		// threeD MUST NOT EXIST in a bar of pie, bubble, doughnut, filled
		// radar, pie of pie, radar, or scatter chart group.
		ThreeD td = cht.getChartObject().getThreeDRec( true );
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					String v = null;
					try
					{
						v = xpp.getAttributeValue( 0 );
					}
					catch(/* XmlPullParser */Exception e )
					{
					}
					if( v != null )
					{
						if( tnm.equals( "rotX" ) )
						{
							td.setAnElev( Integer.valueOf( v ) );
						}
						else if( tnm.equals( "rotY" ) )
						{
							td.setAnRot( Integer.valueOf( v ) );
						}
						else if( tnm.equals( "perspective" ) )
						{
							td.setPcDist( Integer.valueOf( v ) );
						}
						else if( tnm.equals( "depthPercent" ) )
						{
							td.setPcDepth( Integer.valueOf( v ) );
						}
						else if( tnm.equals( "rAngAx" ) )
						{
//						if (v!=null)

						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "view3D" ) )
					{
						lastTag.pop(); // pop layout tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "ThreeD.parseOOXML: " + e.toString() );
		}

		return null;
	}
}
