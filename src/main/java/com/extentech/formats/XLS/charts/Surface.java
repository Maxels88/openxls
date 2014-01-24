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
import org.json.JSONException;
import org.json.JSONObject;

/**
 * <b>Surface: Chart Group is a Surface Chart Group (0x103f) </b>
 * <p/>
 * 4		grbit		2
 * <p/>
 * 0		0x1		fFillSurface		1= chart contains color fill for surface
 * 1		0x2		f3DPhongShade		1= this surface chart has shading
 */
public class Surface extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -3243029185139320374L;
	private short grbit = 0;
	private boolean fFillSurface = true;
	private boolean f3DPhoneShade = false;
	private boolean is3d = false;	/* since all surface charts contain a 3d record, must store 3d setting separately */

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		fFillSurface = (grbit & 0x1) == 0x1;
		f3DPhoneShade = (grbit & 0x2) == 0x2;
		chartType = ChartConstants.SURFACECHART;
	}

	// 20070703 KSC:
	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		Surface b = new Surface();
		b.setOpcode( SURFACE );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	/**
	 * returns true if surface chart is wireframe, false if filled
	 *
	 * @return
	 */
	public boolean isWireframe()
	{
		return !fFillSurface;
	}

	/**
	 * sets this surface chart to wireframe (true) or filled(false)
	 */
	public void setIsWireframe( boolean wireframe )
	{
		fFillSurface = !wireframe;
		grbit = ByteTools.updateGrBit( grbit, fFillSurface, 0 );
	}

	/**
	 * surface charts always contain a 3d record so must determine if 3d separately
	 */
	public boolean getIs3d()
	{
		return is3d;
	}

	/**
	 * set if this surface chart is "truly" 3d as all surface-type charts contain a 3d record
	 *
	 * @param is3d
	 */
	public void setIs3d( boolean is3d )
	{
		this.is3d = is3d;
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( fFillSurface )
		{
			sb.append( " ColorFill=\"true\"" );
		}
		if( f3DPhoneShade )
		{
			sb.append( " Shading=\"true\"" );
		}
		return sb.toString();
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "ColorFill" ) )
		{
			fFillSurface = val.equals( "true" );
			grbit = ByteTools.updateGrBit( grbit, fFillSurface, 0 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shading" ) )
		{
			f3DPhoneShade = val.equals( "true" );
			grbit = ByteTools.updateGrBit( grbit, f3DPhoneShade, 1 );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

	/**
	 * return the (dojo) type JSON for this Chart Object
	 *
	 * @return
	 */
	public JSONObject getTypeJSON() throws JSONException
	{
		JSONObject typeJSON = new JSONObject();
		String dojoType;
		if( !isStacked() )
		{
			dojoType = "Areas";
		}
		else
		{
			dojoType = "StackedAreas";
		}
		typeJSON.put( "type", dojoType );
		return typeJSON;
	}
}
