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
package org.openxls.formats.OOXML;

import org.openxls.ExtenXLS.WorkBookHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * dPt (Data Point)
 * This element specifies a single data point.
 * <p/>
 * parent: series
 * children: idx REQ, invertIfNegative, marker, bubble3D, explosion, spPr, pictureOptions
 */
// TODO: finish pictureOptions
public class DPt implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( DPt.class );
	private static final long serialVersionUID = 8354707071603571747L;
	private int idx;
	private boolean invertIfNegative;
	private boolean bubble3D;
	private Marker marker;
	private SpPr spPr;
	private int explosion;

	public DPt( int idx, boolean invertIfNegative, boolean bubble3D, Marker m, SpPr sp, int explosion )
	{
		this.idx = idx;
		this.invertIfNegative = invertIfNegative;
		this.bubble3D = bubble3D;
		marker = m;
		spPr = sp;
		this.explosion = explosion;
	}

	public DPt( DPt d )
	{
		idx = d.idx;
		invertIfNegative = d.invertIfNegative;
		bubble3D = d.bubble3D;
		marker = d.marker;
		spPr = d.spPr;
		explosion = d.explosion;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		int idx = -1;
		boolean invertIfNegative = true;
		boolean bubble3D = true;
		Marker m = null;
		SpPr sp = null;
		int explosion = 0;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "idx" ) )
					{    // child element only contains 1 element
						if( xpp.getAttributeCount() > 0 )
						{
							idx = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "invertIfNegative" ) )
					{    // child element only contains 1 element
						if( xpp.getAttributeCount() > 0 )
						{
							invertIfNegative = (xpp.getAttributeValue( 0 ).equals( "1" ));
						}
					}
					else if( tnm.equals( "bubble3D" ) )
					{    // child element only contains 1 element
						if( xpp.getAttributeCount() > 0 )
						{
							bubble3D = (xpp.getAttributeValue( 0 ).equals( "1" ));
						}
					}
					else if( tnm.equals( "explosion" ) )
					{    // child element only contains 1 element
						if( xpp.getAttributeCount() > 0 )
						{
							explosion = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
//		            	 sp.setNS("c");
					}
					else if( tnm.equals( "marker" ) )
					{
						lastTag.push( tnm );
						m = (Marker) Marker.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "dPt" ) )
					{
						lastTag.pop();
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "dPt.parseOOXML: " + e.toString() );
		}
		DPt oe = new DPt( idx, invertIfNegative, bubble3D, m, sp, explosion );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:dPt>" );
		ooxml.append( "<c:idx val=\"" + idx + "\"/>" );
		if( !invertIfNegative )
		{
			ooxml.append( "<c:invertIfNegative val=\"0\"/>" ); // default= true
		}
		if( marker != null )
		{
			ooxml.append( marker.getOOXML() );
		}
		if( !bubble3D )
		{
			ooxml.append( "<c:bubble3D val=\"0\"/>" ); // default= true
		}
		if( explosion != 0 )
		{
			ooxml.append( "<c:explosion val=\"" + explosion + "\"/>" );
		}
		if( spPr != null )
		{
			ooxml.append( spPr.getOOXML() );
		}
		ooxml.append( "</c:dPt>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new DPt( this );
	}
}


/*
 * generates the OOXML necessary to represent the data points of the chart 
 * includes fill and line color
 * @return
private String getDataPointOOXML(ChartSeriesHandle[] sh) {
	StringBuffer ooxml= new StringBuffer();
	for (int i= 0; i < sh.length; i++) {
		ooxml.append("<c:dPt>");				ooxml.append("\r\n");
		ooxml.append("<c:idx val=\"" + i + "\"/>");
		if (sh[i].getSpPr()!=null) ooxml.append(sh[i].getSpPr().getOOXML("c"));
		ooxml.append("</c:dPt>");				ooxml.append("\r\n");
	}
	return ooxml.toString();
}
*/