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
 * dLbls (Data Labels)
 * This element serves as a root element that specifies the settings for the data labels for an entire series or the
 * entire chart. It contains child elements that specify the specific formatting and positioning settings.
 * <p/>
 * parent:  chart types, ser
 * children: dLbl (0 or more times), GROUPDLBLS (numFmt, spPr, txPr, dLblPos, showLegendKey, showVal, showCatName, showSerName,showPercent, showBubbleSize, separator, showLeaderLines
 */
// TODO: Finish All Children!!!! leaderLines, numFmt, separator 
public class DLbls implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( DLbls.class );
	private static final long serialVersionUID = -3765320144606034211L;
	private SpPr sp = null;
	private TxPr tx = null;
	// all of these are child elements with one attribute=val, all default to "1" (true) 
	private int showVal = -1;
	private int showLeaderLines = -1;
	private int showLegendKey = -1;
	private int showCatName = -1;
	private int showSerName = -1;
	private int showPercent = -1;
	private int showBubbleSize = -1;

	public DLbls( int showVal,
	              int showLeaderLines,
	              int showLegendKey,
	              int showCatName,
	              int showSerName,
	              int showPercent,
	              int showBubbleSize,
	              SpPr sp,
	              TxPr tx )
	{
		this.showVal = showVal;
		this.showLeaderLines = showLeaderLines;
		this.showLegendKey = showLegendKey;
		this.showCatName = showCatName;
		this.showSerName = showSerName;
		this.showPercent = showPercent;
		this.showBubbleSize = showBubbleSize;
		this.sp = sp;
		this.tx = tx;
	}

	public DLbls( boolean showVal,
	              boolean showLeaderLines,
	              boolean showLegendKey,
	              boolean showCatName,
	              boolean showSerName,
	              boolean showPercent,
	              boolean showBubbleSize,
	              SpPr sp,
	              TxPr tx )
	{
		if( showVal )
		{
			this.showVal = 1;
		}
		if( showLeaderLines )
		{
			this.showLeaderLines = 1;
		}
		if( showLegendKey )
		{
			this.showLegendKey = 1;
		}
		if( showCatName )
		{
			this.showCatName = 1;
		}
		if( showSerName )
		{
			this.showSerName = 1;
		}
		if( showPercent )
		{
			this.showPercent = 1;
		}
		if( showBubbleSize )
		{
			this.showBubbleSize = 1;
		}
		this.sp = sp;
		this.tx = tx;
	}

	public DLbls( DLbls d )
	{
		showVal = d.showVal;
		showLeaderLines = d.showLeaderLines;
		showLegendKey = d.showLegendKey;
		showCatName = d.showCatName;
		showSerName = d.showSerName;
		showPercent = d.showPercent;
		showBubbleSize = d.showBubbleSize;
		sp = d.sp;
		tx = d.tx;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		SpPr sp = null;
		TxPr tx = null;
		int showVal = -1;
		int showLeaderLines = -1;
		int showLegendKey = -1;
		int showCatName = -1;
		int showSerName = -1;
		int showPercent = -1;
		int showBubbleSize = -1;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
//	        			 sp.setNS("c");
					}
					else if( tnm.equals( "txPr" ) )
					{
						lastTag.push( tnm );
						tx = (TxPr) TxPr.parseOOXML( xpp, lastTag, bk ).cloneElement();
					}
					else if( tnm.equals( "showVal" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showVal = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showLeaderLines" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showLeaderLines = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showLegendKey" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showLegendKey = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showCatName" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showCatName = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showSerName" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showSerName = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showPercent" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showPercent = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
					else if( tnm.equals( "showBubbleSize" ) )
					{
						if( xpp.getAttributeCount() > 0 )
						{
							showBubbleSize = Integer.valueOf( xpp.getAttributeValue( 0 ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "dLbls" ) )
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
			log.error( "dLbls.parseOOXML: " + e.toString() );
		}
		DLbls d = new DLbls( showVal, showLeaderLines, showLegendKey, showCatName, showSerName, showPercent, showBubbleSize, sp, tx );
		return d;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<c:dLbls>" );
		// TODO: numFmt
		if( sp != null )
		{
			ooxml.append( sp.getOOXML() );
		}
		if( tx != null )
		{
			ooxml.append( tx.getOOXML() );
		}
		// TODO: dLblPos
		if( showLegendKey != -1 )
		{
			ooxml.append( "<c:showLegendKey val=\"" + showLegendKey + "\"/>" );
		}
		if( showVal != -1 )
		{
			ooxml.append( "<c:showVal val=\"" + showVal + "\"/>" );
		}
		if( showCatName != -1 )
		{
			ooxml.append( "<c:showCatName val=\"" + showCatName + "\"/>" );
		}
		if( showSerName != -1 )
		{
			ooxml.append( "<c:showSerName val=\"" + showSerName + "\"/>" );
		}
		if( showPercent != -1 )
		{
			ooxml.append( "<c:showPercent val=\"" + showPercent + "\"/>" );
		}
		if( showBubbleSize != -1 )
		{
			ooxml.append( "<c:showBubbleSize val=\"" + showBubbleSize + "\"/>" );
		}
		// TODO: separator
		if( showLeaderLines != -1 )
		{
			ooxml.append( "<c:showLeaderLines val=\"" + showLeaderLines + "\"/>" );
		}
		// TODO: leaderLines
		ooxml.append( "</c:dLbls>" );
		return ooxml.toString();
	}

	/**
	 * generate the ooxml necessary to define the data labels for a chart
	 * Controls view of Series name, Category Name, Percents, Leader Lines, Bubble Sizes where applicable
	 *
	 * @return public String getOOXML(ChartFormat cf) {
	 * StringBuffer ooxml= new StringBuffer();
	 * ooxml.append("<c:dLbls>"); ooxml.append("\r\n");
	 * // TODO: c:numFmt, c:spPr, c:txPr
	 * if (cf.getChartOption("ShowBubbleSizes")=="1")
	 * ooxml.append("<c:showBubbleSize val=\"1\"/>");
	 * if (cf.getChartOption("ShowValueLabel")=="1")
	 * ooxml.append("<c:showVal val=\"1\"/>");
	 * if (cf.getChartOption("ShowLabel")=="1")
	 * ooxml.append("<c:showSerName val=\"1\"/>");
	 * if (cf.getChartOption("ShowCatLabel")=="1")
	 * ooxml.append("<c:showCatName val=\"1\"/>");
	 * // Pie specific
	 * if (cf.getChartOption("ShowLabelPct")=="1")
	 * ooxml.append("<c:showPercent val=\"1\"/>");
	 * if (cf.getChartOption("ShowLdrLines")=="true")
	 * ooxml.append("<c:showLeaderLines val=\"1\"/>");
	 * <p/>
	 * <p/>
	 * ooxml.append("</c:dLbls>"); ooxml.append("\r\n");
	 * return ooxml.toString();
	 * }
	 */

	@Override
	public OOXMLElement cloneElement()
	{
		return new DLbls( this );
	}

	/**
	 * get methods
	 */
	public boolean showLegendKey()
	{
		return showLegendKey == 1;
	}

	public boolean showVal()
	{
		return showVal == 1;
	}

	public boolean showCatName()
	{
		return showCatName == 1;
	}

	public boolean showSerName()
	{
		return showSerName == 1;
	}

	public boolean showPercent()
	{
		return showPercent == 1;
	}

	public boolean showBubbleSize()
	{
		return showBubbleSize == 1;
	}

	public boolean showLeaderLines()
	{
		return showLeaderLines == 1;
	}
}



