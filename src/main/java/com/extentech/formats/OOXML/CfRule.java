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
package com.extentech.formats.OOXML;

import com.extentech.formats.XLS.OOXMLAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * cfRule (Conditional Formatting Rule)
 * This collection represents a description of a conditional formatting rule.
 * <br>
 * NOTE: now is used merely to parse the cfRule OOXML element
 * the data is stored in a BIFF-8 Cf object
 * <p/>
 * parent:       conditionalFormatting
 * children:     SEQ: formula (0-3), colorScale, dataBar, iconSet
 * attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
 * percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
 */
//TODO: Finish children colorScale, dataBar, iconSet
public class CfRule implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CfRule.class );
	private static final long serialVersionUID = 8509907308100079138L;
	private HashMap<String, String> attrs;
	private ArrayList<String> formulas;

	public CfRule( HashMap<String, String> attrs, ArrayList<String> formulas )
	{
		this.attrs = attrs;
		this.formulas = formulas;
	}

	public CfRule( CfRule cf )
	{
		attrs = cf.attrs;
		formulas = cf.formulas;
	}

	/**
	 * generate a cfRule based on OOXML input stream
	 *
	 * @param xpp
	 * @return
	 */
	// TODO: finish children colorScale, dataBar, iconSet
	public static CfRule parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<>();
		ArrayList<String> formulas = new ArrayList<>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cfRule" ) )
					{     // get attributes: priority, type, operator, dxfId
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "formula" ) )
					{
						formulas.add( OOXMLAdapter.getNextText( xpp ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cfRule" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "cfRule.parseOOXML: " + e.toString() );
		}
		CfRule cf = new CfRule( attrs, formulas );
		return cf;
	}

	/**
	 * generate OOXML for this cfRule
	 * NOW Cf is parsed to obtain OOXML and CfRule is not retained
	 *
	 * @return
	 * @deprecated
	 */
	public String getOOXML( int priority )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<cfRule" );
		// attributes
		// TODO: use passed in priority
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( formulas != null )
		{
			for( String formula : formulas )
			{
				ooxml.append( "<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote( formula ) + "</formula>" );
			}
		}
		// TODO: finish children dataBar, colorScale, iconSet
		ooxml.append( "</cfRule>" );
		return ooxml.toString();
	}

	/**
	 * get the dxfId (the incremental style associated with this conditional formatting rule)
	 *
	 * @return
	 */
	public int getDxfId()
	{
		if( attrs != null )
		{
			try
			{
				return Integer.valueOf( attrs.get( "dxfId" ) );
			}
			catch( Exception e )
			{
				// it's possible to not specify a dxfId
				;
			}
		}
		return -1;
	}

	/**
	 * set the dxfId (the incremental style associated with this conditional formatting rule)
	 *
	 * @param dxfId
	 * @deprecated
	 */
	public void setDxfId( int dxfId )
	{
		if( attrs == null )
		{
			attrs = new HashMap<>();
		}
		attrs.put( "dxfId", Integer.valueOf( dxfId ).toString() );
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CfRule( this );
	}

	@Override
	public String getOOXML()
	{
		// TODO Auto-generated method stub
		return null;
	}

	/**
	 * get methods
	 */
	public String getOperator()
	{
		// valid only when type=="cellIs"
		if( attrs != null )
		{
			return attrs.get( "operator" );
		}
		return null;
	}

	public String getType()
	{
		String type = null;
		if( attrs != null )
		{
			type = attrs.get( "type" );
		}
		if( type == null )
		{
			type = "cellIs";
		}
		return type;
	}

	/**
	 * returns the text to test in a containsText type of condition
	 * <br>Only valid for containsText conditions
	 *
	 * @return
	 */
	public String getContainsText()
	{
		if( attrs != null )
		{
			return attrs.get( "text" );
		}
		return null;
	}

	public String getFormula1()
	{
		if( (formulas != null) && (formulas.size() > 0) )
		{
			return formulas.get( 0 );
		}
		return null;
	}

	public String getFormula2()
	{
		if( (formulas != null) && (formulas.size() > 1) )
		{
			return formulas.get( 1 );
		}
		return null;
	}
}

/** dataBar (Data Bar)
 * Describes a data bar conditional formatting rule.
 * [Example:
 * In this example a data bar conditional format is expressed, which spreads across all cell values in the cell range,
 * and whose color is blue.
 * <dataBar>
 * <cfvo type="min" val="0"/>
 * <cfvo type="max" val="0"/>
 * <color rgb="FF638EC6"/>
 * </dataBar>
 * end example]
 * The length of the data bar for any cell can be calculated as follows:
 * Data bar length = minLength + (cell value - minimum value in the range) / 1 (maximum value in the range -
 * minimum value in the range) * (maxLength - minLength),
 * where min and max length are a fixed percentage of the column width (by default, 10% and 90% respectively.)
 * The minimum difference in length (or increment amount) is 1 pixel.
 *
 * parent:       cfRule
 * children:     SEQ: cfvo (2), color (1)
 * attributes:   minLength (def=10), maxLength (def=90), showValue (def=true)
 */

/**
 * colorScale (Color Scale)
 * Describes a graduated color scale in this conditional formatting rule.
 *
 * parent:      cfRule
 * children:    SEQ: cfvo (2+), color (2+)
 */

/**
 * cfvo (Conditional Format Value Object)
 * Describes the values of the interpolation points in a gradient scale.
 *
 * parents: colorScale, dataBar, iconSet
 * children: extLst
 * attributes:  type (REQ), val, gte (def=true)
 *
 */

