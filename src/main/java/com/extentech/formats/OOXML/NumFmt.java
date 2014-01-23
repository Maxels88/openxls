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

import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.formats.XLS.Xf;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

/**
 * numFmt  OOXML element
 * <p/>
 * numFmt (Number Format)
 * This element specifies number format properties which indicate how to format and render the numeric value of
 * a cell.
 * Following is a listing of number formats whose formatCode value is implied rather than explicitly saved in
 * the file. In this case a numFmtId value is written on the xf record, but no corresponding numFmt element is
 * written. Some of these Ids are interpreted differently, depending on the UI language of the implementing
 * application.
 * <p/>
 * parent:  styleSheet/numFmts element in styles.xml
 * NOTE:  numFmt element also occurs in drawingML, numFmt is replaced with sourceLinked
 * children: none
 */
public class NumFmt implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( NumFmt.class );
	private static final long serialVersionUID = -206715418106414662L;
	private String formatCode, numFmtId;
	private boolean sourceLinked = false;

	public NumFmt( String formatCode, String numFmtId, boolean sourceLinked )
	{
		this.formatCode = formatCode;
		this.numFmtId = numFmtId;
		this.sourceLinked = sourceLinked;
	}

	public NumFmt( NumFmt n )
	{
		this.formatCode = n.formatCode;
		this.numFmtId = n.numFmtId;
		this.sourceLinked = n.sourceLinked;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp )
	{
		String formatCode = null, numFmtId = null;
		boolean sourceLinked = false;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "numFmt" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( n.equals( "formatCode" ) )
							{
								formatCode = v;
							}
							else if( n.equals( "numFmtId" ) )
							{
								numFmtId = v;
							}
							else if( n.equals( "sourceLinked" ) )
							{
								sourceLinked = (v.equals( "1" ));
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "numFmt" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "numFmt.parseOOXML: " + e.toString() );
		}
		NumFmt oe = new NumFmt( formatCode, numFmtId, sourceLinked );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		return getOOXML( "" );
	}

	public String getOOXML( String ns )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + ns + "numFmt" );
		// attributes
		ooxml.append( " formatCode=\"" + OOXMLAdapter.stripNonAscii( formatCode ) + "\"" );
		if( numFmtId != null )
		{
			ooxml.append( " numFmtId=\"" + numFmtId + "\"" );
		}
		if( sourceLinked )
		{
			ooxml.append( " sourceLinked=\"1\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NumFmt( this );
	}

	/**
	 * returns the format id assoc with this number format
	 *
	 * @return
	 */
	public String getFormatId()
	{
		return numFmtId;
	}

	/**
	 * returns the OOXML specifying the fill based on this FormatHandle object
	 */
	public static String getOOXML( Xf xf )
	{
		// Number Format		  
		StringBuffer ooxml = new StringBuffer();
		if( xf.getIfmt() > FormatConstants.BUILTIN_FORMATS.length )
		{    // only input user defined formats ...
			String s = xf.getFormatPattern();
			if( s != null )
			{
				s = StringTool.replaceText( s,
				                            "\"",
				                            "&quot;" ); // replace internal quotes   // 1.6 only s= s.replace('"', "&quot;"); // replace internal quotes
			}
			ooxml.append( "<numFmt numFmtId=\"" + xf.getIfmt() + "\" formatCode=\"" + s + "\"/>" );
			ooxml.append( "\r\n" );
		}
		return ooxml.toString();
	}
}

