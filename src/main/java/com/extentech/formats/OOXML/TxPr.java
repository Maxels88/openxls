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

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.FormatConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * txpr (Text Properties)
 * <p/>
 * OOXML/DrawingML element specifies text formatting. The lstStyle element is not supported
 * <p/>
 * parent:  axes, title, labels ..
 * children: bodyPr, lstStyle, p
 */
// TODO: Handle lstStyle
public class TxPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( TxPr.class );
	private static final long serialVersionUID = -4293247897525807479L;
	// TODO: handle lstStyle Text List Styles
	private BodyPr bPr;
	private P para;

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		BodyPr b = null;
		P para = null;
		try
		{        // need: endParaRPr?
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "bodyPr" ) )
					{        // default text properties
						lastTag.push( tnm );
						b = (BodyPr) BodyPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "p" ) )
					{    // part of p element
						lastTag.push( tnm );
						para = (P) P.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "txPr" ) )
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
			log.error( "txPr.parseOOXML: " + e.toString() );
		}
		TxPr tPr = new TxPr( b, para );
		return tPr;
	}

	public TxPr( BodyPr b, P para )
	{
		bPr = b;
		this.para = para;
	}

	public TxPr( TxPr tpr )
	{
		bPr = tpr.bPr;
		para = tpr.para;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<c:txPr>" );
		if( bPr != null )
		{
			tooxml.append( bPr.getOOXML() );
		}
		tooxml.append( "<a:lstStyle/>" ); // TODO: HANDLE
		if( para != null )
		{
			tooxml.append( para.getOOXML() );
		}
		tooxml.append( "</c:txPr>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new TxPr( this );
	}

	/**
	 * specify text formatting properties
	 *
	 * @param fontFace font face
	 * @param sz       size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
	 * @param b        true if bold
	 * @param i        true if italic
	 * @param u        underline:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
	 *                 dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
	 * @param strike   one of: dblStrike, noStrike or sngStrike  or null if none
	 * @param clr      fill color in hex form without the #
	 */
	public TxPr( String fontFace, int sz, boolean b, boolean i, String u, String strike, String clr )
	{
		para = new P( fontFace, sz, b, i, u, strike, clr );
	}

	public TxPr( Font fx, int hrot, String vrot )
	{
			/*
			 * Specifies the rotation that is being applied to the text within the bounding box. If it not
specified, the rotation of the accompanying shape is used. If it is specified, then this is
applied independently from the shape. That is the shape can have a rotation applied in
addition to the text itself having a rotation applied to it. If this attribute is omitted, then a
value of 0, is implied.

represents an angle in 60,000ths of a degree
5400000=90 degrees
-5400000=90 degrees counter-clockwise
			 */
		int u = fx.getUnderlineStyle();
		String usty = "none";
		switch( u )
		{
			case FormatConstants.STYLE_UNDERLINE_SINGLE:
				usty = "sng";
				break;
			case FormatConstants.STYLE_UNDERLINE_DOUBLE:
				usty = "dbl";
				break;
			case FormatConstants.STYLE_UNDERLINE_SINGLE_ACCTG:
				usty = "sng";
				break;
			case FormatConstants.STYLE_UNDERLINE_DOUBLE_ACCTG:
				usty = "dbl";
				break;
		}
		String strike = (fx.getStricken() ? "sngStrike" : "noStrike");
		bPr = new BodyPr( hrot, "horz" );
		para = new P( fx.getFontName(),
		                   (int) fx.getFontHeightInPoints() * 100,
		                   fx.getBold(),
		                   fx.getItalic(),
		                   usty,
		                   strike,
		                   fx.getColorAsOOXMLRBG() );
/*
 * <c:txPr><a:bodyPr rot="-5400000" vert="horz"/>
 * <a:lstStyle/>
 * <a:p><a:pPr><a:defRPr sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0">
 * <a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial Narrow"/><a:ea typeface="Arial Narrow"/><a:cs typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>			
 */
	}
}
