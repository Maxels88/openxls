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
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * txBody (Shape Text Body)
 * <p/>
 * OOXML/DrawingML element specifies the existence of text to be contained within the corresponding shape. All visible text and
 * visible text related properties are contained within this element. There can be multiple paragraphs and within
 * paragraphs multiple runs of text
 * <p/>
 * parent:    sp (shape)
 * children:  bodyPr REQ, lstStyle, p REQ
 */
// TODO: handle lstStyle Text List Styles
public class TxBody implements OOXMLElement
{

	private static final long serialVersionUID = 2407194628070113668L;
	private BodyPr bPr;
	private P para;

	public TxBody( BodyPr b, P para )
	{
		this.bPr = b;
		this.para = para;
	}

	public TxBody( TxBody tbd )
	{
		this.bPr = tbd.bPr;
		this.para = tbd.para;
	}

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
					if( endTag.equals( "txBody" ) )
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
			Logger.logErr( "txBody.parseOOXML: " + e.toString() );
		}
		TxBody tBd = new TxBody( b, para );
		return tBd;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<xdr:txBody>" );
		if( bPr != null )
		{
			tooxml.append( bPr.getOOXML() );
		}
		tooxml.append( "<a:lstStyle/>" ); // TODO: HANDLE
		if( para != null )
		{
			tooxml.append( para.getOOXML() );
		}
		tooxml.append( "</xdr:txBody>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new TxBody( this );
	}
}
