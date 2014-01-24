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
import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.formats.XLS.Sst;
import com.extentech.formats.XLS.Unicodestring;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.Stack;

/**
 * text (Comment Text)
 * This element contains rich text which represents the text of a comment. The maximum length for this text is a
 * spreadsheet application implementation detail. A recommended guideline is 32767 chars
 * parent: comment
 * children:	t (text), r (Rich Text Run), rPh (phonetic run), phoneticPr (phonetic properties)
 */
// TODO: finish elements rPh and phoneticPr
// TODO: preserve
public class Text implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Text.class );
	private static final long serialVersionUID = 5886384020139606328L;
	private Unicodestring str = null;

	/**
	 * create a new comment WITH formatting
	 *
	 * @param strref
	 */
	public Text( Unicodestring str )
	{
		this.str = str;
	}

	public Text( Text t )
	{
		str = t.str;
	}

	/**
	 * create a new comment with NO formatting
	 *
	 * @param s
	 */
	public Text( String s )
	{
		str = Sst.createUnicodeString( s, null, Sst.STRING_ENCODING_AUTO );
	}

	/**
	 * parse this Text element into a unicode string with formatting runs
	 */
	public static Text parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		Unicodestring str = null;
		String s = "";
		ArrayList<short[]> formattingRuns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "rPr" ) )
					{    // intra-string formatting properties
						int idx = s.length();    // index into character string to apply formatting to
						Ss_rPr rp = (Ss_rPr) Ss_rPr.parseOOXML( xpp, bk );    //.cloneElement();
						Font f = rp.generateFont( bk );  // NOW CONVERT ss_rPr to a font!!
						int fIndex = bk.getWorkBook().getFontIdx( f );  // index for specific font formatting
						if( fIndex == -1 )  // must insert new font
						{
							fIndex = bk.getWorkBook().insertFont( f ) + 1;
						}
						if( formattingRuns == null )
						{
							formattingRuns = new ArrayList<>();
						}
						formattingRuns.add( new short[]{ Integer.valueOf( idx ).shortValue(), Integer.valueOf( fIndex ).shortValue() } );
					}
					else if( tnm.equals( "t" ) )
					{
				    	 /*boolean bPreserve= false;
		            	 if (xpp.getAttributeCount()>0) {
		            		 if (xpp.getAttributeName(0).equals("space") && xpp.getAttributeValue(0).equals("preserve"))
		            			 bPreserve= true;
		            	 } 
		            	 */
						eventType = xpp.next();
						while( (eventType != XmlPullParser.END_DOCUMENT) &&
								(eventType != XmlPullParser.END_TAG) &&
								(eventType != XmlPullParser.TEXT) )
						{
							eventType = xpp.next();
						}
						if( eventType == XmlPullParser.TEXT )
						{
							s = s + xpp.getText();
						}
					}
				}
				else if( (eventType == XmlPullParser.END_TAG) && xpp.getName().equals( "text" ) )
				{
					str = Sst.createUnicodeString( s,
					                               formattingRuns,
					                               Sst.STRING_ENCODING_UNICODE );    // create a new unicode string with formatting runs
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "r.parseOOXML: " + e.toString() );
		}
		Text oe = new Text( str );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		return null;
	}

	/**
	 * return the OOXML representation of this Text (Comment) element
	 *
	 * @param bk
	 * @return
	 */
	public String getOOXML( com.extentech.formats.XLS.WorkBook bk )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<text>" );
		if( str != null )
		{
			String s = OOXMLAdapter.stripNonAsciiRetainQuote( str.getStringVal() ).toString();
			ArrayList runs = str.getFormattingRuns();
			if( runs == null )
			{    // no intra-string formatting
				if( (s.indexOf( " " ) == 0) || (s.lastIndexOf( " " ) == (s.length() - 1)) )
				{
					ooxml.append( "<t xml:space=\"preserve\">" + s + "</t>" );
				}
				else
				{
					ooxml.append( "<t>" + s + "</t>" );
				}
				ooxml.append( "\r\n" );
			}
			else
			{    // have formatting runs which split up string into areas with separate formats applied
				/*
				 *  <element name="rPr" type="CT_RPrElt" minOccurs="0" maxOccurs="1"/>
					<element name="t" type="ST_Xstring" minOccurs="1" maxOccurs="1"/>
				 */
				int begIdx = 0;
				ooxml.append( "<r>" );    // new rich text run
				for( Object run : runs )
				{
					short[] idxs = (short[]) run;
					if( idxs[0] > begIdx )
					{
						ooxml.append( ("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii( s.substring( begIdx,
						                                                                                       idxs[0] ) ) + "</t>") );
						ooxml.append( "</r>" );
						ooxml.append( "\r\n" );
						ooxml.append( "<r>" );
						begIdx = idxs[0];
					}
					Ss_rPr rp = Ss_rPr.createFromFont( bk.getFont( idxs[1] ) );
					ooxml.append( rp.getOOXML() );
				}
				if( begIdx < s.length() )    // output remaining string
				{
					s = s.substring( begIdx );
				}
				else
				{
					s = "";
				}
				ooxml.append( "<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii( s ) + "</t>" );
				ooxml.append( "\r\n" );
				ooxml.append( "</r>" );
			}
		}
		ooxml.append( "</text>" );
		ooxml.append( "\r\n" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Text( this );
	}

	/**
	 * return the String value of this Text (Comment) element
	 * i.e. without formatting
	 *
	 * @return
	 */
	public String getComment()
	{
		if( str != null )
		{
			return str.getStringVal();
		}
		return null;
	}

	/**
	 * return the String value of this Text (Comment) element
	 * Including formatting runs
	 *
	 * @return
	 */
	public Unicodestring getCommentWithFormatting()
	{
		return str;
	}
}

