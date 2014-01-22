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

import com.extentech.ExtenXLS.DocumentHandle;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.Font;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;

/**
 * rPr (Run Properties for Shared Strings (see SharedStrings.xml))
 * <p/>
 * This element represents a set of properties to apply to the contents of this rich text run.
 * This element corresponds to Unicode String formatting runs where a specific font is applied to a sub-section of a string
 * <p/>
 * parent:  r (rich text run)
 * children:  many
 */
public class Ss_rPr implements OOXMLElement
{

	private static final long serialVersionUID = 8940630588129002652L;
	private HashMap<String, String> attrs = new HashMap<String, String>();
	private Color color = null;

	public Ss_rPr()
	{

	}

	public Ss_rPr( HashMap<String, String> attrs, Color c )
	{
		this.attrs = attrs;
		this.color = c;
	}

	public Ss_rPr( Ss_rPr r )
	{
		this.attrs = r.attrs;
		this.color = r.color;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		Color c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "rFont" ) )
					{
						attrs.put( "rFont", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "charset" ) )
					{
						attrs.put( "charset", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "family" ) )
					{
						attrs.put( "family", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "b" ) )
					{
						attrs.put( "b", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "i" ) )
					{
						attrs.put( "i", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "strike" ) )
					{
						attrs.put( "strike", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "outline" ) )
					{
						attrs.put( "outline", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "shadow" ) )
					{
						attrs.put( "shadow", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "condense" ) )
					{
						attrs.put( "condense", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "extend" ) )
					{
						attrs.put( "extend", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "sz" ) )
					{
						attrs.put( "sz", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "u" ) )
					{
						attrs.put( "u", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "vertAlign" ) )
					{
						attrs.put( "vertAlign", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "scheme" ) )
					{
						attrs.put( "scheme", ((xpp.getAttributeCount() > 0) ? xpp.getAttributeValue( 0 ) : "") );    // val
					}
					else if( tnm.equals( "color" ) )
					{
						c = (Color) Color.parseOOXML( xpp, FormatHandle.colorFONT, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "rPr" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "rPr.parseOOXML: " + e.toString() );
		}
		Ss_rPr r = new Ss_rPr( attrs, c );
		return r;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<rPr>" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			if( val.equals( "" ) ) // same as true for <b/> <u/> <i/> ...
			{
				ooxml.append( "<" + key + "/>" );
			}
			else
			{
				ooxml.append( "<" + key + " val=\"" + val + "\"/>" );
			}
		}
		if( color != null )
		{
			ooxml.append( color.getOOXML() );
		}
		ooxml.append( "</rPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Ss_rPr( this );
	}

	/**
	 * retrieve one of the following key values, if already set
	 * b, charset, condense, extend, family, i, outline, rFont, scheme, shadow, strike, sz, u, vertAlign
	 *
	 * @param key
	 * @return
	 */
	public String getAttr( String key )
	{
		return attrs.get( key );
	}

	/**
	 * set the value for one of the rPr children:
	 * b, charset, condense, extend, family, i, outline, rFont, scheme, shadow, strike, sz, u, vertAlign
	 * <p/>
	 * most are boolean string values "0" or "1"
	 * except for sz, rFont, family, scheme
	 *
	 * @param key
	 * @param val
	 */
	public void setAttr( String key, String val )
	{
		attrs.put( key, val );
	}

	/** family:  
	 0 Not applicable.
	 1 Roman
	 2 Swiss
	 SpreadsheetML Reference Material - Styles
	 2113
	 Value Font Family
	 3 Modern
	 4 Script
	 5 Decorative
	 */
	/**
	 * charset (Character Set)
	 * This element defines the font character set of this font.
	 * This field is used in font creation and selection if a font of the given facename is not available on the system.
	 * Although it is not required to have around when resolving font facename, the information can be stored for
	 * when needed to help resolve which font face to use of all available fonts on a system.
	 *
	 *  int 0-255.
	 *
	 *  The following are some of the possible the character sets:
	 INT
	 Value
	 Character Set
	 0 ANSI_CHARSET
	 1 DEFAULT_CHARSET
	 2 SYMBOL_CHARSET
	 77 MAC_CHARSET
	 128 SHIFTJIS_CHARSET
	 129 HANGEUL_CHARSET
	 129 HANGUL_CHARSET
	 130 JOHAB_CHARSET
	 134 GB2312_CHARSET
	 136 CHINESEBIG5_CHARSET
	 161 GREEK_CHARSET
	 162 TURKISH_CHARSET
	 163 VIETNAMESE_CHARSET
	 177 HEBREW_CHARSET
	 178 ARABIC_CHARSET
	 186 BALTIC_CHARSET
	 204 RUSSIAN_CHARSET
	 222 THAI_CHARSET
	 238 EASTEUROPE_CHARSET
	 255 OEM_CHARSET
	 */
	/**
	 * 	condense (Condense)
	 Macintosh compatibility setting. Represents special word/character rendering on Macintosh, when this flag is
	 set. The effect is to condense the text (squeeze it together). SpreadsheetML applications are not required to
	 render according to this flag.
	 */
	/**
	 * extend (Extend)
	 * This element specifies a compatibility setting used for previous spreadsheet applications, resulting in special
	 * word/character rendering on those legacy applications, when this flag is set. The effect extends or stretches out
	 * the text. SpreadsheetML applications are not required to render according to this flag.
	 */
	/**
	 * shadow (Shadow)
	 * Macintosh compatibility setting. Represents special word/character rendering on Macintosh, when this flag is
	 * set. The effect is to render a shadow behind, beneath and to the right of the text. SpreadsheetML applications
	 * are not required to render according to this flag.
	 */
	/**
	 * outline (Outline)
	 * This element displays only the inner and outer borders of each character. This is very similar to Bold in behavior
	 */
	/**
	 * vertAlign (Vertical Alignment)
	 * This element adjusts the vertical position of the text relative to the text's default appearance for this run. It is
	 * used to get 'superscript' or 'subscript' texts, and shall reduce the font size (if a smaller size is available)
	 * accordingly.
	 *
	 * val= An enumeration representing the vertical-alignment setting.
	 * baseline, subscript, superscript
	 Setting this to either subscript or superscript shall make the font size smaller if a
	 smaller font size is available.
	 */

	/**
	 * given an rPr OOXML text run properties, create an ExtenXLS font
	 *
	 * @param bk
	 * @return
	 */
	public Font generateFont( DocumentHandle bk )
	{
		// not using attributes:  charset, family, condense, extend, shadow, scheme, outline==bold
		Font f = new Font( "Arial", 400, 200 );
		Object o;

		o = this.getAttr( "rFont" );
		f.setFontName( (String) o );
		o = this.getAttr( "sz" );
		if( o != null )
		{
			f.setFontHeight( Font.PointsToFontHeight( Double.parseDouble( (String) o ) ) );
		}

		// boolean attributes
		o = this.getAttr( "b" );
		if( o != null )
		{
			f.setBold( true );
		}
		o = this.getAttr( "i" );
		if( o != null )
		{
			f.setItalic( true );
		}
		o = this.getAttr( "u" );
		if( o != null )
		{
			f.setUnderlined( true );
		}
		o = this.getAttr( "strike" );
		if( o != null )
		{
			f.setStricken( true );
		}
		o = this.getAttr( "outline" );
		if( o != null )
		{
			f.setBold( true );
		}
		o = this.getAttr( "vertAlign" );
		if( o != null )
		{
			String s = (String) o;
			if( s.equals( "baseline" ) )
			{
				f.setScript( 0 );
			}
			else if( s.equals( "superscript" ) )
			{
				f.setScript( 1 );
			}
			else if( s.equals( "subscript" ) )
			{
				f.setScript( 2 );
			}
		}
		f.setOOXMLColor( color );
		return f;
	}

	/**
	 * create a new OOXML ss_rPr shared string table text run properties object using attributes from Font f
	 *
	 * @param f
	 * @return
	 */
	public static Ss_rPr createFromFont( Font f )
	{
		// not using attributes:  charset, family, condense, extend, shadow, scheme, outline==bold
		Ss_rPr rp = new Ss_rPr();
		rp.setAttr( "rFont", f.getFontName() );
		rp.setAttr( "sz", new Double( f.getFontHeightInPoints() ).toString() );

		// boolean attributes
		if( f.getBold() )
		{
			rp.setAttr( "b", "" );
		}
		if( f.getItalic() )
		{
			rp.setAttr( "i", "" );
		}
		if( f.getUnderlined() )
		{
			rp.setAttr( "u", "" );
		}
		if( f.getStricken() )
		{
			rp.setAttr( "strike", "" );
		}
		int s = f.getScript();
		if( s == 1 )
		{
			rp.setAttr( "vertAlign", "superscript" );
		}
		else if( s == 2 )
		{
			rp.setAttr( "vertAlign", "subscript" );
		}
		rp.color = f.getOOXMLColor();
		return rp;
	}
}

