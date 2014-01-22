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
 * style
 * <p/>
 * This element specifies the style information for a shape. This is used to define a shape's appearance in terms of
 * the preset styles defined by the style matrix for the theme.
 * <p/>
 * parent:  sp, pic, cnxSp
 * children (required and in sequence):  lnRef, fillRef, effectRef, fontRef
 */
public class Style implements OOXMLElement
{

	private static final long serialVersionUID = -583023685473342509L;
	private EffectRef effectRef;
	private FontRef fontRef;
	private FillRef fillRef;
	private lnRef lRef;

	public Style( lnRef lr, FillRef flr, EffectRef er, FontRef fr )
	{
		this.lRef = lr;
		this.fillRef = flr;
		this.effectRef = er;
		this.fontRef = fr;
	}

	public Style( Style s )
	{
		this.lRef = s.lRef;
		this.fillRef = s.fillRef;
		this.effectRef = s.effectRef;
		this.fontRef = s.fontRef;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		EffectRef er = null;
		FontRef fr = null;
		FillRef flr = null;
		lnRef lr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "effectRef" ) )
					{
						lastTag.push( tnm );
						er = (EffectRef) EffectRef.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "fontRef" ) )
					{
						lastTag.push( tnm );
						fr = (FontRef) FontRef.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "fillRef" ) )
					{
						lastTag.push( tnm );
						flr = (FillRef) FillRef.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "lnRef" ) )
					{
						lastTag.push( tnm );
						lr = (lnRef) lnRef.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "style" ) )
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
			Logger.logErr( "style.parseOOXML: " + e.toString() );
		}
		Style s = new Style( lr, flr, er, fr );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:style>" );
		ooxml.append( lRef.getOOXML() );
		ooxml.append( fillRef.getOOXML() );
		ooxml.append( effectRef.getOOXML() );
		ooxml.append( fontRef.getOOXML() );
		ooxml.append( "</xdr:style>" );
		return ooxml.toString();
	}

	public String toString()
	{
		return getOOXML();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Style( this );
	}

}

/**
 * effectRef (Effect Reference)
 * This element defines a reference to an effect style within the style matrix. The idx attribute refers the index of
 * an effect style within the effectStyleLst element.
 * <p/>
 * parent: many
 * children:  COLORCHOICE
 */
class EffectRef implements OOXMLElement
{
	private static final long serialVersionUID = -7572271663955122478L;
	private int idx;
	private ColorChoice colorChoice = null;

	protected EffectRef( int idx, ColorChoice c )
	{
		this.idx = idx;
		this.colorChoice = c;
	}

	protected EffectRef( EffectRef er )
	{
		this.colorChoice = er.colorChoice;
		this.idx = er.idx;
	}

	protected static EffectRef parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		int idx = -1;
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "effectRef" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "idx" ) )
							{
								idx = Integer.valueOf( xpp.getAttributeValue( i ) ).intValue();
							}
						}
					}
					else if( tnm.equals( "schemeClr" ) ||
							tnm.equals( "hslClr" ) ||
							tnm.equals( "prstClr" ) ||
							tnm.equals( "scrgbClr" ) ||
							tnm.equals( "srgbClr" ) ||
							tnm.equals( "sysClr" ) )
					{
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "effectRef" ) )
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
			Logger.logErr( "effectRef.parseOOXML: " + e.toString() );
		}
		EffectRef er = new EffectRef( idx, c );
		return er;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:effectRef idx=\"" + idx + "\">" );
		if( colorChoice != null )
		{
			ooxml.append( colorChoice.getOOXML() );
		}
		ooxml.append( "</a:effectRef>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new EffectRef( this );
	}
}

/**
 * fillRef (Fill Reference)
 * This element defines a reference to a fill style within the style matrix. The idx attribute refers to the index of a
 * fill style or background fill style within the presentation's style matrix, defined by the fmtScheme element. A
 * value of 0 or 1000 indicates no background, values 1-999 refer to the index of a fill style within the fillStyleLst
 * element, and values 1001 and above refer to the index of a background fill style within the bgFillStyleLst
 * element. The value 1001 corresponds to the first background fill style, 1002 to the second background fill style,
 * and so on.
 * <p/>
 * parent:  many
 * children: COLORCHOICE
 */
class FillRef implements OOXMLElement
{
	private static final long serialVersionUID = 7691131082710785068L;
	private int idx;
	private ColorChoice colorChoice = null;

	protected FillRef( int idx, ColorChoice c )
	{
		this.idx = idx;
		this.colorChoice = c;
	}

	protected FillRef( FillRef fr )
	{
		this.colorChoice = fr.colorChoice;
		this.idx = fr.idx;
	}

	protected static FillRef parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		int idx = -1;
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "fillRef" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "idx" ) )
							{
								idx = Integer.valueOf( xpp.getAttributeValue( i ) ).intValue();
							}
						}
					}
					else if( tnm.equals( "schemeClr" ) ||
							tnm.equals( "hslClr" ) ||
							tnm.equals( "prstClr" ) ||
							tnm.equals( "scrgbClr" ) ||
							tnm.equals( "srgbClr" ) ||
							tnm.equals( "sysClr" ) )
					{
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk ).cloneElement();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fillRef" ) )
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
			Logger.logErr( "fillRef.parseOOXML: " + e.toString() );
		}
		FillRef fr = new FillRef( idx, c );
		return fr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:fillRef idx=\"" + idx + "\">" );
		if( colorChoice != null )
		{
			ooxml.append( colorChoice.getOOXML() );
		}
		ooxml.append( "</a:fillRef>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FillRef( this );
	}
}

/**
 * fontRef (Font Reference)
 * This element represents a reference to a themed font. When used it specifies which themed font to use along
 * with a choice of color.
 * <p/>
 * parent: many
 * children:  COLORCHOICE
 */
class FontRef implements OOXMLElement
{

	private static final long serialVersionUID = 2907761758443581273L;
	private String idx = null;
	private ColorChoice colorChoice = null;

	protected FontRef( String idx, ColorChoice c )
	{
		this.idx = idx;
		this.colorChoice = c;
	}

	protected FontRef( FontRef fr )
	{
		this.colorChoice = fr.colorChoice;
		this.idx = fr.idx;
	}

	protected static FontRef parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String idx = null;
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "fontRef" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "idx" ) )
							{
								idx = xpp.getAttributeValue( i );
							}
						}
					}
					else if( tnm.equals( "schemeClr" ) ||
							tnm.equals( "hslClr" ) ||
							tnm.equals( "prstClr" ) ||
							tnm.equals( "scrgbClr" ) ||
							tnm.equals( "srgbClr" ) ||
							tnm.equals( "sysClr" ) )
					{
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fontRef" ) )
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
			Logger.logErr( "fontRef.parseOOXML: " + e.toString() );
		}
		FontRef fr = new FontRef( idx, c );
		return fr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:fontRef idx=\"" + idx + "\">" );
		if( colorChoice != null )
		{
			ooxml.append( colorChoice.getOOXML() );
		}
		ooxml.append( "</a:fontRef>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FontRef( this );
	}
}

/**
 * lnRef (Line Reference)
 * This element defines a reference to a line style within the style matrix. The idx attribute refers the index of a
 * line style within the fillStyleLst element
 * <p/>
 * parent: many
 * children: COLORCHOICE
 */
class lnRef implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4349076266006929729L;
	private int idx;
	private ColorChoice colorChoice = null;

	protected lnRef( int idx, ColorChoice c )
	{
		this.idx = idx;
		this.colorChoice = c;
	}

	protected lnRef( lnRef lr )
	{
		this.colorChoice = lr.colorChoice;
		this.idx = lr.idx;
	}

	protected static lnRef parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		int idx = -1;
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "lnRef" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "idx" ) )
							{
								idx = Integer.valueOf( xpp.getAttributeValue( i ) ).intValue();
							}
						}
					}
					else if( tnm.equals( "schemeClr" ) ||
							tnm.equals( "hslClr" ) ||
							tnm.equals( "prstClr" ) ||
							tnm.equals( "scrgbClr" ) ||
							tnm.equals( "srgbClr" ) ||
							tnm.equals( "sysClr" ) )
					{
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "lnRef" ) )
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
			Logger.logErr( "lnRef.parseOOXML: " + e.toString() );
		}
		lnRef lr = new lnRef( idx, c );
		return lr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:lnRef idx=\"" + idx + "\">" );
		if( colorChoice != null )
		{
			ooxml.append( colorChoice.getOOXML() );
		}
		ooxml.append( "</a:lnRef>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new lnRef( this );
	}
}

