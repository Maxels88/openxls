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

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * ColorChoice: choice of: hslClr (Hue, Saturation, Luminance Color Model)
 * prstClr (Preset Color) §5.1.2.2.22 schemeClr (Scheme Color) §5.1.2.2.29
 * scrgbClr (RGB Color Model - Percentage Variant) srgbClr (RGB Color Model -
 * Hex Variant) sysClr (System Color)
 */
// TODO: FINISH: child elements governing color transformations
// finish hslClr
public class ColorChoice implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( ColorChoice.class );
	private static final long serialVersionUID = -4117811305941771643L;
	private SchemeClr s;
	private SrgbClr srgb;
	private SysClr sys;
	private ScrgbClr scrgb;
	private PrstClr p;
	public Theme theme;

	public ColorChoice( SchemeClr s, SrgbClr srgb, SysClr sys, ScrgbClr scrgb, PrstClr p )
	{
		this.s = s;
		this.srgb = srgb;
		this.sys = sys;
		this.scrgb = scrgb;
		this.p = p;
	}

	private void setTheme( Theme t )
	{
		theme = t;
	}

	public ColorChoice( ColorChoice c )
	{
		s = c.s;
		srgb = c.srgb;
		sys = c.sys;
		scrgb = c.scrgb;
		p = c.p;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		SchemeClr s = null;
		SrgbClr srgb = null;
		SysClr sys = null;
		ScrgbClr scrgb = null;
		PrstClr p = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "schemeClr" ) )
					{
						lastTag.push( tnm );
						s = SchemeClr.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "srgbClr" ) )
					{
						lastTag.push( tnm );
						srgb = SrgbClr.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "sysClr" ) )
					{
						lastTag.push( tnm );
						sys = SysClr.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "scrgbClr" ) )
					{
						lastTag.push( tnm );
						scrgb = ScrgbClr.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "prstClr" ) )
					{
						lastTag.push( tnm );
						p = PrstClr.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					/* tnm.equals("hslClr") */// TODO: finish

				}
				else if( eventType == XmlPullParser.END_TAG )
				{ // shouldn't
					// get here
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "ColorChoice.parseOOXML: " + e.toString() );
		}
		ColorChoice c = new ColorChoice( s, srgb, sys, scrgb, p );
		c.setTheme( bk.getWorkBook().getTheme() );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( s != null )
		{
			ooxml.append( s.getOOXML() );
		}
		else if( sys != null )
		{
			ooxml.append( sys.getOOXML() );
		}
		else if( srgb != null )
		{
			ooxml.append( srgb.getOOXML() );
		}
		else if( scrgb != null )
		{
			ooxml.append( scrgb.getOOXML() );
		}
		else if( p != null )
		{
			ooxml.append( p.getOOXML() );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ColorChoice( this );
	}

	public int getColor()
	{
		if( s != null )
		{
			return s.getColor();
		}
		if( sys != null )
		{
			sys.getColor();
		}
		else if( srgb != null )
		{
			srgb.getColor();
		}
		else if( scrgb != null )
		{
			scrgb.getColor();
		}
		else if( p != null )
		{
			p.getColor();
		}
		return -1;

	}

}

/**
 * schemeClr (Scheme Color) This element specifies a color bound to a user's
 * theme. As with all elements which define a color, it is possible to apply a
 * list of color transforms to the base color defined.
 * <p/>
 * accent1 (Accent Color 1) Extra scheme color 1 accent2 (Accent Color 2) Extra
 * scheme color 2 accent3 (Accent Color 3) Extra scheme color 3 accent4 (Accent
 * Color 4) Extra scheme color 4 accent5 (Accent Color 5) Extra scheme color 5
 * accent6 (Accent Color 6) Extra scheme color 6 bg1 (Background Color 1)
 * Semantic background color bg2 (Background Color 2) Semantic additional
 * background color dk1 (Dark Color 1) Main dark color 1 dk2 (Dark Color 2) Main
 * dark color 2 folHlink (Followed Hyperlink Color) Followed Hyperlink Color
 * hlink (Hyperlink Color) Regular Hyperlink Color lt1 (Light Color 1) Main
 * Light Color 1 lt2 (Light Color 2) Main Light Color 2 phClr (Style Color) A
 * color used in theme definitions which means to use the color of the style.
 * tx1 (Text Color 1) Semantic text color tx2 (Text Color 2) Semantic additional
 * text color
 * <p/>
 * parent: many children: many - TODO: handle color transformation children
 * (alpha ...)
 */
class SchemeClr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( SchemeClr.class );
	private static final long serialVersionUID = 2127868578801669266L;
	private String val;
	private ColorTransform clrTransform;
	private Theme theme;

	public SchemeClr( String val, ColorTransform clrTransform, Theme t )
	{
		this.val = val;
		this.clrTransform = clrTransform;
		theme = t;
	}

	public SchemeClr( SchemeClr sc )
	{
		val = sc.val;
		clrTransform = sc.clrTransform;
		theme = sc.theme;
	}

	public static SchemeClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String val = null;
		ColorTransform clrTransform = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "schemeClr" ) )
					{
						val = xpp.getAttributeValue( 0 );
					}
					else
					{
						clrTransform = ColorTransform.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "schemeClr" ) )
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
			log.error( "schemeClr.parseOOXML: " + e.toString() );
		}
		SchemeClr sc = new SchemeClr( val, clrTransform, bk.getWorkBook().getTheme() );
		return sc;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:schemeClr val=\"" + val + "\">" );
		if( clrTransform != null )
		{
			ooxml.append( clrTransform.getOOXML() );
		}
		ooxml.append( "</a:schemeClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SchemeClr( this );
	}

	public int getColor()
	{
		double tint = ((clrTransform == null) ? 0 : clrTransform.getTint());
		Object[] o = Color.parseThemeColor( val, tint, (short) 0, theme );
		return (Integer) o[0];
	}
}

/**
 * sysClr (System Color) This element specifies a color bound to predefined
 * operating system elements. // TODO: appropriate to hard-code???
 * <p/>
 * parent: many children: COLORSTRANSFORM
 */
class SysClr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( SysClr.class );
	private static final long serialVersionUID = 8307422721346337409L;
	private String val; // This simple type specifies a system color
	// value. This color is based upon the value
	// that this color currently has within the
	// system on which the document is being viewed.
	private String lastClr; // Specifies the color value that was last
	// computed by the generating application.
	// Applications shall use the lastClr
	// attribute to determine the absolute value
	// of the last color used if system colors
	// are not supported.
	private ColorTransform clrTransform;

	public SysClr( String val, String lastClr, ColorTransform clrTransform )
	{
		this.val = val;
		this.lastClr = lastClr;
		this.clrTransform = clrTransform;
	}

	public SysClr( SysClr sc )
	{
		val = sc.val;
		lastClr = sc.lastClr;
		clrTransform = sc.clrTransform;
	}

	public static SysClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String val = null;
		String lastClr = null;
		ColorTransform clrTransform = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "sysClr" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "val" ) )
							{ //
								val = xpp.getAttributeValue( i );
							}
							else if( nm.equals( "lastClr" ) )
							{
								lastClr = xpp.getAttributeValue( i );
							}
						}
					}
					else
					{
						clrTransform = ColorTransform.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "sysClr" ) )
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
			log.error( "sysClr.parseOOXML: " + e.toString() );
		}
		SysClr sc = new SysClr( val, lastClr, clrTransform );
		return sc;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:sysClr val=\"" + val + "\"" );
		if( lastClr != null )
		{
			ooxml.append( " lastClr=\"" + lastClr + "\">" );
		}
		if( clrTransform != null )
		{
			ooxml.append( clrTransform.getOOXML() );
		}
		ooxml.append( "</a:sysClr>" );
		// TODO: Handle child elements
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SysClr( this );
	}

	/**
	 * return the color int that represents this system color
	 *
	 * @return
	 */
	public int getColor()
	{
		for( int i = 0; i < OOXMLConstants.systemColors.length; i++ )
		{
			if( OOXMLConstants.systemColors[i][0].equals( val ) )
			{
				return FormatHandle.HexStringToColorInt( OOXMLConstants.systemColors[i][1], (short) 0 );
			}
		}
		return -1;
	}
}

/**
 * srgbClr (RGB Color Model - Hex Variant)
 * <p/>
 * This element specifies a color using the red, green, blue RGB color model.
 * Red, green, and blue is expressed as sequence of hex digits, RRGGBB. A
 * perceptual gamma of 2.2 is used. Specifies the level of red as expressed by a
 * percentage offset increase or decrease relative to the input color.
 * <p/>
 * parent: many children: COLORSTRANSFORM
 */
class SrgbClr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( SrgbClr.class );
	private static final long serialVersionUID = -999813417659560045L;
	private String val;
	private ColorTransform clrTransform;

	public SrgbClr( String val, ColorTransform clrTransform )
	{
		this.val = val;
		this.clrTransform = clrTransform;
	}

	public SrgbClr( SrgbClr sc )
	{
		val = sc.val;
		clrTransform = sc.clrTransform;
	}

	public static SrgbClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String val = null;
		ColorTransform clrTransform = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "srgbClr" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "val" ) )
							{
								val = xpp.getAttributeValue( i );
							}
						}
					}
					else
					{
						clrTransform = ColorTransform.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "srgbClr" ) )
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
			log.error( "srgbClr.parseOOXML: " + e.toString() );
		}
		SrgbClr sc = new SrgbClr( val, clrTransform );
		return sc;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:srgbClr val=\"" + val + "\">" );
		if( clrTransform != null )
		{
			ooxml.append( clrTransform.getOOXML() );
		}
		ooxml.append( "</a:srgbClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SrgbClr( this );
	}

	/**
	 * interpret val and return color int from color table
	 *
	 * @return
	 */
	public int getColor()
	{
		return FormatHandle.HexStringToColorInt( val, (short) 0 );
	}
}

/**
 * scrgbClr (Scheme Color)
 * <p/>
 * This element specifies a color using the red, green, blue RGB color model.
 * Each component, red, green, and blue is expressed as a percentage from 0% to
 * 100%. A linear gamma of 1.0 is assumed.
 * <p/>
 * parent: many children: COLORSTRANSFORM
 */
class ScrgbClr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( ScrgbClr.class );
	private static final long serialVersionUID = -8782954669829478560L;
	private HashMap<String, String> attrs;
	private ColorTransform clrTransform;

	public ScrgbClr( HashMap<String, String> attrs, ColorTransform clrTransform )
	{
		this.attrs = attrs;
		this.clrTransform = clrTransform;
	}

	public ScrgbClr( ScrgbClr sc )
	{
		attrs = sc.attrs;
		clrTransform = sc.clrTransform;
	}

	public static ScrgbClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<>();
		ColorTransform clrTransform = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "scrgbClr" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) ); // r, g, b
						}
					}
					else
					{
						clrTransform = ColorTransform.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "scrgbClr" ) )
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
			log.error( "scrgbClr.parseOOXML: " + e.toString() );
		}
		ScrgbClr sc = new ScrgbClr( attrs, clrTransform );
		return sc;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:scrgbClr" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( clrTransform != null )
		{
			ooxml.append( clrTransform.getOOXML() );
		}
		ooxml.append( "</a:scrgbClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ScrgbClr( this );
	}

	/**
	 * interpret the rbg value into a color int/index into color table
	 *
	 * @return
	 */
	public int getColor()
	{
		// r, g, b in percentages (in 1000th of a percentage)
		double rval = Integer.valueOf( attrs.get( "r" ) ) / 100000; // perecentage
		rval *= 255;
		double gval = Integer.valueOf( attrs.get( "g" ) ) / 100000;
		gval *= 255;
		double bval = Integer.valueOf( attrs.get( "b" ) ) / 100000;
		bval *= 255;
		if( clrTransform != null )
		{
			log.warn( "Scheme Color must process color transforms" );
		}
		java.awt.Color c = new java.awt.Color( (int) rval, (int) gval, (int) bval );
		return FormatHandle.getColorInt( c );
	}
}

/**
 * prstClr (Preset Color) This element specifies a color which is bound to one
 * of a predefined collection of colors.
 * <p/>
 * parent: many children: many
 */
class PrstClr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( PrstClr.class );
	private static final long serialVersionUID = -5773022185972396279L;
	private String val;
	private ColorTransform clrTransform;

	public PrstClr( String val, ColorTransform clrTransform )
	{
		this.val = val;
		this.clrTransform = clrTransform;
	}

	public PrstClr( PrstClr sc )
	{
		val = sc.val;
		clrTransform = sc.clrTransform;
	}

	public static PrstClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String val = null;
		ColorTransform clrTransform = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "prstClr" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "val" ) )
							{
								val = xpp.getAttributeValue( i );
							}
						}
					}
					else
					{
						clrTransform = ColorTransform.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "prstClr" ) )
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
			log.error( "prstClr.parseOOXML: " + e.toString() );
		}
		PrstClr sc = new PrstClr( val, clrTransform );
		return sc;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:prstClr val=\"" + val + "\">" );
		if( clrTransform != null )
		{
			ooxml.append( clrTransform.getOOXML() );
		}
		ooxml.append( "</a:prstClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PrstClr( this );
	}

	public int getColor()
	{
		if( clrTransform != null )
		{
			log.warn( "Preset Color must process color transforms" );
		}
		for( int i = 0; i < FormatHandle.COLORNAMES.length; i++ )
		{
			if( FormatHandle.COLORNAMES[i].equalsIgnoreCase( val ) )
			{
				return i;
			}
		}
		return -1;
	}
}

/**
 * common Color Transformes used by parent elements: schemeColor, systemColor,
 * hslColor, presetColor, sRgbColor, scRgbColor color adjustments are in
 * percentage units <br>
 * // TODO: Finish // comp // inv // gamma // invGamma // gray, red, green, blue
 */
class ColorTransform
{
	private static final Logger log = LoggerFactory.getLogger( ColorTransform.class );
	private int[] lum;
	private int[] hue;
	private int[] sat;
	private int[] alpha;
	private int tint;
	private int shade;

	// TODO: Finish
	// comp
	// inv
	// gamma
	// gray, red, green, blue

	public int getTint()
	{
		return tint;
	}

	public ColorTransform( int[] lum, int[] hue, int[] sat, int[] alpha, int tint, int shade )
	{
		this.lum = lum;
		this.hue = hue;
		this.sat = sat;
		this.alpha = alpha;
		this.tint = tint;
		this.shade = shade;
	}

	/**
	 * parse color transform elements, common children of color-type elements: <br>
	 * schemeColor, systemColor, hslColor, presetColor, sRgbColor, scRgbColor
	 *
	 * @param xpp
	 * @param lastTag
	 * @return
	 */
	public static ColorTransform parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		int[] lum = null;
		int[] hue = null;
		int[] sat = null;
		int[] alpha = null;
		int tint = 0;
		int shade = 0;
		try
		{
			String parentEl = lastTag.peek();

			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "lum" ) )
					{ // This element specifies the input
						// color with its luminance
						// modulated by the given
						// percentage.
						if( lum == null )
						{
							lum = new int[3];
						}
						lum[0] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "lumMod" ) )
					{ // This element specifies
						// the input color with
						// its luminance
						// modulated by the
						// given percentage.
						if( lum == null )
						{
							lum = new int[3];
						}
						lum[1] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "lumOff" ) )
					{ // This element specifies
						// the input color with
						// its luminance
						// shifted, but with its
						// hue and saturation
						// unchanged.
						if( lum == null )
						{
							lum = new int[3];
						}
						lum[2] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "hue" ) )
					{ // This element specifies
						// the input color with the
						// specified hue, but with
						// its saturation and
						// luminance unchanged.
						if( hue == null )
						{
							hue = new int[3];
						}
						hue[0] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "hueMod" ) )
					{
						if( hue == null )
						{
							hue = new int[3];
						}
						hue[1] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "hueOff" ) )
					{
						if( hue == null )
						{
							hue = new int[3];
						}
						hue[2] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "sat" ) )
					{ // This element specifies
						// the input color with the
						// specified saturation, but
						// with its hue and
						// luminance unchanged.
						if( sat == null )
						{
							sat = new int[3];
						}
						sat[0] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "satMod" ) )
					{
						if( sat == null )
						{
							sat = new int[3];
						}
						sat[1] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "satOff" ) )
					{
						if( sat == null )
						{
							sat = new int[3];
						}
						sat[2] = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "shade" ) )
					{ // This element specifies
						// a darker version of
						// its input color
						shade = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "tint" ) )
					{
						tint = Integer.valueOf( xpp.getAttributeValue( 0 ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( parentEl ) )
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
			log.error( "ColorTransform.parseOOXML: " + e.toString() );
		}
		return new ColorTransform( lum, hue, sat, alpha, tint, shade );
	}

	/**
	 * returns the OOXML associated with color transforms of a parent color
	 * element <br>
	 * note that these color transforms must be part of either <br>
	 * schemeColor, systemColor, hslColor, presetColor, sRgbColor, scRgbColor
	 *
	 * @return
	 */
	public StringBuffer getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();

		/**
		 * a:tint Tint a:shade Shade a:comp Complement a:inv Inverse a:gray Gray
		 * a:alpha Alpha a:alphaOff Alpha Offset a:alphaMod Alpha Modulation
		 * a:hue Hue a:hueOff Hue Offset a:hueMod Hue Modulate a:sat Saturation
		 * a:satOff Saturation Offset a:satMod Saturation Modulation a:lum
		 * Luminance a:lumOff Luminance Offset a:lumMod Luminance Modulation
		 * a:red Red a:redOff Red Offset a:redMod Red Modulation a:green Green
		 * a:greenOff Green Offset a:greenMod Green Modification a:blue Blue
		 * a:blueOff Blue Offset a:blueMod Blue Modification a:gamma Gamma
		 * a:invGamma Inverse Gamma
		 */
		if( tint != 0 )
		{
			ooxml.append( "<a:tint val=\"" + tint + "\"/>" );
		}
		if( shade != 0 )
		{
			ooxml.append( "<a:shade val=\"" + shade + "\"/>" );
		}
		// Complement
		// Inverse
		// Gray
		if( alpha != null )
		{
			if( alpha[0] != 0 )
			{
				ooxml.append( "<a:alpha val=\"" + alpha[0] + "\"/>" );
			}
			if( alpha[2] != 0 )
			{
				ooxml.append( "<a:alphaOff val=\"" + alpha[2] + "\"/>" );
			}
			if( alpha[1] != 0 )
			{
				ooxml.append( "<a:alphaMod val=\"" + alpha[1] + "\"/>" );
			}
		}
		if( hue != null )
		{
			if( hue[0] != 0 )
			{
				ooxml.append( "<a:hue val=\"" + hue[0] + "\"/>" );
			}
			if( hue[2] != 0 )
			{
				ooxml.append( "<a:hueOff val=\"" + hue[2] + "\"/>" );
			}
			if( hue[1] != 0 )
			{
				ooxml.append( "<a:hueMod val=\"" + hue[1] + "\"/>" );
			}
		}
		if( sat != null )
		{
			if( sat[0] != 0 )
			{
				ooxml.append( "<a:sat val=\"" + sat[0] + "\"/>" );
			}
			if( sat[2] != 0 )
			{
				ooxml.append( "<a:satOff val=\"" + sat[2] + "\"/>" );
			}
			if( sat[1] != 0 )
			{
				ooxml.append( "<a:satMod val=\"" + sat[1] + "\"/>" );
			}
		}
		if( lum != null )
		{
			if( lum[0] != 0 )
			{
				ooxml.append( "<a:lum val=\"" + lum[0] + "\"/>" );
			}
			if( lum[2] != 0 )
			{
				ooxml.append( "<a:lumOff val=\"" + lum[2] + "\"/>" );
			}
			if( lum[1] != 0 )
			{
				ooxml.append( "<a:lumMod val=\"" + lum[1] + "\"/>" );
			}
		}
		// red,
		// green
		// blue
		// gamma
		// invGamma
		return ooxml;
	}
}