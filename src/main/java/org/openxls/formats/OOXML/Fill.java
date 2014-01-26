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

import org.openxls.ExtenXLS.FormatHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.formats.XLS.Xf;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * fill OOXML element
 * <p/>
 * fill (Fill) This element specifies fill formatting
 * <p/>
 * parent: styleSheet/fills element in styles.xml, dxf->fills children: REQ
 * CHOICE OF: patternFill, gradientFill
 */
public class Fill implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Fill.class );
	private static final long serialVersionUID = -4510508531435037641L;
	private PatternFill patternFill = null;
	private GradientFill gradientFill = null;
	private Theme theme = null;

	public Fill( PatternFill p, GradientFill g, Theme t )
	{
		patternFill = p;
		gradientFill = g;
		theme = t;
	}

	public Fill( Fill f )
	{
		if( f.patternFill != null )
		{
			patternFill = (PatternFill) f.patternFill.cloneElement();
		}
		if( f.gradientFill != null )
		{
			gradientFill = (GradientFill) f.gradientFill.cloneElement();
		}
		theme = f.theme;
	}

	/**
	 * create a new Fill from external vals
	 *
	 * @param fs String pattern type
	 * @param fg int color index
	 * @param bg int color index
	 */
	public Fill( String fs, int fg, int bg, Theme t )
	{
		patternFill = new PatternFill( fs, fg, bg );
		theme = t;
	}

	/**
	 * create a new Fill from external vals
	 *
	 * @param i             XLS indexed pattern
	 * @param fg            int color index
	 * @param fgColorCustom
	 * @param bg            int color index
	 */
	public Fill( int pattern, int fg, String fgColorCustom, int bg, String bgColorCustom, Theme t )
	{

		patternFill = new PatternFill( PatternFill.translateIndexedFillPattern( pattern ), fg, fgColorCustom, bg, bgColorCustom );
		theme = t;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, boolean isDxf, WorkBookHandle bk )
	{
		PatternFill p = null;
		GradientFill g = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "patternFill" ) )
					{
						p = PatternFill.parseOOXML( xpp, isDxf, bk );

					}
					else if( tnm.equals( "gradientFill" ) )
					{
						g = GradientFill.parseOOXML( xpp, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fill" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "fill.parseOOXML: " + e.toString() );
		}
		Fill oe = new Fill( p, g, bk.getWorkBook().getTheme() );
		return oe;
	}

	/**
	 * OOXML values are stored with an intiall FF in their value,
	 * <p/>
	 * this method assures that values returned are 6 digits + # for web usage
	 *
	 * @param rgbcolor
	 */
	protected static String transformToWebRGBColor( String rgbcolor )
	{
		if( rgbcolor.indexOf( "#" ) == 0 )
		{
			return rgbcolor;
		}
		if( (rgbcolor.indexOf( "FF" ) == 0) && (rgbcolor.length() == 8) )
		{
			return "#" + rgbcolor.substring( 2, rgbcolor.length() );
		}
		return "#" + rgbcolor;
	}

	/**
	 * OOXML values are stored with an intiall FF in their value,
	 * <p/>
	 * this method assures that values returned are 8 digits with
	 * ff appended to rgb value for ooxml usage
	 *
	 * @param rgbcolor
	 */
	protected static String transformToOOXMLRGBColor( String rgbcolor )
	{
		if( rgbcolor.indexOf( "#" ) == 0 )
		{
			return "FF" + rgbcolor.substring( 1, rgbcolor.length() );
		}
		if( (rgbcolor.indexOf( "FF" ) == 0) && (rgbcolor.length() == 8) )
		{
			return rgbcolor;
		}
		return "FF" + rgbcolor;
	}

	@Override
	public String getOOXML()
	{
		return getOOXML( false );
	}

	/**
	 * dxfs apparently have different pattern fill syntax -- UNDOCUMENTED ****
	 *
	 * @param isDxf if this is an Dxf-generated fill, solid fills are handled
	 *              differently than regular fills
	 * @return
	 */
	public String getOOXML( boolean isDxf )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<fill>" );
		if( patternFill != null )
		{
			ooxml.append( patternFill.getOOXML( isDxf ) );
		}
		if( gradientFill != null )
		{
			ooxml.append( gradientFill.getOOXML() );
		}
		ooxml.append( "</fill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Fill( this );
	}

	/**
	 * returns the OOXML specifying the fill based on this FormatHandle object
	 */
	public static String getOOXML( Xf xf )
	{
		if( xf.getFill() != null )
		{
			return xf.getFill().getOOXML();
		}

		// otherwise, create fill from 2003-style xf
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<fill>" );
		try
		{
			ooxml.append( "<patternFill patternType=\"" + OOXMLConstants.patternFill[xf.getFillPattern()] + "\">" );
		}
		catch( IndexOutOfBoundsException e )
		{
			ooxml.append( "<patternFill>" ); // apparently there are less patterns
			// in xlsx? some other way of
			// storage
		}
		int fg = xf.getForegroundColor();
		if( (fg > -1) && (fg != 64) )
		{
			ooxml.append( "<fgColor rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[fg] ).substring( 1 ) + "\"/>" );
		}
		int bg = xf.getBackgroundColor();
		if( (bg > -1) && (bg != 64) )
		{
			ooxml.append( "<bgColor rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[bg] ).substring( 1 ) + "\"/>" );
		}
		ooxml.append( "</patternFill>" );
		ooxml.append( "\r\n" );
		ooxml.append( "</fill>" );
		ooxml.append( "\r\n" );
		return ooxml.toString();
	}

	/**
	 * returns the OOXML specifying the fill based on fill pattern fs,
	 * foreground color fg, background color bg
	 */
	public static String getOOXML( int fs, int fg, int bg )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<fill>" );
		try
		{
			ooxml.append( "<patternFill patternType=\"" + OOXMLConstants.patternFill[fs] + "\">" );
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
			ooxml.append( "<patternFill>" );
		}
		if( (fg > -1) && (fg != 64) )
		{
			ooxml.append( "<fgColor rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[fg] ).substring( 1 ) + "\"/>" );
			ooxml.append( "\r\n" );
		}
		if( (bg > -1) && (bg != 64) )
		{
			ooxml.append( "<bgColor rgb=\"" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[bg] ).substring( 1 ) + "\"/>" );
			ooxml.append( "\r\n" );
		}
		ooxml.append( "</patternFill>" );
		ooxml.append( "\r\n" );
		ooxml.append( "</fill>" );
		ooxml.append( "\r\n" );
		return ooxml.toString();
	}

	/**
	 * return the foreground color of this fill, if any
	 *
	 * @return
	 */
	public String getFgColorAsRGB( Theme t )
	{
		if( patternFill != null )
		{
			return patternFill.getFgColorAsRGB( t );
		}
		return null;
	}

	/**
	 * return the fg color in indexed (int) representation
	 *
	 * @return
	 */
	public int getFgColorAsInt( Theme t )
	{
		if( patternFill != null )
		{
			return patternFill.getFgColorAsInt( t );
		}
		return 0; // default= black
	}

	/**
	 * return the bg color of this fill, if any
	 *
	 * @return
	 */
	public String getBgColorAsRGB( Theme t )
	{
		if( patternFill != null )
		{
			return patternFill.getBgColorAsRGB( t );
		}
		return null;
	}

	/**
	 * return the bg color in indexed (int) representation
	 *
	 * @return
	 */
	public int getBgColorAsInt( Theme t )
	{
		if( patternFill != null )
		{
			return patternFill.getBgColorAsInt( t );
		}
		return -1;
	}

	/**
	 * sets the foreground fill color via color int
	 *
	 * @param t
	 */
	public void setFgColor( int t )
	{
		if( patternFill != null )
		{
			patternFill.setFgColor( t );
			return;
		}
	}

	/**
	 * sets the foreground fill color to a color string and Excel-2003-mapped
	 * color int
	 *
	 * @param t           Excel-2003-mapped color int for the hex color string
	 * @param colorString hex color string
	 */
	public void setFgColor( int t, String colorString )
	{
		if( patternFill != null )
		{
			patternFill.setFgColor( t, colorString );
		}
		else
		{
			patternFill = new PatternFill( "none", t, colorString, -1, null );
		}
	}

	/**
	 * sets the bg fill color via color int
	 *
	 * @param t
	 */
	public void setBgColor( int t )
	{
		if( patternFill != null )
		{
			patternFill.setBgColor( t );
			return;
		}
	}

	/**
	 * sets the foreground fill color to a color string and Excel-2003-mapped
	 * color int
	 *
	 * @param t           Excel-2003-mapped color int for the hex color string
	 * @param colorString hex color string
	 */
	public void setBgColor( int t, String colorString )
	{
		if( patternFill != null )
		{
			patternFill.setBgColor( t, colorString );
		}
		else
		{
			patternFill = new PatternFill( "none", -1, null, t, colorString );
		}
	}

	public String getFillPattern()
	{
		if( patternFill != null )
		{
			return patternFill.getFillPattern();
		}
		return null;
	}

	/**
	 * return the fill pattern in 2003 int representation
	 *
	 * @return
	 */
	public int getFillPatternInt()
	{
		if( patternFill != null )
		{
			return patternFill.getFillPatternInt();
		}
		return -1;
	}

	/**
	 * sets the fill pattern
	 */
	public void setFillPattern( int t )
	{
		if( patternFill != null )
		{
			patternFill.setFillPattern( t );
		}
	}

	/**
	 * returns true if the background pattern is solid
	 *
	 * @return
	 */
	public boolean isBackgroundSolid()
	{
		if( patternFill != null )
		{
			return (patternFill.getFillPattern().equalsIgnoreCase( "solid" ));
		}
		return false;
	}
}

/**
 * patternFill (Pattern) This element is used to specify cell fill information
 * for pattern and solid color cell fills. For solid cell fills (no pattern),
 * fgColor is used. For cell fills with patterns specified, then the cell fill
 * color is specified by the bgColor element.
 * <p/>
 * parent: fill children: SEQ: fgColor, bgColor
 */
class PatternFill implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( PatternFill.class );
	private static final long serialVersionUID = -4399355217499895956L;
	private String patternType = null;
	private FgColor fgColor = null;
	private BgColor bgColor = null;
	private Theme theme = null;

	public PatternFill( String patternType, FgColor fg, BgColor bg, Theme t )
	{
		this.patternType = patternType;
		fgColor = fg;
		bgColor = bg;
		theme = t;
	}

	public PatternFill( PatternFill p )
	{
		patternType = p.patternType;
		if( p.fgColor != null )
		{
			fgColor = (FgColor) p.fgColor.cloneElement();
		}
		if( p.bgColor != null )
		{
			bgColor = (BgColor) p.bgColor.cloneElement();
		}
		theme = p.theme;
	}

	/**
	 * create a new pattern fill from external vals
	 *
	 * @param patternType String OOXML pattern type
	 * @param fg          int color index
	 * @param bg          int color index
	 */
	public PatternFill( String patternType, int fg, int bg )
	{
		this.patternType = patternType;
		if( (fg > -1) && (fg != 64) )
		{
			HashMap<String, String> attrs = new HashMap<>();
			attrs.put( "rgb", "FF" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[fg] ).substring( 1 ) );
			fgColor = new FgColor( attrs );
		}
		if( (bg > -1) && (bg != 65) )
		{
			HashMap<String, String> attrs = new HashMap<>();
			attrs.put( "rgb", "FF" + FormatHandle.colorToHexString( FormatHandle.COLORTABLE[bg] ).substring( 1 ) );
			bgColor = new BgColor( attrs );
		}
	}

	public PatternFill( String patternType, int fg, String fgCustom, int bg, String bgCustom )
	{
		this.patternType = patternType;
		if( (fg > 0) || (fgCustom != null) )
		{ // 64= default fg color
			HashMap<String, String> attrs = new HashMap<>();
			if( fgCustom == null )
			{
				attrs.put( "indexed", String.valueOf( fg ) );
			}
			else
			{
				attrs.put( "rgb", Fill.transformToOOXMLRGBColor( fgCustom ) );
			}
			fgColor = new FgColor( attrs );
		}
		if( (bg > -1) || (bgCustom != null) )
		{ // 65= default bg color
			HashMap<String, String> attrs = new HashMap<>();
			if( bgCustom == null )
			{
				attrs.put( "indexed", String.valueOf( bg ) );
			}
			else
			{
				attrs.put( "rgb", Fill.transformToOOXMLRGBColor( bgCustom ) );
			}
			bgColor = new BgColor( attrs );
		}
	}

	public static PatternFill parseOOXML( XmlPullParser xpp, boolean isDxf, WorkBookHandle bk )
	{
		String patternType = null; // "none"; // default when missing -- so sez
		// the doc but doesn't appear what Excel
		// does
		/**
		 * APPARENLTY patternFills in dxfs are DIFFERENT AND NOT FOLLOWING THE
		 * DOCUMENTATION on regular patternFills
		 *
		 * APPARENTLY patternType="none" and missing patternType ARE NOT THE
		 * SAME: if Missing, APPARENTLY means to fill with BG color ("solid"
		 * means to fill with FG color)
		 */
		FgColor fg = null;
		BgColor bg = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "patternFill" ) )
					{ // get attributes
						if( xpp.getAttributeCount() > 0 )
						{
							patternType = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "fgColor" ) )
					{
						fg = FgColor.parseOOXML( xpp );
					}
					else if( tnm.equals( "bgColor" ) )
					{
						bg = BgColor.parseOOXML( xpp );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "patternFill" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "patternFill.parseOOXML: " + e.toString() );
		}
		if( isDxf )
		{
			if( patternType == null ) // null apparently does NOT mean none for
			// Dxf's
			{
				patternType = "solid"; // see Dxf and Cf handling: null means
			}
			// solid fill with bg color as cell
			// background
			if( patternType.equals( "solid" ) )
			{
				if( bg != null )
				{// shouldn't!
					fg = new FgColor( bg.getAttrs() ); // so 2003-v can properly
					// set solid cell
					// pattern color
				}
				bg = new BgColor( 64 );
			}
		}
		PatternFill p = new PatternFill( patternType, fg, bg, bk.getWorkBook().getTheme() );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<patternFill" );
		ooxml.append( " patternType=\"" + patternType + "\">" );
		if( fgColor != null )
		{
			ooxml.append( fgColor.getOOXML() );
		}
		if( bgColor != null )
		{
			ooxml.append( bgColor.getOOXML() );
		}
		ooxml.append( "</patternFill>" );
		return ooxml.toString();
	}

	/**
	 * apparently Fill OOXML from Dxf has differnt syntax - UNDOCUMENTED
	 *
	 * @param isDxf if this is an Dxf-generated fill, solid fills are handled
	 *              differently than regular fills
	 * @return
	 */
	public String getOOXML( boolean isDxf )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<patternFill" );
		if( !isDxf )
		{
			ooxml.append( " patternType=\"" + patternType + "\">" );
			if( fgColor != null )
			{
				ooxml.append( fgColor.getOOXML() );
			}
			if( bgColor != null )
			{
				ooxml.append( bgColor.getOOXML() );
			}
		}
		else
		{
			if( patternType.equals( "solid" ) )
			{ // dxf needs "none" or "soild" to
				// have bg color set, not fg
				// color as is normal
				ooxml.append( ">" );
				if( fgColor != null )
				{ // shoudln't!
					BgColor tempbg = new BgColor( fgColor.getAttrs() );
					ooxml.append( tempbg.getOOXML() );
				}
			}
			else
			{
				ooxml.append( " patternType=\"" + patternType + "\">" );
				if( fgColor != null )
				{
					ooxml.append( fgColor.getOOXML() );
				}
				if( bgColor != null )
				{
					ooxml.append( bgColor.getOOXML() );
				}
			}
		}
		ooxml.append( "</patternFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PatternFill( this );
	}

	public String toString()
	{
		return ((patternType != null) ? patternType : "<none>") + " fg:" + ((fgColor != null) ? fgColor.toString() : "<none>") + " bg:" + ((bgColor != null) ? bgColor
				.toString() : "<none>");
	}

	/**
	 * return the pattern type of this Fill in String representation
	 *
	 * @return
	 */
	public String getFillPattern()
	{
		return patternType;
	}

	/**
	 * return the pattern type of this fill pattern in 2003 int representation <br>
	 * BIG NOTE: Apparently doc is wrong in that if patternType is missing it
	 * does NOT == none; instead, it is a SOLID FILL with BG==fill color <br>
	 * THIS METHOD returns -1 in those cases
	 *
	 * @return int pattern fill integer
	 * @see PatternFill.setFillPattern
	 */
	public int getFillPatternInt()
	{
		if( patternType == null ) // a missing entry *should*==none or 0, but
		// apparently is a distinct value in and of
		// itself
		{
			return OOXMLConstants.patternFill.length + 1; // none --SHOULD BE
		}
		// NONE BUT IS
		// 1==SOLID????
		for( int i = 0; i < OOXMLConstants.patternFill.length; i++ )
		{
			if( OOXMLConstants.patternFill[i].equals( patternType ) )
			{
				return i;
			}
		}
		return -1; // none
	}

	/**
	 * sets the pattern type of this Fill <br>
	 * One of: <li>"none", <li>"solid", <li>"mediumGray", <li>"darkGray", <li>
	 * "lightGray", <li>"darkHorizontal", <li>"darkVertical", <li>"darkDown",
	 * <li>"darkUp", <li>"darkGrid", <li>"darkTrellis", <li>"lightHorizontal",
	 * <li>"lightVertical", <li>"lightDown", <li>"lightUp", <li>"lightGrid", <li>
	 * "lightTrellis", <li>"gray125", <li>"gray0625",
	 *
	 * @param s
	 */
	public void setFillPattern( String s )
	{
		patternType = s;
	}

	/**
	 * sets the pattern type of this fill via a 2003-style pattern int <br>
	 * One of: <li>FLSNULL 0x00 No fill pattern <li>FLSSOLID 0x01 Solid <li>
	 * FLSMEDGRAY 0x02 50% gray <li>FLSDKGRAY 0x03 75% gray <li>FLSLTGRAY 0x04
	 * 25% gray <li>FLSDKHOR 0x05 Horizontal stripe <li>FLSDKVER 0x06 Vertical
	 * stripe <li>FLSDKDOWN 0x07 Reverse diagonal stripe <li>FLSDKUP 0x08
	 * Diagonal stripe <li>FLSDKGRID 0x09 Diagonal crosshatch <li>FLSDKTRELLIS
	 * 0x0A Thick Diagonal crosshatch <li>FLSLTHOR 0x0B Thin horizontal stripe
	 * <li>FLSLTVER 0x0C Thin vertical stripe <li>FLSLTDOWN 0x0D Thin reverse
	 * diagonal stripe <li>FLSLTUP 0x0E Thin diagonal stripe <li>FLSLTGRID 0x0F
	 * Thin horizontal crosshatch <li>FLSLTTRELLIS 0x10 Thin diagonal crosshatch
	 * <li>FLSGRAY125 0x11 12.5% gray <li>FLSGRAY0625 0x12 6.25% gray <br>
	 * NOTE: There is a "Special Code" that indicates a missing patternType,
	 * which has a significant meaning in OOXML
	 *
	 * @param t
	 */
	public void setFillPattern( int t )
	{
		patternType = translateIndexedFillPattern( t );
	}

	public static String translateIndexedFillPattern( int pattern )
	{
		String newPattern = null;
		if( pattern == (OOXMLConstants.patternFill.length + 1) )// special code
		{
			newPattern = null;
		}
		else if( (pattern >= 0) && (pattern < OOXMLConstants.patternFill.length) )
		{
			newPattern = OOXMLConstants.patternFill[pattern];
		}

		return newPattern;
	}

	/**
	 * return the foreground color of this fill as an RGB string
	 *
	 * @return
	 */
	public String getFgColorAsRGB( Theme t )
	{
		if( fgColor != null )
		{
			return Fill.transformToWebRGBColor( fgColor.getColorAsRGB( t ) );
		}
		return null;
	}

	/**
	 * return the foreground color of this fill as indexed color int
	 *
	 * @return
	 */
	public int getFgColorAsInt( Theme t )
	{
		if( fgColor != null )
		{
			return fgColor.getColorAsInt( t );
		}
		if( "solid".equals( patternType ) )
		{
			return 0;
		}
		return -1;
	}

	/**
	 * sets the foreground color of this pattern fill via color int
	 *
	 * @param t
	 */
	public void setFgColor( int t )
	{
		if( fgColor != null )
		{
			fgColor.setColor( t );
		}
	}

	/**
	 * sets the foreground fill color to a color string and Excel-2003-mapped
	 * color int
	 *
	 * @param t           Excel-2003-mapped color int for the hex color string
	 * @param colorString hex color string
	 */
	public void setFgColor( int t, String colorString )
	{
		if( (t > 0) || (colorString != null) )
		{ // 64= default fg color
			HashMap<String, String> attrs = new HashMap<>();
			if( colorString == null )
			{
				attrs.put( "indexed", String.valueOf( t ) );
			}
			else
			{
				attrs.put( "rgb", colorString );
			}
			fgColor = new FgColor( attrs );
		}
	}

	/**
	 * return the background color of this fill as an RGB string
	 *
	 * @return
	 */
	public String getBgColorAsRGB( Theme t )
	{
		if( bgColor != null )
		{
			return Fill.transformToWebRGBColor( bgColor.getColorAsRGB( t ) );
		}
		return null;
	}

	/**
	 * return the background color of this fill as an indexed color int
	 *
	 * @return
	 */
	public int getBgColorAsInt( Theme t )
	{
		if( bgColor != null )
		{
			return bgColor.getColorAsInt( t );
		}
		return -1; // default=white
	}

	/**
	 * sets the background color of this pattern fill via color int
	 *
	 * @param t
	 */
	public void setBgColor( int t )
	{
		if( bgColor != null )
		{
			bgColor.setColor( t );
		}
	}

	/**
	 * sets the background fill color to a color string and Excel-2003-mapped
	 * color int
	 *
	 * @param t           Excel-2003-mapped color int for the hex color string
	 * @param colorString hex color string
	 */
	public void setBgColor( int t, String colorString )
	{
		if( (t > 0) || (colorString != null) )
		{    // 65= default bg color
			HashMap<String, String> attrs = new HashMap<>();
			if( colorString == null )
			{
				attrs.put( "indexed", String.valueOf( t ) );
			}
			else
			{
				attrs.put( "rgb", Fill.transformToOOXMLRGBColor( colorString ) );
			}
			bgColor = new BgColor( attrs );
		}
	}

}

/**
 * gradientFill (Gradient) This element defines a gradient-style cell fill.
 * Gradient cell fills can use one or two colors as the end points of color
 * interpolation.
 * <p/>
 * parent: fill children: stop (0 or more) attributes: many
 */
class GradientFill implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( GradientFill.class );
	private static final long serialVersionUID = 3633230059631047503L;
	private HashMap<String, String> attrs = null;
	private ArrayList<Stop> stops = null;

	public GradientFill( HashMap<String, String> attrs, ArrayList<Stop> stops )
	{
		this.attrs = attrs;
		this.stops = stops;
	}

	public GradientFill( GradientFill g )
	{
		attrs = g.attrs;
		stops = g.stops;
	}

	public static GradientFill parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<>();
		ArrayList<Stop> stops = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gradientFill" ) )
					{ // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "stop" ) )
					{
						if( stops == null )
						{
							stops = new ArrayList<>();
						}
						stops.add( Stop.parseOOXML( xpp, bk ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "gradientFill" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "gradientFill.parseOOXML: " + e.toString() );
		}
		GradientFill g = new GradientFill( attrs, stops );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<gradientFill" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( stops != null )
		{
			for( Stop stop : stops )
			{
				ooxml.append( stop.getOOXML() );
			}
		}
		ooxml.append( "</gradientFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GradientFill( this );
	}
}

/**
 * fgColor (Foreground Color) Foreground color of the cell fill pattern. Cell
 * fill patterns operate with two colors: a background color and a foreground
 * color. These combine together to make a patterned cell fill
 * <p/>
 * parent: patternFill chilren: none
 */
class FgColor implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( FgColor.class );
	private static final long serialVersionUID = -1274598491373019241L;
	private HashMap<String, String> attrs = null;

	protected FgColor( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	protected FgColor( FgColor f )
	{
		attrs = (HashMap<String, String>) f.attrs.clone();
	}

	protected FgColor( int c )
	{
		attrs = new HashMap<>();
		attrs.put( "indexed", String.valueOf( c ) );
	}

	protected HashMap<String, String> getAttrs()
	{
		return attrs;
	}

	protected static FgColor parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "fgColor" ) )
					{ // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fgColor" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "fgColor.parseOOXML: " + e.toString() );
		}
		FgColor f = new FgColor( attrs );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<fgColor" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FgColor( this );
	}

	/**
	 * return the Html string representation of this foreground color
	 * <p/>
	 * Note that this will return an HTML correct value, such as #000000.
	 * <p/>
	 * Values set in ooxml need to look like FF000000
	 *
	 * @return
	 */
	protected String getColorAsRGB( Theme t )
	{
		String val = attrs.get( "rgb" );
		if( val != null )
		{
			return val;
		}
		val = attrs.get( "indexed" );
		if( val != null )
		{
			if( Integer.parseInt( val ) == 64 ) // default fg color
			{
				return null; // return "#000000"; //null;
			}
			return Color.parseColor( val, Color.COLORTYPEINDEXED, FormatHandle.colorFOREGROUND, t );
		}
		val = attrs.get( "theme" );
		if( val != null )
		{
			return Color.parseColor( val, Color.COLORTYPETHEME, FormatHandle.colorFOREGROUND, t );
		}
		val = attrs.get( "auto" );
		if( val != null )
		{
			return "#000000";
		}
		return null;
	}

	/**
	 * sets the foreground color to the indexed color integer
	 *
	 * @param c
	 */
	protected void setColor( int c )
	{
		attrs.clear();
		attrs.put( "indexed", String.valueOf( c ) );
	}

	/**
	 * returns the fg color as an indexed color int
	 */
	protected int getColorAsInt( Theme t )
	{
		String val = attrs.get( "auto" );
		if( val != null )
		{
			return 0;
		}
		val = attrs.get( "rgb" );
		if( val != null )
		{
			return Color.parseColorInt( val, Color.COLORTYPERGB, FormatHandle.colorFOREGROUND, t );
		}
		val = attrs.get( "indexed" );
		if( val != null )
		{
			return Integer.valueOf( val );
		}
		val = attrs.get( "theme" );
		if( val != null )
		{
			return Color.parseColorInt( val, Color.COLORTYPETHEME, FormatHandle.colorFOREGROUND, t );
		}
		return -1;
	}

	public String toString()
	{
		if( attrs != null )
		{
			String s = attrs.toString();
			return s.substring( 1, s.length() - 1 );
		}
		return "none";
	}
}

/**
 * bgColor (Background Color) Background color of the cell fill pattern. Cell
 * fill patterns operate with two colors: a background color and a foreground
 * color. These combine together to make a patterned cell fill.
 * <p/>
 * parent: patternFill children: none
 */
class BgColor implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( BgColor.class );
	private static final long serialVersionUID = 43028503491956217L;
	private HashMap<String, String> attrs = null;

	protected BgColor( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	protected BgColor( BgColor f )
	{
		attrs = (HashMap<String, String>) f.attrs.clone();
	}

	protected BgColor( int c )
	{
		attrs = new HashMap<>();
		attrs.put( "indexed", String.valueOf( c ) );
	}

	protected HashMap<String, String> getAttrs()
	{
		return attrs;
	}

	public static BgColor parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "bgColor" ) )
					{ // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "bgColor" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "bgColor.parseOOXML: " + e.toString() );
		}
		BgColor f = new BgColor( attrs );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<bgColor" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new BgColor( this );
	}

	/**
	 * return the Html string representation of this background color
	 *
	 * @return
	 */
	protected String getColorAsRGB( Theme t )
	{
		String val = attrs.get( "rgb" );
		if( val != null )
		{
			return val;
		}
		val = attrs.get( "indexed" );
		if( val != null )
		{
			if( Integer.parseInt( val ) == 65 ) // default bg color
			{
				return null; // return "#FFFFFF"; //return null;
			}
			return Color.parseColor( val, Color.COLORTYPEINDEXED, FormatHandle.colorBACKGROUND, t );
		}
		val = attrs.get( "theme" );
		if( val != null )
		{
			return Color.parseColor( val, Color.COLORTYPETHEME, FormatHandle.colorBACKGROUND, t );
		}
		return null;
	}

	/**
	 * sets the background color to the indexed color integer
	 */
	protected void setColor( int c )
	{
		attrs.clear();
		attrs.put( "indexed", String.valueOf( c ) );
	}

	/**
	 * returns the bg color as an indexed color int
	 */
	protected int getColorAsInt( Theme t )
	{
		String val = attrs.get( "rgb" );
		if( val != null )
		{
			return Color.parseColorInt( val, Color.COLORTYPERGB, FormatHandle.colorBACKGROUND, t );
		}
		val = attrs.get( "indexed" );
		if( val != null )
		{
			return Integer.valueOf( val );
		}
		val = attrs.get( "theme" );
		if( val != null )
		{
			return Color.parseColorInt( val, Color.COLORTYPETHEME, FormatHandle.colorBACKGROUND, t );
		}
		return 64; // default
	}

	public String toString()
	{
		if( attrs != null )
		{
			String s = attrs.toString();
			return s.substring( 1, s.length() - 1 );
		}
		return "none";
	}
}

/**
 * stop (Gradient Stop) One of a sequence of two or more gradient stops,
 * constituting this gradient fill.
 * <p/>
 * parent: gradientFill children: color REQ
 */
class Stop implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Stop.class );
	private static final long serialVersionUID = -9215564484103992694L;
	private String position = null;
	private Color c = null;

	public Stop( String position, Color c )
	{
		this.position = position;
		this.c = c;
	}

	public Stop( Stop s )
	{
		position = s.position;
		c = s.c;
	}

	public static Stop parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		String position = null;
		Color c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "stop" ) )
					{
						position = xpp.getAttributeValue( 0 ); // position=
						// REQUIRED
					}
					else if( tnm.equals( "color" ) )
					{
						c = (Color) Color.parseOOXML( xpp, (short) -1, bk ).cloneElement();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "stop" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "stop.parseOOXML: " + e.toString() );
		}
		Stop s = new Stop( position, c );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<stop" );
		ooxml.append( " position=\"" + position + "\"" );
		ooxml.append( ">" );
		if( c != null )
		{
			ooxml.append( c.getOOXML() );
		}
		ooxml.append( "</stop>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Stop( this );
	}

}

