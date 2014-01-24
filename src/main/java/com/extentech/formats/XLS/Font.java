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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.Color;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.io.UnsupportedEncodingException;

/**
 * <b>Font: Font Description (231h)</b><br>
 * <p/>
 * Font records describe a font in the workbook
 * <p/>
 * <p>
 * <p/>
 * <pre>
 *     offset  name            size    contents
 *     ---
 *     4       dyHeight        2       Height in 1/20 point (twips)
 *     6       grbit           2       attributes
 * 									grbit Mask Contents
 * 										0 0001H 1 = Characters are bold
 * 										1 0002H 1 = Characters are italic
 * 										2 0004H 1 = Characters are underlined
 * 										3 0008H 1 = Characters are struck out
 *     8       icv             2       index to color palette
 *     10      bls             2       bold style (weight)  100-1000 default is 190h norm 2bch bold
 *     12      sss             2       super/sub (0 = none, 1 = super, 2 = sub)
 *     14      uls             1       Underline Style (0 = none, 1 = single, 2 = double, 21h = single acctg, 22h = dble acctg)
 *     15      bFamily         1       Font Family (WinAPI LOGFONT struct)
 *     16      bCharSet        1       Characterset (WinAPI LOGFONT struct)
 *     17      reserved        0
 *     18      cch             1       Length of font name
 * 19      rgch            var     Font name
 *
 * </p>
 * </pre>
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 *
 * @see XF
 * @see FORMAT
 */

public final class Font extends com.extentech.formats.XLS.XLSRecord implements FormatConstants
{
	private static final Logger log = LoggerFactory.getLogger( Font.class );
	private static final long serialVersionUID = -398444997553403671L;
	private short grbit = -1;
	private short cch = -1;
	private short dyHeight = -1;
	private short icv = -1;
	private short bls = -1;
	private short sss = -1;
	private short uls = -1;
	private short bFamily = -1;
	private short bCharSet;
	private String fontName = "";
	// OOXML specifics:
	private Color customColor = null; // holds custom color (OOXML or other use) 
	private boolean condensed;
	private boolean extended;

	// grbit flags
	static final int BITMASK_BOLD = 0x0001;
	static final int BITMASK_ITALIC = 0x0002;
	static final int BITMASK_UNDERLINED = 0x0004;
	static final int BITMASK_STRIKEOUT = 0x0008;

	// charset values
	static final int ANSI_CHARSET = 0;
	static final int DEFAULT_CHARSET = 1;
	static final int SYMBOL_CHARSET = 2;
	static final int SHIFTJIS_CHARSET = 128;
	static final int HANGEUL_CHARSET = 129;
	static final int HANGUL_CHARSET = 129;
	static final int GB2312_CHARSET = 134;
	static final int CHINESEBIG5_CHARSET = 136;
	static final int OEM_CHARSET = 255;
	static final int JOHAB_CHARSET = 130;
	static final int HEBREW_CHARSET = 177;
	static final int ARABIC_CHARSET = 178;
	static final int GREEK_CHARSET = 161;
	static final int TURKISH_CHARSET = 162;
	static final int VIETNAMESE_CHARSET = 163;
	static final int THAI_CHARSET = 222;
	static final int EASTEUROPE_CHARSET = 238;
	static final int RUSSIAN_CHARSET = 204;
	static final int MAC_CHARSET = 77;
	static final int BALTIC_CHARSET = 186;

	/**
	 * Initialize the font record
	 */
	@Override
	public void init()
	{
		super.init();
		dyHeight = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );// Height
		// in
		// 1/20
		// point
		grbit = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );// attributes
		icv = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );// index
		// to
		// color
		// palette
		bls = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );// bold
		// style
		// (weight)
		// 100-1000
		// default
		// is
		// 190h
		// norm
		// 2bch
		// bold
		sss = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );// super/sub
		// (0 =
		// none,
		// 1 =
		// super,
		// 2 =
		// sub)
		uls = getByteAt( 10 );// Underline Style (0 = none, 1 = single, 2 =
		// double, 21h = single acctg, 22h = dble
		// acctg)
		bFamily = getByteAt( 11 );// Font Family (WinAPI LOGFONT struct)
		/**
		 * lfCharSet
		 *
		 * The character set. The following values are predefined.
		 *
		 * ANSI_CHARSET BALTIC_CHARSET xx CHINESEBIG5_CHARSET xx DEFAULT_CHARSET
		 * xx EASTEUROPE_CHARSET GB2312_CHARSET xx GREEK_CHARSET xx
		 * HANGUL_CHARSET xx MAC_CHARSET xx OEM_CHARSET xx RUSSIAN_CHARSET xx
		 * SHIFTJIS_CHARSET xx SYMBOL_CHARSET xx TURKISH_CHARSET xx
		 * VIETNAMESE_CHARSET
		 *
		 *
		 * Korean language edition of Windows: JOHAB_CHARSET
		 *
		 * Middle East language edition of Windows: ARABIC_CHARSET
		 * HEBREW_CHARSET
		 *
		 * Thai language edition of Windows: THAI_CHARSET
		 *
		 * The OEM_CHARSET value specifies a character set that is
		 * operating-system dependent.
		 *
		 * DEFAULT_CHARSET is set to a value based on the current system locale.
		 * For example, when the system locale is English (United States), it is
		 * set as ANSI_CHARSET.
		 *
		 * Fonts with other character sets may exist in the operating system. If
		 * an application uses a font with an unknown character set, it should
		 * not attempt to translate or interpret strings that are rendered with
		 * that font.
		 *
		 * This parameter is important in the font mapping process. To ensure
		 * consistent results, specify a specific character set. If you specify
		 * a typeface name in the lfFaceName member, make sure that the
		 * lfCharSet value matches the character set of the typeface specified
		 * in lfFaceName.
		 */
		bCharSet = (short) ByteTools.readUnsignedShort( getByteAt( 12 ), (byte) 0 );// Characterset (WinAPI LOGFONT struct)

		// this.getData()[13]= 0;// set byte to 0 for comparisons
		// get the Name
		int pos = 14;
		cch = getByteAt( pos++ );
		int buflen = cch * 2;
		pos++;
		boolean compressed = false;

		if( (buflen + pos) >= getLength() )
		{
			buflen = getLength() - pos;
			compressed = true;
		}

		if( buflen < 0 )
		{
			log.warn( "could not parse font: length reported as " + buflen );
			return;
		}

		byte[] namebytes = getBytesAt( pos, buflen );
		if( !compressed )
		{
			pos = 0;
			try
			{
				fontName = new String( namebytes, XLSConstants.UNICODEENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				log.error( "Font name decoding failed.", e );
			}
		}
		else
		{ // compressed
			fontName = new String( namebytes );
		}
	}

	/**
	 * Get if the font record is striken our or not
	 *
	 * @return
	 */
	public boolean getStricken()
	{
		return ((grbit & BITMASK_STRIKEOUT) == BITMASK_STRIKEOUT);
	}

	public void setStricken( boolean b )
	{
		if( b )
		{
			grbit = (short) (grbit | BITMASK_STRIKEOUT);
		}
		else
		{
			int grbittemp = grbit ^ BITMASK_STRIKEOUT;
			grbit = (short) (grbittemp & grbit);
		}
		setGrbit();
	}

	/**
	 * Get if the font is italic or not
	 *
	 * @return
	 */
	public boolean getItalic()
	{
		int isItalic = grbit & BITMASK_ITALIC;
		if( isItalic == BITMASK_ITALIC )
		{
			return true;
		}
		return false;
	}

	public void setItalic( boolean b )
	{
		if( b )
		{
			grbit = (short) (grbit | BITMASK_ITALIC);
		}
		// false: need to account for multiple font formats - add step 2: &
		// original grbit
		else
		{
			int grbittemp = grbit ^ BITMASK_ITALIC;
			grbit = (short) (grbittemp & grbit);
		}
		setGrbit();
	}

	/**
	 * Get if the font is underlined or not
	 *
	 * @return
	 */
	public boolean getUnderlined()
	{
		return ((grbit & BITMASK_UNDERLINED) == BITMASK_UNDERLINED);
	}

	public void setUnderlined( boolean b )
	{
		if( b )
		{
			grbit = (short) (grbit | BITMASK_UNDERLINED);
		}
		else
		{
			int grbittemp = grbit ^ BITMASK_UNDERLINED;
			grbit = (short) (grbittemp & grbit);
		}
		setGrbit();
		setUnderlineStyle( (byte) 1 ); // 20070821 KSC: should also set underline
		// style ...
	}

	/**
	 * get if the font is bold or not
	 *
	 * @return
	 */
	public boolean getBold()
	{
		return ((grbit & BITMASK_BOLD) == BITMASK_BOLD);
	}

	/**
	 * Set or unset bold attribute of the font record
	 *
	 * @param b
	 */
	public void setBold( boolean b )
	{
		if( data == null )
		{
			setData( getData() );
		}
		if( b )
		{
			byte[] boldbytes = ByteTools.shortToLEBytes( (short) 0x2bc );
			System.arraycopy( boldbytes, 0, data, 6, 2 );
			bls = 0x2bc;
			grbit = (short) (grbit | BITMASK_BOLD);
		}
		else
		{
			byte[] boldbytes = ByteTools.shortToLEBytes( (short) 0x190 );
			System.arraycopy( boldbytes, 0, data, 6, 2 );
			bls = 0x190;
			int grbittemp = grbit ^ BITMASK_BOLD;
			grbit = (short) (grbittemp & grbit);
		}
		setGrbit();
	}

	/**
	 * update the Grbit bytes in the underlying byte stream
	 */
	public void setGrbit()
	{
		byte[] data = getData();
		byte[] b = ByteTools.shortToLEBytes( grbit );
		System.arraycopy( b, 0, data, 2, 2 );
		setData( data );

	}

	private int tableidx = -1;

	public int getIdx()
	{
		return tableidx;
	}

	/**
	 * @param idx
	 */
	public void setIdx( int idx )
	{
		tableidx = idx;
	}

	/**
	 * add to Fonts table in Workbook
	 */
	@Override
	public void setWorkBook( WorkBook b )
	{
		super.setWorkBook( b );
		if( tableidx == -1 )
		{
			tableidx = getWorkBook().addFont( this );
		}
	}

	public String toString()
	{
		return fontName + "," + bls + "," + dyHeight + " " + getColorAsColor() + " font style:[" + getBold() + getItalic() + getStricken() + getUnderlined() + getColor() + getUnderlineStyle() + "]";
	}

	public String getFontName()
	{
		return fontName;
	}

	/**
	 * Get an int representing the underline style of this record, matches int
	 * records in FormatConstants.STYLE_UNDERLINE*****
	 *
	 * @return
	 * @see
	 */
	public int getUnderlineStyle()
	{
		return getData()[10];
	}

	/**
	 * Set the underline style of this font recotd
	 *
	 * @param styl
	 */
	public void setUnderlineStyle( byte styl )
	{
		uls = styl;
		getData()[10] = styl;
	}

	public Font()
	{
	}

	/**
	 * Create a New Font from the String definition.
	 * <p/>
	 * Roughly matches the functionality of the java.awt.Font class.
	 *
	 * @param String font name
	 * @param int    font style
	 * @param int    font size in Points
	 */
	public Font( String nm, int stl, int sz )
	{
		byte[] bl = new byte[]{
				-56, 0, 0, 0, -1, 127, -112, 1, 0, 0, 0, 0, 0, 0, 5, 1, 65, 0, 114, 0, 105, 0, 97, 0, 108, 0,
		};
		setOpcode( FONT );
		setLength( (short) (bl.length) );
		setData( bl );
		init();
		setFontName( nm );
		setFontWeight( stl );
		setFontHeight( sz );
	}

	/**
	 * Set the Font name.
	 * <p/>
	 * To be valid, this font name must be available on the client system.
	 */
	public void setFontName( String fn )
	{
		byte[] namebytes = null;
		try
		{
			namebytes = fn.getBytes( XLSConstants.UNICODEENCODING );
		}
		catch( UnsupportedEncodingException e )
		{
			log.error( "setting Font Name using Default Encoding failed: " + e );
			namebytes = fn.getBytes();
		}
		cch = (short) (namebytes.length / 2);
		fontName = fn;
		byte[] newdata = new byte[namebytes.length + 16];
		System.arraycopy( getBytesAt( 0, 13 ),
		                  0,
		                  newdata,
		                  0,
		                  13 );// 20061027 KSC: keep 13th byte for sake of comparisons - 20070816 - revert to original
		System.arraycopy( getBytesAt( 0, 14 ), 0, newdata, 0, 14 );
		newdata[14] = (byte) cch;
		newdata[15] = 1;
		System.arraycopy( namebytes, 0, newdata, 16, namebytes.length );
		setData( newdata );
		init();
	}

	/**
	 *
	 */
	/**
	 * Set the super/sub script for the Font
	 * <p/>
	 * super/sub (0 = none, 1 = super, 2 = sub)
	 */
	public void setScript( int ss )
	{
		if( data == null )
		{
			setData( getData() );
		}
		byte[] newss = ByteTools.shortToLEBytes( (short) ss );
		System.arraycopy( newss, 0, data, 8, 2 );
		sss = (short) ss;
	}

	/**
	 * returns the super/sub script for the Font
	 *
	 * @return int (0 = none, 1 = super, 2 = sub)
	 */
	public int getScript()
	{
		return sss;
	}

	/**
	 * Set the font color via index into 2003 color table
	 */
	public void setColor( int cl )
	{
		if( data == null )
		{
			setData( getData() );
		}
		if( cl != icv )
		{ // don't do it if the font is already this color
			byte[] newcl = ByteTools.shortToLEBytes( (short) cl );
			System.arraycopy( newcl, 0, data, 4, 2 );
			icv = (short) cl;
		}
		if( customColor != null )
		{
			customColor.setColorInt( cl );
		}
	}

	/**
	 * Sets the font color via java.awt.Color
	 */
	public void setColor( java.awt.Color color )
	{
		if( customColor != null )
		{
			customColor.setColor( color );
		}
		else
		{
			customColor = new Color( color, "color", getWorkBook().getTheme() );
		}
		icv = (short) customColor.getColorInt();
		byte[] newcl = ByteTools.shortToLEBytes( icv );
		System.arraycopy( newcl, 0, data, 4, 2 );
	}

	/**
	 * Sets the font color via a web-compliant Hex String
	 */
	public void setColor( String clr )
	{
		if( customColor != null )
		{
			customColor.setColor( clr );
		}
		else
		{
			customColor = new Color( clr, "color", getWorkBook().getTheme() );
		}
		icv = (short) customColor.getColorInt();
		byte[] newcl = ByteTools.shortToLEBytes( icv );
		System.arraycopy( newcl, 0, data, 4, 2 );
	}

	/**
	 * Set the size of the font in 1/20 point units
	 */
	public void setFontHeight( int ht )
	{
		if( data == null )
		{
			setData( getData() );
		}
		byte[] newht = ByteTools.shortToLEBytes( (short) ht );
		System.arraycopy( newht, 0, data, 0, 2 );
		dyHeight = (short) ht;
	}

	/**
	 * utility to convert points to correct font height
	 *
	 * @param h
	 * @return
	 */
	public static int PointsToFontHeight( double h )
	{
		return (int) (h * 20);
	}

	public static double FontHeightToPoints( int h )
	{
		return h / 20.0;
	}

	public int getFontWeight()
	{
		return bls;
	}

	public int getFontHeight()
	{
		return dyHeight;
	}

	public double getFontHeightInPoints()
	{
		return dyHeight / 20.0;
	}

	/**
	 * Get the color for this Font as a avt.Color
	 *
	 * @return
	 */
	public java.awt.Color getColorAsColor()
	{
		if( customColor != null )
		{
			return customColor.getColorAsColor();
		}
		// If icv is System window text color=7FFF, default fg color or default tooltip text color, return black
		if( (icv == 0x7FFF) || (icv == 0x40) || (icv == 0x51) )
		{
			return java.awt.Color.BLACK;
		}
		if( icv > FormatHandle.COLORTABLE.length )
		{
			return java.awt.Color.BLACK;
		}
		if( getWorkBook() == null )
		{
			return FormatHandle.COLORTABLE[icv];
		}
		/* notes: special icv values:
		0x0040	Default foreground color. This is the window text color in the sheet display.
		0x0041	Default background color. This is the window background color in the sheet display and is the default background color for a cell.
		0x004D	Default chart foreground color. This is the window text color in the chart display.
		0x004E	Default chart background color. This is the window background color in the chart display.
		0x004F	Chart neutral color which is black, an RGB value of (0,0,0).
		0x0051	ToolTip text color. This is the automatic font color for comments.
		0x7FFF	Font automatic color. This is the window text color.
		*/
		return getColorTable()[icv];
	}

	/**
	 * Get the color of this Font as a web-compliant Hex String
	 */
	public String getColorAsHex()
	{
		if( (customColor != null) && (customColor.getColorAsOOXMLRBG() != null) )
		{
			return "#" + customColor.getColorAsOOXMLRBG().substring( 2 ); // remove "FF" from beginning
		}
		return FormatHandle.colorToHexString( getColorAsColor() );
	}

	/**
	 * returns the font color as an OOXML-compliant Hex Stringf	 *
	 *
	 * @return
	 */
	public String getColorAsOOXMLRBG()
	{
		String rgbcolor = getColorAsHex();
		return "FF" + rgbcolor.substring( 1, rgbcolor.length() );
	}

	/**
	 * gets the color of this font as an index into excel 2003 color table
	 *
	 * @return int
	 * @deprecated use getColor()
	 */
	public int getFontColor()
	{
		return getColor();
	}

	/**
	 * gets the color of this font as an index into excel 2003 color table
	 *
	 * @return int
	 */
	public int getColor()
	{
		if( customColor != null )
		{
			return customColor.getColorInt();
		}
		if( icv == 32767 ) // this is a value for system font color, default
		// to black
		{
			return 0;
		}
		return icv;
	}

	/**
	 * Set the weight of the font in 1/20 point units 100-1000 range.
	 */
	public void setFontWeight( int wt )
	{
		if( data == null )
		{
			setData( getData() );
		}
		byte[] newwt = ByteTools.shortToLEBytes( (short) wt );
		System.arraycopy( newwt, 0, data, 6, 2 );
		bls = (short) wt;
	}

	/**
	 * Get if the font is bold or not
	 */
	public boolean getIsBold()
	{
		if( bls > 0x190 )
		{
			return true;
		}
		return false;
	}

	/**
	 * @return an XML descriptor for this Font
	 */
	// changed from 'getFontInfoXML'
	public String getXML()
	{
		return getXML( false );
	}

	/**
	 * return an XML desciptor for this font
	 *
	 * @param convertToUnicodeFont if true, font family will be changed to ArialUnicodeMS
	 *                             (standard unicode) for non-ascii fonts
	 * @return
	 */
	public String getXML( boolean convertToUnicodeFont )
	{
		StringBuffer sb = new StringBuffer( "" );
		if( !convertToUnicodeFont || !isUnicodeCharSet() )
		{
			sb.append( "name=\"" + StringTool.convertXMLChars( getFontName() ) + "\"" );
		}
		else
		{
			sb.append( "name=\"ArialUnicodeMS\"" );
		}
		sb.append( " size=\"" + getFontHeightInPoints() + "\"" );
		sb.append( " color=\"" + FormatHandle.colorToHexString( getColorAsColor() ) + "\"" );
		sb.append( " weight=\"" + getFontWeight() + "\"" );
		if( getIsBold() )
		{
			sb.append( " bold=\"1\"" );
		}
		if( getUnderlineStyle() != Font.STYLE_UNDERLINE_NONE )
		{
			sb.append( " underline=\"" + getUnderlineStyle() + "\"" );
		}
		return sb.toString();
	}

	/**
	 * return true if font f matches key attributes of this font
	 *
	 * @param f
	 * @return
	 */
	public boolean matches( Font f )
	{
		return (fontName.equals( f.fontName ) && (dyHeight == f.dyHeight) && (bls == f.bls) && (getColor() == f.getColor()) && (sss == f.sss) && (uls == f.uls) && (grbit == f.grbit));
	}

	/**
	 * store OOXML font color
	 *
	 * @param c
	 */
	public void setOOXMLColor( Color c )
	{
		if( c != null )
		{
			setColor( c.getColorInt() );
		}
		customColor = c;
	}

	/**
	 * return the OOXML font color element
	 *
	 * @return
	 */
	public Color getOOXMLColor()
	{
		return customColor;
	}

	/**
	 * set whether this font is Condensed (OOXML-specific) Macintosh
	 * compatibility setting.
	 *
	 * @param condensed
	 */
	public void setCondensed( boolean condensed )
	{
		this.condensed = condensed;
	}

	/**
	 * return whether this font is Condensed (OOXML-specific) Macintosh
	 * compatibility setting.
	 *
	 * @return
	 */
	public boolean isCondensed()
	{
		return condensed;
	}

	/**
	 * set whether this font is Extended (OOXML-specific) Macintosh
	 * compatibility setting.
	 *
	 * @param expanded
	 */
	public void setExtended( boolean extended )
	{
		this.extended = extended;
	}

	/**
	 * return whether this font is Expanded (OOXML-specific) Macintosh
	 * compatibility setting.
	 *
	 * @return expanded
	 */
	public boolean isExtended()
	{
		return extended;
	}

	/**
	 * generate the OOXML to define this Font
	 *
	 * @return
	 */
	// TODO: family, scheme
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<font>" );
		if( getIsBold() )
		{
			ooxml.append( "<b/>" );
		}
		if( getItalic() )
		{
			ooxml.append( "<i/>" );
		}
		if( getUnderlined() )
		{
			int u = getUnderlineStyle();
			if( u == 1 )// the default
			{
				ooxml.append( "<u/>" );
			}
			else if( u == 2 )
			{
				ooxml.append( "<u val=\"double\"/>" );
			}
			else if( u == 0x21 )
			{
				ooxml.append( "<u val=\"singleAccounting\"/>" );
			}
			else if( u == 0x22 )
			{
				ooxml.append( "<u val=\"doubleAccounting\"/>" );
			}
		}
		if( getStricken() )
		{
			ooxml.append( "<strike/>" );
		}
		if( !isCondensed() )
		{
			ooxml.append( "<condense val=\"0\"/>" );
		}
		if( !isExtended() )
		{
			ooxml.append( "<extend val=\"0\"/>" );
		}
		Color c = getOOXMLColor();
		if( c != null )
		{
			ooxml.append( c.getOOXML() );
		}
		else
		{
			// KSC: modify due to certain XLS->XLSX issues with automatic color
			if( (icv != 9) && (icv != 64) )
			{ // leave automatic "blank"
				int cl = getColor();
				if( cl > 0 )
				{
					ooxml.append( "<color rgb=\"FF" + FormatHandle.colorToHexString( getColorTable()[cl] ).substring( 1 ) + "\"/>" );
				}
			}
		}
		double sz = getFontHeightInPoints();
		if( sz > 0 ) // for incremental styles, font size may not be set
		{
			ooxml.append( "<sz val=\"" + sz + "\"/>" );
		}
		String n = getFontName();
		if( (n != null) && !n.equals( "" ) ) // for incremental styles, font name may
		// not be set
		{
			ooxml.append( "<name val=\"" + n + "\"/>" );
		}
		// TODO: family val= # (see OOXMLConstants)

		ooxml.append( "</font>" );
		ooxml.append( "\r\n" );
		return ooxml.toString();
	}

	/**
	 * parse incoming OOXML into a Font object
	 *
	 * @param xpp
	 * @return
	 */
	// TODO: family, scheme
	public static Font parseOOXML( XmlPullParser xpp, WorkBookHandle bk )
	{
		Color c = null;
		String sz = null;
		String name = "";
		Object u = null;
		boolean b = false;
		boolean strike = false;
		boolean ital = false;
		boolean condense = false;
		boolean expand = false;
		try
		{
			int eventType = xpp.next();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "sz" ) )
					{
						sz = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "name" ) )
					{
						name = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "b" ) )
					{
						if( xpp.getAttributeCount() == 0 )
						{
							b = true;
						}
						else
						{
							b = (xpp.getAttributeValue( 0 ).equals( "1" ));
						}
					}
					else if( tnm.equals( "i" ) )
					{
						if( xpp.getAttributeCount() == 0 )
						{
							ital = true;
						}
						else
						{
							ital = (xpp.getAttributeValue( 0 ).equals( "1" ));
						}
					}
					else if( tnm.equals( "u" ) )
					{
						if( xpp.getAttributeCount() == 0 )
						{
							u = true;
						}
						else
						{
							u = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "strike" ) )
					{
						strike = true;
					}
					else if( tnm.equals( "condense" ) )
					{
						condense = false;
					}
					else if( tnm.equals( "expand" ) )
					{
						expand = false;
					}
					else if( tnm.equals( "color" ) )
					{
						c = (Color) Color.parseOOXML( xpp, FormatHandle.colorFONT, bk );
					}
				}
				else if( (eventType == XmlPullParser.END_TAG) && xpp.getName().equals( "font" ) )
				{
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "Font.parseOOXML: " + e.toString() );
		}
		// for incremental styles, font size may not be set
		int size = (sz == null) ? -1 : Font.PointsToFontHeight( new Double( sz ) );
		Font f = new Font( name, 400, size );
		if( c != null )
		{
			f.setOOXMLColor( c );
		}
		if( u != null )
		{
			f.setUnderlined( true );
			if( u instanceof String )
			{
				if( u.equals( "double" ) )
				{
					f.setUnderlineStyle( (byte) 2 );
				}
				else if( u.equals( "singleAccounting" ) )
				{
					f.setUnderlineStyle( (byte) 0x21 );
				}
				else if( u.equals( "doubleAccounting" ) )
				{
					f.setUnderlineStyle( (byte) 0x22 );
				}
			}
		}
		if( b )
		{
			f.setBold( b );
		}
		if( ital )
		{
			f.setItalic( true );
		}
		if( strike )
		{
			f.setStricken( true );
		}
		if( !condense )
		{
			f.setCondensed( false );
		}
		if( !expand )
		{
			f.setExtended( false );
		}
		return f;
	}

	/**
	 * return the appropriate SVG string to define this font
	 *
	 * @return
	 */
	public String getSVG()
	{
		StringBuffer sbf = new StringBuffer( "font-family='" + getFontName() + "'" );
		sbf.append( " font-size='" + getFontHeightInPoints() + "pt'" );
		sbf.append( " font-weight='" + getFontWeight() + "'" );
		// sbf.append(" fill='#222222'"); // TODO: get proper text color
		if( icv != 9 )
		{
			sbf.append( " fill='" + FormatHandle.colorToHexString( FormatHandle.getColor( getColor() ) ) + "'" );
		}
		else
		{
			sbf.append( " fill='" + FormatHandle.colorToHexString( FormatHandle.getColor( 0 ) ) + "'" );
		}
		return sbf.toString();
	}

	/**
	 * EXPERIMENTAL AND MAY NOT BE COMPLETE <br>
	 * Returns true if this font is a Unicode (non-ascii) Charset
	 *
	 * @return
	 */
	private boolean isUnicodeCharSet()
	{
		return ((bCharSet == SHIFTJIS_CHARSET) || (bCharSet == HANGEUL_CHARSET) || (bCharSet == HANGUL_CHARSET) || (bCharSet == GB2312_CHARSET) || (bCharSet == CHINESEBIG5_CHARSET) || (bCharSet == HEBREW_CHARSET) || (bCharSet == ARABIC_CHARSET) || (bCharSet == GREEK_CHARSET) || (bCharSet == TURKISH_CHARSET) || (bCharSet == VIETNAMESE_CHARSET) || (bCharSet == THAI_CHARSET) || (bCharSet == EASTEUROPE_CHARSET) || (bCharSet == RUSSIAN_CHARSET) || (bCharSet == BALTIC_CHARSET));
	}
}