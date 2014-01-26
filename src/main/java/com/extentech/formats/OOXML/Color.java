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

/**
 * color (Data Bar Color)
 * One of the colors associated with the data bar or color scale.  Also font ...
 * Note: the auto attribute is not used in the context of data bars
 * <p/>
 * NOTE: both color, fgColor and bgColor use same CT_COLOR schema:
 * <complexType name="CT_Color">
 * <attribute name="auto" type="xsd:boolean" use="optional"/>
 * <attribute name="indexed" type="xsd:unsignedInt" use="optional"/>
 * <attribute name="rgb" type="ST_UnsignedIntHex" use="optional"/>
 * <attribute name="theme" type="xsd:unsignedInt" use="optional"/>
 * <attribute name="tint" type="xsd:double" use="optional" default="0.0"/>
 * </complexType>
 * <p/>
 * fgColor (Foreground Color)
 * Foreground color of the cell fill pattern. Cell fill patterns operate with two colors: a background color and a
 * foreground color. These combine together to make a patterned cell fill.
 * <p/>
 * <p/>
 * bgColor (Background Color)
 * Background color of the cell fill pattern. Cell fill patterns operate with two colors: a background color and a
 * foreground color. These combine together to make a patterned cell fill.
 */
public class Color implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Color.class );
	private static final long serialVersionUID = 2546003092245407502L;
	private boolean auto = false;
	public static int COLORTYPEINDEXED = 0;
	public static int COLORTYPERGB = 1;
	public static int COLORTYPETHEME = 2;
	private int colortype = -1;
	private String colorval = null;        // value of colortype
	private double tint = 0.0;
	private String element = null;
	private int colorint = -1;        // parsed color (tint + value) translated to ExtenXLS color int value
	private String colorstr = null;    // parsed color (tint + value) translated to HTML color string
	private Theme theme = null;

	public Color()
	{
	}

	public Color( String element, boolean auto, int colortype, String colorval, double tint, short type, Theme t )
	{
		this.element = element;
		this.auto = auto;
		this.colortype = colortype;
		this.colorval = colorval;
		this.tint = tint;
		theme = t;
		parseColor( type );
	}

	public Color( Color c )
	{
		element = c.element;
		auto = c.auto;
		colortype = c.colortype;
		colorval = c.colorval;
		tint = c.tint;
		colorint = c.colorint;
		colorstr = c.colorstr;
		theme = c.theme;
	}

	/**
	 * creates a new Color object based upon a java.awt.Color
	 *
	 * @param clr     Color objeect
	 * @param element "color", "fgColor" or "bgColor"
	 */
	public Color( java.awt.Color c, String element, Theme t )
	{
		colortype = COLORTYPERGB;
		colorval = "FF" + FormatHandle.colorToHexString( c ).substring( 1 );
		this.element = element;
		theme = t;    // ok if it's null
		parseColor( (short) 0 );
	}

	/**
	 * creates a new Color object based upon a web-compliant Hex Color string
	 *
	 * @param clr
	 * @param element "color", "fgColor" or "bgColor"
	 */
	public Color( String clr, String element, Theme t )
	{
		colortype = COLORTYPERGB;
		if( clr.startsWith( "#" ) )
		{
			colorval = "FF" + clr.substring( 1 );
		}
		else if( clr.length() == 6 )
		{
			colorval = "FF" + clr;
		}
		else
		{
			colorval = clr;
		}
		this.element = element;
		theme = t;    // ok if it's null
		parseColor( (short) 0 );
	}

	/**
	 * parses a color element
	 * root= color, fgColor or bgColor
	 *
	 * @param xpp
	 * @return
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, short type, WorkBookHandle bk )
	{
		String element = null;
		boolean auto = false;
		int colortype = -1;
		String colorval = null;
		double tint = 0.0;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "color" ) ||
							tnm.equals( "fgColor" ) ||
							tnm.equals( "bgColor" ) )
					{        // get attributes
						element = tnm;    // save element name
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String val = xpp.getAttributeValue( i );
							if( n.equals( "auto" ) )
							{
								auto = true;
							}
							else if( n.equals( "indexed" ) )
							{
								colortype = Color.COLORTYPEINDEXED;
								colorval = val;
							}
							else if( n.equals( "rgb" ) )
							{
								colortype = COLORTYPERGB;
								colorval = val;
							}
							else if( n.equals( "theme" ) )
							{
								colortype = COLORTYPETHEME;
								colorval = val;
							}
							else if( n.equals( "tint" ) )
							{
								tint = new Double( val );
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( element ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "color.parseOOXML: " + e.toString() );
		}
		Color c = new Color( element, auto, colortype, colorval, tint, type, bk.getWorkBook().getTheme() );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + element );
		if( auto )
		{
			ooxml.append( " auto=\"1\"" );
		}
		else if( colortype == COLORTYPERGB )// rgb
		{
			ooxml.append( " rgb=\"" + colorstr + "\"" );
		}
		else if( colortype == COLORTYPEINDEXED )
		{
			ooxml.append( " indexed=\"" + colorval + "\"" );
		}
		else if( colortype == COLORTYPETHEME ) // theme
		{
			ooxml.append( " theme=\"" + colorval + "\"" );
		}
		if( tint != 0 )
		{
			ooxml.append( " tint=\"" + tint + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	public static String getOOXML( String element, int colortype, int colorval, String colorstr, double tint )
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + element );
		if( colortype == COLORTYPERGB )        // rgb
		{
			ooxml.append( " rgb=\"" + colorstr + "\"" );
		}
		else if( colortype == COLORTYPEINDEXED )    // indexed
		{
			ooxml.append( " indexed=\"" + colorval + "\"" );
		}
		else if( colortype == COLORTYPETHEME )    // theme
		{
			ooxml.append( " theme=\"" + colorval + "\"" );
		}
		if( tint != 0 )
		{
			ooxml.append( " tint=\"" + tint + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Color( this );
	}

	public String toString()
	{
		String ret = "";
		if( colortype == COLORTYPERGB )        // rgb
		{
			ret = " rgb=" + colorstr;
		}
		else if( colortype == COLORTYPEINDEXED )    // indexed
		{
			ret = " indexed=" + colorval;
		}
		else if( colortype == COLORTYPETHEME )    // theme
		{
			ret = " theme=" + colorval;
		}
		if( tint != 0 )
		{
			ret = " tint=" + tint;
		}
		return ret;
	}

	/**
	 * return the translated color int for this OOXML color
	 *
	 * @return
	 */
	public int getColorInt()
	{
		return colorint;
	}

	/**
	 * return the color type for this OOXML color
	 * (indexed= 0, rgb= 1, theme= 2)
	 *
	 * @return
	 */
	public int getColorType()
	{
		return colortype;
	}

	/**
	 * return the translated HTML color string for this OOXML color
	 *
	 * @return
	 */
	public String getColorAsOOXMLRBG()
	{
		return colorstr;
	}

	public java.awt.Color getColorAsColor()
	{
		return FormatHandle.HexStringToColor( getColorAsOOXMLRBG() );
	}

	/**
	 * manually set color int
	 * automatically defaults to indexed color, and looks up color in COLORTABLE
	 *
	 * @param clr
	 */
	public void setColorInt( int clr )
	{
		colorint = clr;
		colortype = 0;    // indexed
		// reset other vars as well
		if( clr > -1 )
		{
			colorstr = FormatHandle.colorToHexString( FormatHandle.COLORTABLE[clr] ).substring( 1 );
		}
		else
		{
			colorstr = null;
		}
		tint = 0.0;
		auto = false;
	}

	/**
	 * sets the color via java.awt.Color
	 *
	 * @param c
	 */
	public void setColor( java.awt.Color c )
	{
		colortype = COLORTYPERGB;
		colorval = "FF" + FormatHandle.colorToHexString( c ).substring( 1 );
		parseColor( (short) 0 );
		tint = 0.0;
		auto = false;
	}

	/**
	 * sets the color via a web-compliant Hex Color String
	 *
	 * @param clr
	 */
	public void setColor( String clr )
	{
		colortype = COLORTYPERGB;
		if( clr.startsWith( "#" ) )
		{
			colorval = "FF" + clr.substring( 1 );
		}
		else if( clr.length() == 6 )
		{
			colorval = "FF" + clr;
		}
		else
		{
			colorval = clr;
		}
		parseColor( (short) 0 );
	}

	/**
	 * static version of parseColor --takes a color value of OOXML type type
	 * <br>COLORTYPERBG, COLORTYPEINDEXED or COLORTYPETHEME
	 * <br>and returns the
	 *
	 * @param val       String OOXML color value, value depends upon colortype
	 * @param colortype int val's colortype, one of:  Color.COLORTYPERGB, Color.COLORTYPEINDEXED, Color.COLORTYPETHEME
	 * @param type      0= font, 1= fill
	 * @return String    HEX-style color string
	 */
	public static String parseColor( String val, int colortype, short type, Theme t )
	{
		Color c = new Color( "", false, colortype, val, 0, type, t );
		return c.colorstr;
	}

	/**
	 * static version of parseColor --takes a color value of OOXML type type
	 * <br>COLORTYPERBG, COLORTYPEINDEXED or COLORTYPETHEME
	 * <br>and returns the indexed color int it represents
	 *
	 * @param val       String OOXML color value, value depends upon colortype
	 * @param colortype int val's colortype, one of:  Color.COLORTYPERGB, Color.COLORTYPEINDEXED, Color.COLORTYPETHEME
	 * @param type      0= font, 1= fill
	 * @return int
	 */
	public static int parseColorInt( String val, int colortype, short type, Theme t )
	{
		Color c = new Color( "", false, colortype, val, 0, type, t );
//		c.parseColor(type);
		return c.colorint;
	}

	/**
	 * simple utility to parse correct colors (int + html string) from OOXML style color
	 *
	 * @param type 0= font, -1= fill
	 */
	private void parseColor( short type )
	{
		try
		{
			if( colortype == COLORTYPERGB )
			{        // rgb - color string
				colorint = FormatHandle.HexStringToColorInt( colorval, type ); // find best match
				colorstr = colorval;
			}
			else if( colortype == COLORTYPEINDEXED )
			{    // indexed (corresponds to either our color int or a custom set of indexed colors
				colorint = Integer.valueOf( colorval );
				if( (colorint == 64) && (type == FormatHandle.colorFONT) ) // means system foreground: Default foreground color. This is the window text color in the sheet display.
				{
					colorstr = FormatHandle.colorToHexString( FormatHandle.getColor( 0 ) );
				}
//					this.colorint= 0; // black
				else
				{
					colorstr = FormatHandle.colorToHexString( FormatHandle.getColor( colorint ) );
				}
			}
			else if( colortype == COLORTYPETHEME )
			{    // theme
				Object[] o = Color.parseThemeColor( colorval, tint, type, theme );
				colorint = (Integer) o[0];
				colorstr = (String) o[1];
			}
		}
		catch( Exception e )
		{
			log.error( "color.parseColor: " + colortype + ":" + colorval + ": " + e.toString() );
		}
	}

	/**
	 * interprets theme colorval and tint
	 * TODO: read in theme colors from theme1.xml
	 *
	 * @param colorval String theme colorval (see SchemeClr and Theme for details)
	 *                 EITHER an Index into the <clrScheme> collection, referencing a particular <sysClr> or 	<srgbClr> value expressed in the Theme part.
	 *                 OR the actual clrScheme string
	 * @param tint     double tint to apply to the base color (0 for none)
	 * @param type     short 0= font, 1= fill
	 * @return
	 */
	public static Object[] parseThemeColor( String colorval, double tint, short type, Theme t )
	{
		/**
		 * TODO: read this in from THEME!!!
		 */
		final String[][] themeColors = {
				{ "dk1", "000000" },
				{ "lt1", "FFFFFF" },
				{ "dk2", "1F497D" },
				{ "lt2", "EEECE1" },
				{ "accent1", "4F81BD" },
				{ "accent1", "C0504D" },
				{ "accent3", "9BBB59" },
				{ "accent4", "8064A2" },
				{ "accent5", "4BACC6" },
				{ "accent6", "F79646" },
				{ "hlink", "0000FF" },
				{ "folHlink", "800080" },
		};

		// TODO: read theme colors from file
		Object[] o = new Object[2];
		int i = 0;
		String clr = "";
		try
		{
			i = Integer.valueOf( colorval );
			clr = t.genericThemeClrs[i];
		}
		catch( Exception e )
		{        // ExtenXLS-created xlsx files will not have theme entries
			for( i = 0; i < themeColors.length; i++ )
			{
				if( themeColors[i][0].equals( colorval ) )
				{
					clr = themeColors[i][1];
					break;
				}
			}
		}
		if( Math.abs( tint ) > .005 )
		{
			java.awt.Color c = applyTint( tint, clr );
			o[0] = FormatHandle.getColorInt( c );
			o[1] = FormatHandle.colorToHexString( c ).substring( 1 );    // avoid #
		}
		else
		{
			o[0] = FormatHandle.HexStringToColorInt( "#" + clr, type );
			o[1] = "#" + clr;
		}
		return o;
	}

	/**
	 * If tint is supplied, then it is applied to the RGB value of the color to determine the final
	 * color applied.
	 * The tint value is stored as a double from -1.0 .. 1.0, where -1.0 means 100% darken and
	 * 1.0 means 100% lighten. Also, 0.0 means no change.
	 * In loading the RGB value, it is converted to HLS where HLS values are (0..HLSMAX), where
	 * HLSMAX is currently 255.
	 */
	private static java.awt.Color applyTint( double tint, String clr )
	{
		final int HLSMAX = 255;
		int r = Integer.parseInt( clr.substring( 0, 2 ), 16 );
		int g = Integer.parseInt( clr.substring( 2, 4 ), 16 );
		int b = Integer.parseInt( clr.substring( 4, 6 ), 16 );

		HSLColor hsl = new HSLColor();
		hsl.initHSLbyRGB( r, g, b );
		double l = hsl.getLuminence();
		if( tint < 0 )        //darken
		{
			l = l * (1 + tint);
		}
		else                // lighten
		{
			l = l * (1 - tint) + (HLSMAX - (HLSMAX * (1.0 - tint)));
		}
		l = Math.round( l );
		hsl.initRGBbyHSL( hsl.getHue(), hsl.getSaturation(), new Double( l ).intValue() );
		return new java.awt.Color( hsl.getRed(), hsl.getGreen(), hsl.getBlue() );

	}
}




	/* HSL stands for hue, saturation, lightness 
	 * Both HSL and HSV describe colors as points in a cylinder 
	 * whose central axis ranges from black at the bottom to white at the top 
	 * with neutral colors between them, 
	 * where angle around the axis corresponds to hue, 
	 * distance from the axis corresponds to saturation, 
	 * and distance along the axis corresponds to lightness, value, or brightness.*/

/*
 * Conversion from RGB to HSL

	Let r, g, b E [0,1] be the red, green, and blue coordinates, respectively, of a color in RGB space.
	Let max be the greatest of r, g, and b, and min the least.

	hue angle h:

	h= 0 													if max=min (for grays)
	 = 60 deg x (g-b)/(max-min) + 360 deg) mod 360 deg		if max=r
	 = 60 deg x (b-r)/(max-min) + 120 deg)					if max=g
	 = 60 deg x (r-g)/(max-min) + 240 deg) 					if max=b

	lightness l= (max-min)/2

	Saturation s

	s= 0													if max=min
	 = (max-min)/2l											if l <= 1/2
	 = (max-min)/(2-2l)										if l > 1/2

	The value of h is generally normalized to lie between 0 and 360, and h = 0 is used when max = min (that is, for grays)
	though the hue has no geometric meaning there, where the saturation s is zero.
	Similarly, the choice of 0 as the value for s when l is equal to 0 or 1 is arbitrary.

	Conversion from HSL to RGB

	Given a color defined by (h, s, l) values in HSL space, with h in the range [0, 360), indicating the angle,
	in degrees of the hue, and with s and l in the range [0, 1], representing the saturation and lightness, respectively,
	a corresponding (r, g, b) triplet in RGB space, with r, g, and b also in range [0, 1], and corresponding to red, green, and blue,
	respectively, can be computed as follows:

	First, if s = 0, then the resulting color is achromatic, or gray. In this special case, r, g, and b all equal l.
	Note that the value of h is ignored, and may be undefined in this situation.

	The following procedure can be used, even when s is zero:

	q= l x (1+s) 					if l < 1/2
	 = l + s - (l x s) 				if l >= 1/2

	p= 2 x l - q

	hk= h / 360 (h normalized in the range 0-1)

	tr= hk + 1/3

	tg= hk

	tb= hk - 1/3

	if tc

 */
class HSLColor
{
	private final static int HSLMAX = 255;
	private final static int RGBMAX = 255;
	private final static int UNDEFINED = 170;

	private int pHue;
	private int pSat;
	private int pLum;
	private int pRed;
	private int pGreen;
	private int pBlue;

	public void initHSLbyRGB( int R, int G, int B )
	{
		// sets Hue, Sat, Lum
		int cMax;
		int cMin;
		int RDelta;
		int GDelta;
		int BDelta;
		int cMinus;
		int cPlus;

		pRed = R;
		pGreen = G;
		pBlue = B;

		//Set Max & MinColor Values
		cMax = iMax( iMax( R, G ), B );
		cMin = iMin( iMin( R, G ), B );

		cMinus = cMax - cMin;
		cPlus = cMax + cMin;

		// Calculate luminescence (lightness)
		pLum = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX);

		if( cMax == cMin )
		{
			// greyscale
			pSat = 0;
			pHue = UNDEFINED;
		}
		else
		{
			// Calculate color saturation
			if( pLum <= (HSLMAX / 2) )
			{
				pSat = (int) (((cMinus * HSLMAX) + 0.5) / cPlus);
			}
			else
			{
				pSat = (int) (((cMinus * HSLMAX) + 0.5) / ((2 * RGBMAX) - cPlus));
			}

			//Calculate hue
			RDelta = (int) ((((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus);
			GDelta = (int) ((((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus);
			BDelta = (int) ((((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus);

			if( cMax == R )
			{
				pHue = BDelta - GDelta;
			}
			else if( cMax == G )
			{
				pHue = (HSLMAX / 3) + RDelta - BDelta;
			}
			else if( cMax == B )
			{
				pHue = ((2 * HSLMAX) / 3) + GDelta - RDelta;
			}

			if( pHue < 0 )
			{
				pHue = pHue + HSLMAX;
			}
		}
	}

	public void initRGBbyHSL( int H, int S, int L )
	{
		int Magic1;
		int Magic2;

		pHue = H;
		pLum = L;
		pSat = S;

		if( S == 0 )
		{ //Greyscale
			pRed = (L * RGBMAX) / HSLMAX; //luminescence: set to range
			pGreen = pRed;
			pBlue = pRed;
		}
		else
		{
			if( L <= (HSLMAX / 2) )
			{
				Magic2 = ((L * (HSLMAX + S)) + (HSLMAX / 2)) / (HSLMAX);
			}
			else
			{
				Magic2 = L + S - ((L * S) + (HSLMAX / 2)) / HSLMAX;
			}
			Magic1 = 2 * L - Magic2;

			//get R, G, B; change units from HSLMAX range to RGBMAX range
			pRed = ((hueToRGB( Magic1, Magic2, H + (HSLMAX / 3) ) * RGBMAX) + (HSLMAX / 2)) / HSLMAX;
			if( pRed > RGBMAX )
			{
				pRed = RGBMAX;
			}

			pGreen = ((hueToRGB( Magic1, Magic2, H ) * RGBMAX) + (HSLMAX / 2)) / HSLMAX;
			if( pGreen > RGBMAX )
			{
				pGreen = RGBMAX;
			}

			pBlue = ((hueToRGB( Magic1, Magic2, H - (HSLMAX / 3) ) * RGBMAX) + (HSLMAX / 2)) / HSLMAX;
			if( pBlue > RGBMAX )
			{
				pBlue = RGBMAX;
			}
		}
	}

	private static int hueToRGB( int mag1, int mag2, int Hue )
	{
		// check the range
		if( Hue < 0 )
		{
			Hue = Hue + HSLMAX;
		}
		else if( Hue > HSLMAX )
		{
			Hue = Hue - HSLMAX;
		}

		if( Hue < (HSLMAX / 6) )
		{
			return (mag1 + ((((mag2 - mag1) * Hue) + (HSLMAX / 12)) / (HSLMAX / 6)));
		}

		if( Hue < (HSLMAX / 2) )
		{
			return mag2;
		}

		if( Hue < ((HSLMAX * 2) / 3) )
		{
			return (mag1 + ((((mag2 - mag1) * (((HSLMAX * 2) / 3) - Hue)) + (HSLMAX / 12)) / (HSLMAX / 6)));
		}

		return mag1;
	}

	private static int iMax( int a, int b )
	{
		if( a > b )
		{
			return a;
		}
		return b;
	}

	private static int iMin( int a, int b )
	{
		if( a < b )
		{
			return a;
		}
		return b;
	}

	public void greyscale()
	{
		initRGBbyHSL( UNDEFINED, 0, pLum );
	}

	// --

	public int getHue()
	{
		return pHue;
	}

	public void setHue( int iToValue )
	{
		while( iToValue < 0 )
		{
			iToValue = HSLMAX + iToValue;
		}
		while( iToValue > HSLMAX )
		{
			iToValue = iToValue - HSLMAX;
		}

		initRGBbyHSL( iToValue, pSat, pLum );
	}

	// --

	public int getSaturation()
	{
		return pSat;
	}

	public void setSaturation( int iToValue )
	{
		if( iToValue < 0 )
		{
			iToValue = 0;
		}
		else if( iToValue > HSLMAX )
		{
			iToValue = HSLMAX;
		}

		initRGBbyHSL( pHue, iToValue, pLum );
	}

	// --

	public int getLuminence()
	{
		return pLum;
	}

	public void setLuminence( int iToValue )
	{
		if( iToValue < 0 )
		{
			iToValue = 0;
		}
		else if( iToValue > HSLMAX )
		{
			iToValue = HSLMAX;
		}

		initRGBbyHSL( pHue, pSat, iToValue );
	}

	// --

	public int getRed()
	{
		return pRed;
	}

	public void setRed( int iNewValue )
	{
		initHSLbyRGB( iNewValue, pGreen, pBlue );
	}

	// --

	public int getGreen()
	{
		return pGreen;
	}

	public void setGreen( int iNewValue )
	{
		initHSLbyRGB( pRed, iNewValue, pBlue );
	}

	// --

	public int getBlue()
	{
		return pBlue;
	}

	public void setBlue( int iNewValue )
	{
		initHSLbyRGB( pRed, pGreen, iNewValue );
	}

	// --

	public void reverseColor()
	{
		setHue( pHue + (HSLMAX / 2) );
	}

	// --

	public void reverseLight()
	{
		setLuminence( HSLMAX - pLum );
	}

	// --

	public void brighten( float fPercent )
	{
		int L;

		if( fPercent == 0 )
		{
			return;
		}

		L = (int) (pLum * fPercent);
		if( L < 0 )
		{
			L = 0;
		}
		if( L > HSLMAX )
		{
			L = HSLMAX;
		}

		setLuminence( L );
	}

	// --
	// --

	public void blend( int R, int G, int B, float fPercent )
	{
		if( fPercent >= 1 )
		{
			initHSLbyRGB( R, G, B );
			return;
		}
		if( fPercent <= 0 )
		{
			return;
		}

		int newR = (int) ((R * fPercent) + (pRed * (1.0 - fPercent)));
		int newG = (int) ((G * fPercent) + (pGreen * (1.0 - fPercent)));
		int newB = (int) ((B * fPercent) + (pBlue * (1.0 - fPercent)));

		initHSLbyRGB( newR, newG, newB );
	}
}

	/*
	Color.getHSBColor(hue, saturation, brightness
	 * The hue parameter is a decimal number between 0.0 and 1.0 which indicates the hue of the color. You'll have to experiment with the hue number to find out what color it represents.
	 * The saturation is a decimal number between 0.0 and 1.0 which indicates how deep the color should be. Supplying a "1" will make the color as deep as possible, and to the other extreme, 
	 * supplying a "0," will take all the color out of the mixture and make it a shade of gray.
	 * The brightness is also a decimal number between 0.0 and 1.0 which obviously indicates how bright the color should be. 
	 * A 1 will make the color as light as possible and a 0 will make it very dark. 
	 */

	/*
			int maxC = Math.max(Math.max(r,g),b);
			int minC = Math.min(Math.min(r,g),b);
			int l = (((maxC + minC)*HLSMAX) + RGBMAX)/(2*RGBMAX);
			int delta= maxC - minC;
			int sum= maxC + minC;
			if (delta != 0) {
					if (l <= (HLSMAX/2))
						s = ( (delta*HLSMAX) + (sum/2) ) / sum;
					else
						s = ( (delta*HLSMAX) + ((2*RGBMAX - sum)/2) ) / (2*RGBMAX - sum);
			
					if (r == maxC)
						h =                ((g - b) * HLSMAX) / (6 * delta);
					else if (g == maxC)
						h = (  HLSMAX/3) + ((b - r) * HLSMAX) / (6 * delta);
					else if (b == maxC)
						h = (2*HLSMAX/3) + ((r - g) * HLSMAX) / (6 * delta);

					if (h < 0)
						h += HLSMAX;
					else if (h >= HLSMAX)
						h -= HLSMAX;
			} else {
					h = 0;
					s = 0;
			}
			

	*/
	
	