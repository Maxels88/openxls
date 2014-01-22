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
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * p (Text Paragraph)
 * <p/>
 * This element specifies the presence of a paragraph of text within the containing text body. The paragraph is the
 * highest level text separation mechanism within a text body. A paragraph may contain text paragraph properties
 * associated with the paragraph. If no properties are listed then properties specified in the defPPr element are
 * used.
 * <p/>
 * parent:  r, t, txBody, txpr
 * children:  pPr, (r, br or fld), endParaRPr
 */
// TODO: Finish pPr Text Paragraph Properties -- MANY child elements not handled
//TODO: Finish endParaRPr children TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver, 
public class P implements OOXMLElement
{

	private static final long serialVersionUID = 6302706683933521698L;
	private TextRun run = null;
	private PPr ppr = null;
	private EndParaRPr ep = null;

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		PPr ppr = null;
		TextRun run = null;
		EndParaRPr ep = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pPr" ) )
					{        // paragraph-level text props
						lastTag.push( tnm );
						ppr = PPr.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "r" ) ||
							tnm.equals( "br" ) ||
							tnm.equals( "fld" ) )
					{    // text run
						lastTag.push( tnm );
						run = (TextRun) TextRun.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "endParaRPr" ) )
					{
						lastTag.push( tnm );
						ep = EndParaRPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "p" ) )
					{
						lastTag.pop();    // pop this tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "p.parseOOXML: " + e.toString() );

		}
		P pr = new P( ppr, run, ep );
		return pr;
	}

	public P( String s )
	{
		this.run = new TextRun( s );
		this.ppr = new PPr( new DefRPr(), null );
	}

	public P( PPr ppr, TextRun run, EndParaRPr ep )
	{
		this.ppr = ppr;
		this.ep = ep;
		this.run = run;
	}

	public P( P p )
	{
		this.ppr = p.ppr;
		this.ep = p.ep;
		this.run = p.run;
	}

	/**
	 * create a default paragraph property from the specified information
	 *
	 * @param fontFace String font face e.g. "Arial"
	 * @param sz       int size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
	 * @param b        boolean true if bold
	 * @param i        boolean true if italic
	 * @param u        String underline.  One of the following Strings:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
	 *                 dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
	 * @param strike   String strike setting.  One of the following Strings: dblStrike, noStrike or sngStrike  or null if none
	 * @param clr      String fill color in hex form without the #
	 */
	public P( String fontFace, int sz, boolean b, boolean i, String u, String strike, String clr )
	{
		this.ppr = new PPr( fontFace, sz, b, i, u, strike, clr );
	}

	/**
	 * create a default paragraph property from the specified Font and text s
	 *
	 * @param f Font
	 */
	public P( Font f, String s )
	{
		int u = f.getUnderlineStyle();
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
		String strike = (f.getStricken() ? "sngStrike" : "noStrike");
		String clr = FormatHandle.colorToHexString( FormatHandle.getColor( f.getColor() ) ).substring( 1 );
		this.ppr = new PPr( f.getFontName(), (int) f.getFontHeightInPoints() * 100, f.getBold(), f.getItalic(), usty, strike, clr );
		this.run = new TextRun( s );
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:p>" );
		if( ppr != null )
		{
			ooxml.append( ppr.getOOXML() );
		}
		if( run != null )
		{
			ooxml.append( run.getOOXML() );    // 20090526 KSC: order of children was wrong (Kaylan/Rajesh chart error)
		}
		if( ep != null )
		{
			ooxml.append( ep.getOOXML() );
		}
		ooxml.append( "</a:p>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new P( this );
	}

	public String getTitle()
	{
		if( run != null )
		{
			return run.getTitle();
		}
		return "";
	}

	/**
	 * returns a map of all text properties defined for this paragraph
	 *
	 * @return
	 */
	public HashMap<String, String> getTextProperties()
	{
		HashMap<String, String> textprops = new HashMap<String, String>();
		if( ppr != null )
		{
			// algn- left, right, centered, just, distributed
			// defTabSz
			// fontAlgn (Font Alignment)
			// hangingPunct (Hanging Punctuation)  bool
			// indent (Indent)
			// lvl (Level)
			// marL (Left Margin)
			// marR (Right Margin)
			// rtl (Right To Left) bool
				 /* This element contains all default run level text properties for the text runs within a containing paragraph. These
				10 properties are to be used when overriding properties have not been defined within the rPr element.*/
			textprops.putAll( ppr.getTextProperties() );
			textprops.putAll( ppr.getDefaultTextProperties() );
					/*
					 *	altLang (Alternative Language)
					 *  b (Bold)		bool
					 *  baseline (Baseline)
					 *  bmk (Bookmark Link Target)
					 *  cap (Capitalization)
					 *  i (Italics)		bool 
					 *  kern (Kerning)
					 *  kumimoji
					 *  lang (Language ID)
					 *  spc (Spacing)
					 *  strike (Strikethrough) 
					 *  sz (Font Size)	size
					 *  u (Underline)	underline style
					 *  
					 *  PLUS-- fill, line, blipFill, cs/ea/latin font (attribute: typeface) **** 
					 */
		}
		if( run != null )
		{
			textprops.putAll( run.getTextProperties() );
		}
		return textprops;
	}
}

/**
 * pPr (Text Paragraph Properties)
 * <p/>
 * This element contains all paragraph level text properties for the containing paragraph. These paragraph
 * properties should override any and all conflicting properties that are associated with the paragraph in question
 * <p/>
 * parent: p, fld
 * children: many
 * attributes: many
 */
// TODO: Handle child elements: lnSpc, spcBef, spcAft, TEXTBULLETCOLOR, TEXTBULLETSIZE, TEXTBULLET, tabLst
class PPr implements OOXMLElement
{

	private static final long serialVersionUID = -6909210948618654877L;
	private DefRPr dp;
	private HashMap<String, String> attrs;

	public PPr( DefRPr dp, HashMap<String, String> attrs )
	{
		this.dp = dp;
		this.attrs = attrs;
	}

	public PPr( PPr p )
	{
		this.dp = p.dp;
		this.attrs = p.attrs;
	}

	/**
	 * create a default paragraph property from the specified information
	 *
	 * @param fontFace String font face e.g. "Arial"
	 * @param sz       int size in 100 pts (e.g. font size 12.5 pts,, sz= 1250)
	 * @param b        boolean true if bold
	 * @param i        boolean true if italic
	 * @param u        String underline.  One of the following Strings:  dash, dashHeavy, dashLong, dashLongHeavy, dbl, dotDash, dotDashHeavy, dotDotDash, dotDotDashHeavy, dotted
	 *                 dottedHeavy, heavy, none, sng, wavy, wavyDbl, wavyHeavy, words (underline only words not spaces)
	 * @param strike   String strike setting.  One of the following Strings: dblStrike, noStrike or sngStrike  or null if none
	 * @param clr      String fill color in hex form without the #
	 */
	public PPr( String fontFace, int sz, boolean b, boolean i, String u, String strike, String clr )
	{
		this.dp = new DefRPr( fontFace, sz, b, i, u, strike, clr );
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:pPr" );
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = i.next();
				String val = attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		ooxml.append( ">" );
		if( dp != null )
		{
			ooxml.append( dp.getOOXML() );
		}
		ooxml.append( "</a:pPr>" );
		return ooxml.toString();
	}

	public static PPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		DefRPr dp = null;
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pPr" ) )
					{        // t element of text run -- the title string we are interested in
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{    // align, defTabSz, fontAlgn, hangingPunct, indent, lvl, rtl, marL, marR ...
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "defRPr" ) )
					{        // default text properties
						lastTag.push( tnm );
						dp = (DefRPr) DefRPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "pPr" ) )
					{
						lastTag.pop();    // pop this element
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "paraText.parseOOXML: " + e.toString() );
		}
		PPr pt = new PPr( dp, attrs );
		return pt;
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PPr( this );
	}

	/**
	 * return the default run properties for this para
	 *
	 * @return
	 */
	public HashMap<String, String> getDefaultTextProperties()
	{
		return dp.getTextProperties();
	}

	/**
	 * return the text properties defined by PPr:
	 * algn- left, right, centered, just, distributed
	 * defTabSz
	 * fontAlgn (Font Alignment)
	 * hangingPunct (Hanging Punctuation)  bool
	 * indent (Indent)
	 * lvl (Level)
	 * marL (Left Margin)
	 * marR (Right Margin)
	 * rtl (Right To Left) bool
	 *
	 * @return
	 */
	public HashMap<String, String> getTextProperties()
	{
		if( attrs != null )
		{
			return attrs;
		}
		return new HashMap<String, String>();
	}
}

/**
 * endParaRPr (End Paragraph Run Properties)
 * <p/>
 * This element specifies the text run properties that are to be used if another run is inserted after the last run
 * specified. This effectively saves the run property state so that it may be applied when the user enters additional
 * text. If this element is omitted, then the application may determine which default properties to apply. It is
 * recommended that this element be specified at the end of the list of text runs within the paragraph so that an
 * orderly list is maintained.
 * <p/>
 * parent: p
 * children: ln, FillGroup, EffectGroup, highlight, TEXTUNDERLINE, TEXTUNDERLINEFILL, latin, ea, cs, sym, hlinkClick, hlinkMouseOver,
 */
// TODO: Finish children TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver, 
class EndParaRPr implements OOXMLElement
{

	private static final long serialVersionUID = -7094231887468090281L;
	private HashMap<String, String> attrs = null;
	private Ln l;
	private FillGroup fill;
	private EffectPropsGroup effect;
	private String latin, ea, cs;    // really children but only have 1 attribute and no children

	public EndParaRPr( HashMap<String, String> attrs, Ln l, FillGroup fill, EffectPropsGroup effect, String latin, String ea, String cs )
	{
		this.attrs = attrs;
		this.l = l;
		this.fill = fill;
		this.effect = effect;
		this.latin = latin;
		this.ea = ea;
		this.cs = cs;
	}

	public EndParaRPr( EndParaRPr ep )
	{
		this.attrs = ep.attrs;
		this.l = ep.l;
		this.fill = ep.fill;
		this.effect = ep.effect;
		this.latin = ep.latin;
		this.ea = ep.ea;
		this.cs = ep.cs;
	}

	public static EndParaRPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		FillGroup fill = null;
		EffectPropsGroup effect = null;
		Ln l = null;
		String latin = null, ea = null, cs = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "endParaRPr" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "ln" ) )
					{
						lastTag.push( tnm );
						l = (Ln) Ln.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "solidFill" ) ||
							tnm.equals( "noFill" ) ||
							tnm.equals( "gradFill" ) ||
							tnm.equals( "grpFill" ) ||
							tnm.equals( "pattFill" ) ||
							tnm.equals( "blipFill" ) )
					{
						lastTag.push( tnm );
						fill = (FillGroup) FillGroup.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "effectLst" ) || tnm.equals( "effectDag" ) )
					{
						lastTag.push( tnm );
						effect = (EffectPropsGroup) EffectPropsGroup.parseOOXML( xpp, lastTag );
						// TODO: Eventually these will be objects
					}
					else if( tnm.equals( "latin" ) )
					{
						latin = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "ea" ) )
					{
						ea = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "cs" ) )
					{
						cs = xpp.getAttributeValue( 0 );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "endParaRPr" ) )
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
			Logger.logErr( "endParaRPr.parseOOXML: " + e.toString() );
		}
		EndParaRPr oe = new EndParaRPr( attrs, l, fill, effect, latin, ea, cs );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:endParaRPr" );
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = i.next();
				String val = attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		ooxml.append( ">" );
		if( l != null )
		{
			ooxml.append( l.getOOXML() );
		}
		if( fill != null )
		{
			ooxml.append( fill.getOOXML() );        // group fill
		}
		if( latin != null )
		{
			ooxml.append( "<a:latin typeface=\"" + latin + "\"/>" );
		}
		if( ea != null )
		{
			ooxml.append( "<a:ea typeface=\"" + ea + "\"/>" );
		}
		if( cs != null )
		{
			ooxml.append( "<a:cs typeface=\"" + cs + "\"/>" );
		}
		// sym
		ooxml.append( "</a:endParaRPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new EndParaRPr( this );
	}
}

