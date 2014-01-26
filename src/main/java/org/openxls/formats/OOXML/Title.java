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
import org.openxls.formats.XLS.Font;
import org.openxls.formats.XLS.WorkBook;
import org.openxls.formats.XLS.charts.TextDisp;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Stack;

/**
 * class holds OOXML title Property used to define chart and axis titles
 */
public class Title implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Title.class );
	private static final long serialVersionUID = -3889674575558708481L;
	private Layout layout = null;
	private SpPr sp = null;
	private ChartText chartText = null; // tx
	private TxPr txpr = null; // xPr

	public Title( ChartText ct, TxPr txpr, Layout l, SpPr sp )
	{
		layout = l;
		this.sp = sp;
		chartText = ct;
		this.txpr = txpr;
	}

	/**
	 * for BIFF8 compatibility, create a Title element from the title string
	 *
	 * @param t
	 */
	public Title( String t )
	{
		chartText = new ChartText( t );
		sp = new SpPr( "c" );
		// no spPr
	}

	/**
	 * create an OOXML title from a 2003-v TextDisp record
	 *
	 * @param td
	 */
	public Title( TextDisp td, WorkBook bk )
	{
		P para = new P( td.getFont( bk ), td.toString() );
		chartText = new ChartText( null, para, null );
	}

	public void setLayout( double x, double y )
	{
		layout = new Layout( null, new double[]{ x, y, -1, -1 } );
	}

	/**
	 * parse title OOXML element title
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spPr object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		TxPr txpr = null;
		ChartText ct = null;
		Layout l = null;
		SpPr sp = null;

		/*
		 * TextDisp td= (TextDisp) TextDisp.getPrototype(ObjectLink.TYPE_TITLE,
		 * str, this.getWorkBook()); this.addChartRecord((BiffRec) td); // add
		 * TextDisp title to end of chart recs ... charttitle= td;
		 */
		/**
		 * contains (in Sequence) layout overlay -- not handled yet spPr tx txPr
		 */
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "tx" ) )
					{ // chart text
						lastTag.push( tnm );
						ct = ChartText.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "manualLayout" ) )
					{
						lastTag.push( tnm );
						l = (Layout) Layout.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk );
						// sp.setNS("c");
					}
					else if( tnm.equals( "txPr" ) )
					{
						lastTag.push( tnm );
						txpr = (TxPr) TxPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "title" ) )
					{
						lastTag.pop(); // pop title tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "title.parseOOXML: " + e.toString() );
		}
		Title tt = new Title( ct, txpr, l, sp );
		return tt;
	}

	public Layout getLayout()
	{
		return layout;
	}

	public SpPr getSpPr()
	{
		return sp;
	}

	/**
	 * generate ooxml to define a title
	 *
	 * @return
	 */
	/**
	 * tx chart text layout overlay spPr txPr
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer tooxml = new StringBuffer();
		tooxml.append( "<c:title>" );
		if( chartText != null )
		{
			tooxml.append( chartText.getOOXML() );
		}
		if( layout != null )
		{
			tooxml.append( layout.getOOXML() );
		}
		// TODO: overlay
		if( sp != null )
		{
			tooxml.append( sp.getOOXML() );
		}
		if( txpr != null )
		{
			tooxml.append( txpr.getOOXML() );
		}
		tooxml.append( "</c:title>" );
		return tooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Title( chartText, txpr, layout, sp );
	}

	public String getTitle()
	{
		if( chartText != null )
		{
			return chartText.getTitle();
		}
		return "";
	}

	/**
	 * return the font index for this title
	 *
	 * @param wb
	 * @return
	 */
	public int getFontId( WorkBookHandle wb )
	{
		if( chartText != null )
		{
			return chartText.getFontId( wb );
		}
		return -1;
	}
}

/**
 * chart text element tx
 *
 *
 */

/**
 * contains either strRef -- contains f, strCache rich -- contains bodyPr,
 * lstStyle, p
 */
class ChartText implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( ChartText.class );
	private static final long serialVersionUID = -1175394918747218776L;
	StrRef strref = null;
	P para = null;
	BodyPr bpr = null;

	/**
	 * for BIFF8 compatibility, create an OOXML chartText element from title
	 * string
	 *
	 * @param s
	 */
	public ChartText( String s )
	{
		para = new P( s );
	}

	public ChartText( StrRef s, P para, BodyPr bpr )
	{
		strref = s;
		this.para = para;
		this.bpr = bpr;
	}

	public static ChartText parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		P para = null;
		BodyPr bpr = null;
		StrRef s = null;

		try
		{ // title->tx->rich->bodyPr lstStyle, p->pPr, r
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "p" ) )
					{ // text Paragraph props - part of rich
						lastTag.push( tnm );
						para = (P) P.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "bodyPr" ) )
					{ // body-level Paragraph props -- part of rich
						lastTag.push( tnm );
						bpr = (BodyPr) BodyPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "strRef" ) )
					{
						lastTag.push( tnm );
						s = (StrRef) StrRef.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "tx" ) )
					{
						lastTag.pop(); // pop title tag
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "chartText.parseOOXML: " + e.toString() );
		}
		ChartText ct = new ChartText( s, para, bpr );
		return ct;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer cooxml = new StringBuffer();
		cooxml.append( "<c:tx>" ); // chart text
		if( strref == null )
		{ // it has a rich element
			cooxml.append( "<c:rich>" );
			if( bpr != null )
			{
				cooxml.append( bpr.getOOXML() );
			}
			else
			{
				cooxml.append( "<a:bodyPr/>" );
			}
			cooxml.append( "<a:lstStyle/>" ); // TODO: Handle!!!
			if( para != null )
			{
				cooxml.append( para.getOOXML() ); // text paragraph
			}
			cooxml.append( "</c:rich>" );
		}
		else
		{ // it has a strRef element
			cooxml.append( strref.getOOXML() );
		}
		cooxml.append( "</c:tx>" );
		return cooxml.toString();
	}

	public String getTitle()
	{
		if( para != null )
		{
			return para.getTitle();
		}
		return "";
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ChartText( strref, para, bpr );
	}

	/**
	 * concatenate the 3 levels of text properties and either find an existing
	 * font or add new
	 *
	 * @return
	 */
	public int getFontId( WorkBookHandle wb )
	{
		HashMap<?, ?> textprops = new HashMap<Object, Object>();
		if( bpr != null )
		{
			/*
			 * noAutofit (No AutoFit) §5.1.5.1.2 normAutofit (Normal AutoFit)
			 * §5.1.5.1.3 prstTxWarp (Preset Text Warp) §5.1.11.19 scene3d (3D
			 * Scene Properties) §5.1.4.1.26 sp3d (Apply 3D shape properties)
			 * §5.1.7.12 spAutoFit (Shape AutoFit)
			 */
			// TODO: set any of the above??
		}
		if( para != null )
		{
			textprops = para.getTextProperties();
		}
		/*
		 * altLang (Alternative Language) b (Bold) bool baseline (Baseline) bmk
		 * (Bookmark Link Target) cap (Capitalization) i (Italics) bool kern
		 * (Kerning) kumimoji lang (Language ID) spc (Spacing) strike
		 * (Strikethrough) sz (Font Size) size u (Underline) underline style
		 * 
		 * PLUS-- fill, line, blipFill, cs/ea/latin font (attribute: typeface)
		 * ****
		 */
		int w = 400;
		int u = 0;
		double h = 200; // default
		boolean b = false;
		boolean i = false;
		String face = "Arial";
		if( textprops.get( "b" ) != null )
		{
			b = ("1".equals( textprops.get( "b" ) ));
		}
		if( textprops.get( "i" ) != null )
		{
			i = ("1".equals( textprops.get( "i" ) ));
		}
		if( textprops.get( "latin_typeface" ) != null )
		{
			face = (String) textprops.get( "latin_typeface" );
		}
		// if (textprops.get("u")!=null)
		// u= textprops.get("u").toString();
		Object o = textprops.get( "sz" ); // Whole points are specified in
		// increments of 100 starting with 100
		// being a point size of 1
		if( o != null )
		{
			h = Font.PointsToFontHeight( Integer.parseInt( (String) o ) / 100 );
		}
		Font f = new Font( face, w, new Float( h ).intValue() );
		if( b )
		{
			f.setBold( true );
		}
		if( i )
		{
			f.setItalic( i );
		}
		if( u != 0 )
		{
			f.setUnderlineStyle( (byte) u );
		}
		o = textprops.get( "vertAlign" );
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
		o = textprops.get( "strike" );
		if( o != null )
		{
			f.setStricken( true );
		}
		// f.setFontColor(cl);
		return FormatHandle.addFont( f, wb );
	}

}
