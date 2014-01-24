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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * textRun group, either r (regular text), br (line break) or fld (text Field)
 * parent:  p
 * children: either r, br or fld
 */
//TODO: Finish rPr children highlight TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver, 
public class TextRun implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( TextRun.class );
	private static final long serialVersionUID = -6224636879471246452L;
	private r run = null;
	private Br brk = null;
	private Fld f = null;

	public TextRun( r run, Br brk, Fld f )
	{
		this.run = run;
		this.brk = brk;
		this.f = f;
	}

	public TextRun( TextRun r )
	{
		run = r.run;
		brk = r.brk;
		f = r.f;
	}

	/**
	 * create a new regular text text run (OOXML element r)
	 *
	 * @param s
	 */
	public TextRun( String s )
	{
		run = new r( s, null );
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		r run = null;
		Br brk = null;
		Fld f = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "r" ) )
					{
						lastTag.push( tnm );
						run = r.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "br" ) )
					{
						lastTag.push( tnm );
						brk = Br.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "fld" ) )
					{
						lastTag.push( tnm );
						f = Fld.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{    // shouldn't get here
					lastTag.pop();
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "textRun.parseOOXML: " + e.toString() );
		}
		TextRun oe = new TextRun( run, brk, f );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( run != null )
		{
			ooxml.append( run.getOOXML() );
		}
		else if( brk != null )
		{
			ooxml.append( brk.getOOXML() );
		}
		else if( f != null )
		{
			ooxml.append( f.getOOXML() );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new TextRun( this );
	}

	public String getTitle()
	{
		if( run != null )
		{
			return run.getTitle();
		}
		if( f != null )
		{
			return f.getTitle();
		}
		return null;
	}

	/**
	 * return the text properties for this text run
	 *
	 * @return
	 */
	public HashMap<String, String> getTextProperties()
	{
		if( run != null )
		{
			return run.getTextProperties();
		}
		return new HashMap<>();
	}
}

/**
 * OOXML element r, text run, sub-element of p (paragraph)
 * <p/>
 * children:  rPr, t (actual text string)
 */
class r implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( r.class );
	private static final long serialVersionUID = 863254651451294443L;
	private String t = "";        // t element just contains string
	private RPr rp = null;

	public r( String title, RPr rp )
	{
		this.rp = rp;
		t = title;
	}

	public r( r run )
	{
		rp = run.rp;
		t = run.t;
	}

	@Override
	public String getOOXML()
	{
		if( (t == null) || t.equals( "" ) )
		{
			return "";
		}

		t = com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii( t ).toString();
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:r>" );                    // text run
		if( rp != null )
		{
			ooxml.append( rp.getOOXML() );
		}
		ooxml.append( "<a:t>" + t + "</a:t>" );
		ooxml.append( "</a:r>" );
		return ooxml.toString();
	}

	public static r parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String t = "";
		RPr rp = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "t" ) )
					{        // t element of text run -- the title string we are interested in
						t = com.extentech.formats.XLS.OOXMLAdapter.getNextText( xpp );
					}
					else if( tnm.equals( "rPr" ) )
					{    // text run properties
						lastTag.push( tnm );
						rp = RPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "r" ) )
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
			log.error( "r.parseOOXML: " + e.toString() );
		}
		r run = new r( t, rp );
		return run;
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new r( this );
	}

	public String getTitle()
	{
		return t;
	}

	public HashMap<String, String> getTextProperties()
	{
		if( rp != null )
		{
			return rp.getTextProperties();
		}
		return new HashMap<>();
	}
}

/**
 * OOXML element br, vertical break, sub-element of p (paragraph)
 * <p/>
 * children:  rPr
 */
class Br implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Br.class );
	private static final long serialVersionUID = -1724086871866480013L;
	private RPr rp;

	public Br( RPr rp )
	{
		this.rp = rp;
	}

	public Br( Br b )
	{
		rp = b.rp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:br>" );
		if( rp != null )
		{
			ooxml.append( rp.getOOXML() );
		}
		ooxml.append( "</a:br>" );
		return ooxml.toString();
	}

	public static Br parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		RPr rp = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "rPr" ) )
					{    // text run properties
						lastTag.push( tnm );
						rp = RPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "br" ) )
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
			log.error( "br.parseOOXML: " + e.toString() );
		}
		Br b = new Br( rp );
		return b;
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Br( this );
	}
}

/**
 * OOXML element fld, text field, sub-element of p (paragraph)
 * <p/>
 * children:  pPr, rPr, t (actual text string)
 */
class Fld implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Fld.class );
	private static final long serialVersionUID = -7060602732912595402L;
	private String t;        // t element just contains string
	private RPr rp;
	private PPr p;
	String id;
	String type;

	public Fld( String id, String type, String title, RPr rp, PPr p )
	{
		this.id = id;
		this.type = type;
		this.rp = rp;
		t = title;
		this.p = p;
	}

	public Fld( Fld f )
	{
		id = f.id;
		type = f.type;
		rp = f.rp;
		t = f.t;
		p = f.p;
	}

	@Override
	public String getOOXML()
	{
		if( (t == null) || t.equals( "" ) )
		{
			return "";
		}

		t = com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii( t ).toString();
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:fld" );                    // text field
		ooxml.append( " id=\"" + id + "\"" );
		if( type != null )
		{
			ooxml.append( " type=\"" + type + "\"" );
		}
		ooxml.append( ">" );
		if( rp != null )
		{
			ooxml.append( rp.getOOXML() );
		}
		if( p != null )
		{
			ooxml.append( p.getOOXML() );
		}
		ooxml.append( "<a:t>" + t + "</a:t>" );
		ooxml.append( "</a:fld>" );
		return ooxml.toString();
	}

	public static Fld parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String t = "";
		String id = "";
		String type = "";
		PPr p = null;
		RPr rp = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "fld" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "id" ) )
							{
								id = xpp.getAttributeValue( i );
							}
							else if( n.equals( "type" ) )
							{
								type = xpp.getAttributeValue( i );
							}
						}
					}
					else if( tnm.equals( "t" ) )
					{        // t element -- the title string we are interested in
						t = com.extentech.formats.XLS.OOXMLAdapter.getNextText( xpp );
					}
					else if( tnm.equals( "rPr" ) )
					{    // text run properties
						lastTag.push( tnm );
						rp = RPr.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "pPr" ) )
					{    // text field properties
						lastTag.push( tnm );
						p = PPr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fld" ) )
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
			log.error( "textRun.parseOOXML: " + e.toString() );
		}
		Fld f = new Fld( id, type, t, rp, p );
		return f;
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Fld( this );
	}

	public String getTitle()
	{
		return t;
	}

}

class RPr implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( RPr.class );
	private static final long serialVersionUID = 228716184734751439L;
	private HashMap<String, String> attrs;
	private Ln l;
	private FillGroup fill;
	private EffectPropsGroup effect;
	private String latin;
	private String ea;
	private String cs;

	public RPr( HashMap<String, String> attrs, Ln l, FillGroup fill, EffectPropsGroup effect, String latin, String ea, String cs )
	{
		this.attrs = attrs;
		this.l = l;
		this.fill = fill;
		this.effect = effect;
		this.latin = latin;
		this.ea = ea;
		this.cs = cs;
	}

	public RPr( RPr rp )
	{
		attrs = rp.attrs;
		l = rp.l;
		fill = rp.fill;
		effect = rp.effect;
		latin = rp.latin;
		ea = rp.ea;
		cs = rp.cs;
	}

	public static RPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<>();
		Ln l = null;
		FillGroup fill = null;
		EffectPropsGroup effect = null;
		String latin = null;
		String ea = null;
		String cs = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "rPr" ) )
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
					if( endTag.equals( "rPr" ) )
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
			log.error( "rPr.parseOOXML: " + e.toString() );
		}
		RPr rp = new RPr( attrs, l, fill, effect, latin, ea, cs );
		return rp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:rPr" );
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
		if( effect != null )
		{
			ooxml.append( effect.getOOXML() );    // group effect
		}
		// highlight
		// TEXTUNDERLINELINE
		// TEXTUNDERLINEFILL
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
		// hlinkClick
		// hlinkMouseOver
		ooxml.append( "</a:rPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new RPr( this );
	}

	/**
	 * return the text properties for this text run
	 *
	 * @return
	 */
	public HashMap<String, String> getTextProperties()
	{
		HashMap<String, String> textprops = new HashMap<>();
		textprops.putAll( attrs );
		if( latin != null )
		{
			textprops.put( "latin_typeface", latin );
		}
		if( ea != null )
		{
			textprops.put( "ea_typeface", ea );
		}
		if( cs != null )
		{
			textprops.put( "cs_typeface", cs );
		}
		// TODO: Fill, line ...
		return textprops;
	}
}
