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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * FillGroup -- holds ONE OF the Fill-type OOXML/DrawingML Elements: blipFill
 * gradFill grpFill pattFill solidFill noFill
 * <p/>
 * parents: bgPr, spPr, fmtScheme, endParaRPr, rPr ...
 */
// TODO: Handle blip element children
// TODO: Handle gradFill element shade properties
public class FillGroup implements OOXMLElement
{
	private static final long serialVersionUID = 8320871291479597945L;
	private BlipFill bf;
	private GradFill gpf;
	private GrpFill grpf;
	private PattFill pf;
	private SolidFill sf;

	public FillGroup( BlipFill bf, GradFill gpf, GrpFill grpf, PattFill pf, SolidFill sf )
	{
		this.bf = bf;
		this.gpf = gpf;
		this.grpf = grpf;
		this.pf = pf;
		this.sf = sf;
	}

	public FillGroup( FillGroup f )
	{
		this.bf = f.bf;
		this.gpf = f.gpf;
		this.grpf = f.grpf;
		this.pf = f.pf;
		this.sf = f.sf;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		BlipFill bf = null;
		GradFill gpf = null;
		GrpFill grpf = null;
		PattFill pf = null;
		SolidFill sf = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "solidFill" ) )
					{
						lastTag.push( tnm );
						sf = (SolidFill) SolidFill.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					else if( tnm.equals( "noFill" ) )
					{
						// do nothing
					}
					else if( tnm.equals( "gradFill" ) )
					{
						lastTag.push( tnm );
						gpf = (GradFill) GradFill.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					else if( tnm.equals( "grpFill" ) )
					{
						lastTag.push( tnm );
						grpf = (GrpFill) GrpFill.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					else if( tnm.equals( "pattFill" ) )
					{
						lastTag.push( tnm );
						pf = (PattFill) PattFill.parseOOXML( xpp, lastTag, bk );
						lastTag.pop();
						break;
					}
					else if( tnm.equals( "blipFill" ) )
					{
						lastTag.push( tnm );
						bf = (BlipFill) BlipFill.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{ // shouldn't get here
					lastTag.pop();
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "FillGroup.parseOOXML: " + e.toString() );
		}
		FillGroup f = new FillGroup( bf, gpf, grpf, pf, sf );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		// CHOICE OF fill
		if( sf != null )
		{
			ooxml.append( sf.getOOXML() );
		}
		else if( gpf != null )
		{
			ooxml.append( gpf.getOOXML() );
		}
		else if( grpf != null )
		{
			ooxml.append( grpf.getOOXML() );
		}
		else if( bf != null )
		{
			ooxml.append( bf.getOOXML() );
		}
		else if( pf != null )
		{
			ooxml.append( pf.getOOXML() );
		}
		else
		// no fill
		{
			ooxml.append( "<a:noFill/>" );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FillGroup( this );
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( bf != null )
		{
			return bf.getEmbed();
		}
		return null;
	}

	/**
	 * return the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( bf != null )
		{
			return bf.getLink();
		}
		return null;
	}

	/**
	 * set the embed attribute for this blip (the id for the embedded picture)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( bf == null )
		{
			bf = new BlipFill();
		}
		bf.setEmbed( embed );
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( bf == null )
		{
			bf = new BlipFill();
		}
		bf.setLink( link );
	}

	/**
	 * set the color for this fill
	 * NOTE: at this time only solid fills may be specified
	 *
	 * @param clr
	 */
	public void setColor( String clr )
	{
		if( bf != null )
		{
			return;
		}
		gpf = null;
		grpf = null;
		pf = null;
		sf = new SolidFill();
		sf.setColor( clr );
	}

	public int getColor()
	{
		if( sf != null )
		{
			return sf.getColor();
		}
		return -1;
	}
}

/**
 * gradFill (Gradient Fill)
 * <p/>
 * This element defines a gradient fill.
 * <p/>
 * parents: many children: gsLst, SHADEPROPERTIES, tileRect
 */
// TODO: finish SHADEPROPERTIES
class GradFill implements OOXMLElement
{

	private static final long serialVersionUID = 8965776942160065286L;
	private GsLst g;
	// private ShadeProps sp= null;
	private TileRect tr;
	private HashMap<String, String> attrs;

	public GradFill( GsLst g, /* ShadeProps sp, */TileRect tr, HashMap<String, String> attrs )
	{
		this.g = g;
		// this.sp= sp;
		this.tr = tr;
		this.attrs = attrs;
	}

	public GradFill( GradFill gf )
	{
		this.g = gf.g;
		// this.sp= gf.sp;
		this.tr = gf.tr;
		this.attrs = gf.attrs;
	}

	public static GradFill parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		GsLst g = null;
		// ShadeProps sp= null;
		TileRect tr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gradFill" ) )
					{ // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "tileRect" ) )
					{
						lastTag.push( tnm );
						tr = (TileRect) TileRect.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "gsLst" ) )
					{
						lastTag.push( tnm );
						g = (GsLst) GsLst.parseOOXML( xpp, lastTag, bk );
						/*
						 * } else if (tnm.equals("lin") || tnm.equals("path")) {
						 * lastTag.push(tnm); sp= (shadeProps)
						 * shadeProps.parseOOXML(xpp, lastTag).clone();
						 */
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "gradFill" ) )
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
			Logger.logErr( "gradFill.parseOOXML: " + e.toString() );
		}
		GradFill gf = new GradFill( g, /* sp, */tr, attrs );
		return gf;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:gradFill" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( g != null )
		{
			ooxml.append( g.getOOXML() );
		}
		// if (sp!=null) ooxml.append(sp.getOOXML());
		if( tr != null )
		{
			ooxml.append( tr.getOOXML() );
		}
		ooxml.append( "</a:gradFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GradFill( this );
	}

}

/**
 * blipFill (Picture Fill)
 * <p/>
 * This element specifies the type of picture fill that the picture object will
 * have. Because a picture has a picture fill already by default, it is possible
 * to have two fills specified for a picture object. An example of this is shown
 * below.
 * <p/>
 * parents: many children: blip, srcRect, FillModeProperties
 */
class BlipFill implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2030570462677450734L;
	private HashMap<String, String> attrs = null;
	private Blip blip = null;
	private SrcRect srcRect;
	private FillMode fillMode;
	private String ns = "xdr";    //	default namespace

	public BlipFill()
	{    // no-parameter constructor, set common defaults
		srcRect = new SrcRect();
		fillMode = new FillMode( null, false );
	}

	public BlipFill( String ns, HashMap<String, String> attrs, Blip b, SrcRect s, FillMode f )
	{
		this.ns = ns;
		this.attrs = attrs;
		this.blip = b;
		this.srcRect = s;
		this.fillMode = f;
	}

	public BlipFill( BlipFill bf )
	{
		this.ns = bf.ns;
		this.attrs = bf.attrs;
		this.blip = bf.blip;
		this.srcRect = bf.srcRect;
		this.fillMode = bf.fillMode;
	}

	public static BlipFill parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		Blip b = null;
		SrcRect s = null;
		FillMode f = null;
		String ns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "blipFill" ) )
					{ // get attributes
						ns = xpp.getPrefix();
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "blip" ) )
					{
						lastTag.push( tnm );
						b = (Blip) Blip.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "srcRect" ) )
					{
						lastTag.push( tnm );
						s = (SrcRect) SrcRect.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "stretch" ) || tnm.equals( "tile" ) )
					{
						lastTag.push( tnm );
						f = (FillMode) FillMode.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "blipFill" ) )
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
			Logger.logErr( "blipFill.parseOOXML: " + e.toString() );
		}
		BlipFill bf = new BlipFill( ns, attrs, b, s, f );
		return bf;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( ns == null )
		{
			Logger.logErr( "Error: BlipFill Namespace is null" );
		}
		ooxml.append( "<" + ns + ":blipFill" );
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = (String) i.next();
				String val = (String) attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		ooxml.append( ">" );
		if( blip != null )
		{
			ooxml.append( blip.getOOXML() );
		}
		if( srcRect != null )
		{
			ooxml.append( srcRect.getOOXML() );
		}
		if( fillMode != null )
		{
			ooxml.append( fillMode.getOOXML() );
		}
		ooxml.append( "</" + ns + ":blipFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new BlipFill( this );
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( blip != null )
		{
			return blip.getEmbed();
		}
		return null;
	}

	/**
	 * return the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( blip != null )
		{
			return blip.getLink();
		}
		return null;
	}

	/**
	 * set the embed attribute for this blip (the id for the embedded picture)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		if( blip == null )
		{
			blip = new Blip();
		}
		blip.setEmbed( embed );
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( blip == null )
		{
			blip = new Blip();
		}
		blip.setLink( link );
	}
}

/**
 * solidFill (Solid Fill) This element specifies a solid color fill. The shape
 * is filled entirely with the specified color
 * <p/>
 * parents: many children: COLORCHOICE
 */
class SolidFill implements OOXMLElement
{

	private static final long serialVersionUID = 3341509200573989744L;
	private ColorChoice color;

	public SolidFill()
	{ // no-param constructor, set up common defaults
		this.color = new ColorChoice( null, new SrgbClr( "000000", null ), null, null, null );
	}

	public SolidFill( ColorChoice c )
	{
		this.color = c;
	}

	public SolidFill( SolidFill s )
	{
		this.color = s.color;
	}

	/**
	 * creates a solid fill of a specific color
	 *
	 * @param clr hex color string without #
	 */
	public SolidFill( String clr )
	{
		if( clr == null )
		{
			clr = "FFFFFF";
		}
		this.color = new ColorChoice( null, new SrgbClr( clr, null ), null, null, null );
	}

	public static SolidFill parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "hslClr" ) || tnm.equals( "prstClr" ) || tnm.equals( "schemeClr" ) || tnm.equals( "scrgbClr" ) || tnm.equals(
							"srgbClr" ) || tnm.equals( "sysClr" ) )
					{ // get attributes
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk ).cloneElement();

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "solidFill" ) )
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
			Logger.logErr( "solidFill.parseOOXML: " + e.toString() );
		}
		SolidFill s = new SolidFill( c );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:solidFill>" );
		if( this.color != null )
		{
			ooxml.append( this.color.getOOXML() );
		}
		ooxml.append( "</a:solidFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SolidFill( this );
	}

	/**
	 * set the color for this solid fill in hex (html) string format
	 *
	 * @param clr
	 */
	public void setColor( String clr )
	{
		this.color = new ColorChoice( null, new SrgbClr( clr, null ), null, null, null );
	}

	public int getColor()
	{
		return color.getColor();
	}
}

/**
 * pattFill (Pattern Fill) This element specifies a pattern fill. A repeated
 * pattern is used to fill the object.
 * <p/>
 * parent: many children: bgClr, fgClr
 */
class PattFill implements OOXMLElement
{

	private static final long serialVersionUID = -1052627959661249692L;
	private String prst;
	private BgClr bg;
	private FgClr fg;

	public PattFill( String prst, BgClr bg, FgClr fg )
	{
		this.prst = prst;
		this.bg = bg;
		this.fg = fg;
	}

	public PattFill( PattFill p )
	{
		this.prst = p.prst;
		this.bg = p.bg;
		this.fg = p.fg;
	}

	public static PattFill parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String prst = null;
		BgClr bg = null;
		FgClr fg = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pattFill" ) )
					{ // get attributes
						prst = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "bgClr" ) )
					{
						lastTag.push( tnm );
						bg = (BgClr) BgClr.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "fgClr" ) )
					{
						lastTag.push( tnm );
						fg = (FgClr) FgClr.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "pattFill" ) )
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
			Logger.logErr( "pattFill.parseOOXML: " + e.toString() );
		}
		PattFill p = new PattFill( prst, bg, fg );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:pattFill prst=\"" + prst + "\">" );
		if( fg != null )
		{
			ooxml.append( fg.getOOXML() );
		}
		if( bg != null )
		{
			ooxml.append( bg.getOOXML() );
		}
		ooxml.append( "</a:pattFill>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PattFill( this );
	}
}

/**
 * grpFill (Group Fill) This element specifies a group fill. When specified,
 * this setting indicates that the parent element is part of a group and should
 * inherit the fill properties of the group.
 * <p/>
 * parent: many children: CT_GROUPFILL: ??? contains nothing?
 */
class GrpFill implements OOXMLElement
{
	private static final long serialVersionUID = 2388879629485740996L;

	public static GrpFill parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "grpFill" ) )
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
			Logger.logErr( "grpFill.parseOOXML: " + e.toString() );
		}
		GrpFill g = new GrpFill();
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		// TODO: no attributes or children?
		ooxml.append( "<a:grpFill/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GrpFill();
	}
}

/**
 * bgClr background color
 * <p/>
 * parent: pattFill
 */
class BgClr implements OOXMLElement
{

	private static final long serialVersionUID = -879409152334931909L;
	private ColorChoice colorChoice;

	public BgClr( ColorChoice c )
	{
		this.colorChoice = c;
	}

	public BgClr( BgClr s )
	{
		this.colorChoice = s.colorChoice;
	}

	public static BgClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "hslClr" ) || tnm.equals( "prstClr" ) || tnm.equals( "schemeClr" ) || tnm.equals( "scrgbClr" ) || tnm.equals(
							"srgbClr" ) || tnm.equals( "sysClr" ) )
					{ // get attributes
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "bgClr" ) )
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
			Logger.logErr( "bgClr.parseOOXML: " + e.toString() );
		}
		BgClr s = new BgClr( c );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:bgClr>" );
		if( this.colorChoice != null )
		{
			ooxml.append( this.colorChoice.getOOXML() );
		}
		ooxml.append( "</a:bgClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new BgClr( this );
	}
}

/**
 * fgClr Foreground Color
 * <p/>
 * parent: pattFill
 */
class FgClr implements OOXMLElement
{
	private static final long serialVersionUID = 6836994790529289731L;
	private ColorChoice colorChoice;

	public FgClr( ColorChoice c )
	{
		this.colorChoice = c;
	}

	public FgClr( FgClr s )
	{
		this.colorChoice = s.colorChoice;
	}

	public static FgClr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "hslClr" ) || tnm.equals( "prstClr" ) || tnm.equals( "schemeClr" ) || tnm.equals( "scrgbClr" ) || tnm.equals(
							"srgbClr" ) || tnm.equals( "sysClr" ) )
					{ // get attributes
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "fgClr" ) )
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
			Logger.logErr( "fgClr.parseOOXML: " + e.toString() );
		}
		FgClr s = new FgClr( c );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:fgClr>" );
		if( this.colorChoice != null )
		{
			ooxml.append( this.colorChoice.getOOXML() );
		}
		ooxml.append( "</a:fgClr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FgClr( this );
	}
}

/**
 * blip (Blip) This element specifies the existence of an image (binary large
 * image or picture) and contains a reference to the image data.
 * <p/>
 * parent: blipFill, buFill children: MANY
 */
// TODO: HANDLE THE MANY CHILDREN
class Blip implements OOXMLElement
{

	private static final long serialVersionUID = 5188967633123620513L;
	private HashMap<String, String> attrs = new HashMap<String, String>();

	public Blip()
	{
	}

	public Blip( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Blip( Blip b )
	{
		this.attrs = b.attrs;
	}

	public static Blip parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "blip" ) )
					{ // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							if( xpp.getAttributePrefix( i ) != null )
							{
								attrs.put( xpp.getAttributePrefix( i ) + ":" + xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
							}
							else
							{
								attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
							}
						}
						// } else if (tnm.equals("CHILDELEMENT")) {
						// lastTag.push(tnm);
						// layout = (layout) layout.parseOOXML(xpp,
						// lastTag).clone();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "blip" ) )
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
			Logger.logErr( "blip.parseOOXML: " + e.toString() );
		}
		Blip b = new Blip( attrs );
		return b;
	}

	// TODO: cstate= "print"
	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		// TODO: HANDLE CHILDREN
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Blip( this );
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( attrs != null && attrs.get( "r:embed" ) != null )
		{
			return (String) attrs.get( "r:embed" );
		}
		return null;
	}

	/**
	 * return the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		if( attrs != null && attrs.get( "link" ) != null )
		{
			return (String) attrs.get( "link" );
		}
		return null;
	}

	/**
	 * set the embed attribute for this blip (the id for the embedded picture)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		attrs.put( "r:embed", embed );
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		attrs.put( "link", link );
	}
}

/**
 * tileRect (Tile Rectangle) This element specifies a rectangular region of the
 * shape to which the gradient is applied. This region is then tiled across the
 * remaining area of the shape to complete the fill. The tile rectangle is
 * defined by percentage offsets from the sides of the shape's bounding box.
 * parent: gradFill children: none
 */
class TileRect implements OOXMLElement
{

	private static final long serialVersionUID = 5380575948049571420L;
	private HashMap<String, String> attrs;

	public TileRect( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public TileRect( TileRect t )
	{
		this.attrs = t.attrs;
	}

	public static TileRect parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "tileRect" ) )
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
					if( endTag.equals( "tileRect" ) )
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
			Logger.logErr( "tileRect.parseOOXML: " + e.toString() );
		}
		TileRect t = new TileRect( attrs );
		return t;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:tileRect" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new TileRect( this );
	}
}

/**
 * srcRect (Source Rectangle) This element specifies the portion of the blip
 * used for the fill
 * <p/>
 * parent: blipFill children: NONE
 */
class SrcRect implements OOXMLElement
{

	private static final long serialVersionUID = -6407800173040857433L;
	private HashMap<String, String> attrs = null;

	public SrcRect()
	{
	}

	public SrcRect( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public SrcRect( SrcRect s )
	{
		this.attrs = s.attrs;
	}

	public static SrcRect parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "srcRect" ) )
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
					if( endTag.equals( "srcRect" ) )
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
			Logger.logErr( "srcRect.parseOOXML: " + e.toString() );
		}
		SrcRect s = new SrcRect( attrs );
		return s;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:srcRect" );
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = (String) i.next();
				String val = (String) attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SrcRect( this );
	}
}

/**
 * fillRect (Fill Rectangle) This element specifies a fill rectangle. When
 * stretching of an image is specified, a source rectangle, srcRect, is scaled
 * to fit the specified fill rectangle. parent: stretch children: NONE
 */
class FillRect implements OOXMLElement
{

	private static final long serialVersionUID = 7200764163180402065L;
	private HashMap<String, String> attrs = null;

	public FillRect()
	{
	}

	public FillRect( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public FillRect( FillRect f )
	{
		this.attrs = f.attrs;
	}

	public static FillRect parseOOXML( XmlPullParser xpp, Stack<?> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "fillRect" ) )
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
					if( endTag.equals( "fillRect" ) )
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
			Logger.logErr( "fillRect.parseOOXML: " + e.toString() );
		}
		FillRect oe = new FillRect();
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:fillRect" );
		// attributes
		Iterator<?> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FillRect( this );
	}
}

/**
 * FillMode Choice of either tile or stretch elements
 * <p/>
 * tile (Tile) This element specifies that a BLIP should be tiled to fill the
 * available space. This element defines a "tile" rectangle within the bounding
 * box. The image is encompassed within the tile rectangle, and the tile
 * rectangle is tiled across the bounding box to fill the entire area. stretch
 * (Stretch) This element specifies that a BLIP should be stretched to fill the
 * target rectangle. The other option is a tile where a BLIP is tiled to fill
 * the available area.
 * <p/>
 * parent: blipFill choice of : stretch or tile
 */
class FillMode implements OOXMLElement
{

	private static final long serialVersionUID = 967269629502516244L;
	// Since both "child" elements are so small, just output OOXML "by hand"
	private HashMap<String, String> attrs = null;
	private boolean tile = false;

	public FillMode()
	{
	}

	public FillMode( HashMap<String, String> attrs, boolean tile )
	{
		this.attrs = attrs;
		this.tile = tile;
	}

	public FillMode( FillMode f )
	{
		this.attrs = f.attrs;
		this.tile = f.tile;
	}

	public static FillMode parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		boolean tile = false;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "tile" ) )
					{
						tile = true;
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "fillRect" ) )
					{ // only child of
						// stretch
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "tile" ) || endTag.equals( "stretch" ) )
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
			Logger.logErr( "FillMode.parseOOXML: " + e.toString() );
		}
		FillMode fm = new FillMode( attrs, tile );
		return fm;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		// Since both "child" elements are so small, just output OOXML "by hand"
		if( this.tile )
		{
			ooxml.append( "<a:tile" );
		}
		else
		{
			ooxml.append( "<a:stretch><a:fillRect" );
		}
		// attributes
		if( attrs != null )
		{
			Iterator<String> i = attrs.keySet().iterator();
			while( i.hasNext() )
			{
				String key = (String) i.next();
				String val = (String) attrs.get( key );
				ooxml.append( " " + key + "=\"" + val + "\"" );
			}
		}
		if( this.tile )
		{
			ooxml.append( "/>" );
		}
		else
		{
			ooxml.append( "/></a:stretch>" );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FillMode( this );
	}
}

/**
 * gsLst (Gradient Stop List) The list of gradient stops that specifies the
 * gradient colors and their relative positions in the color band.
 * <p/>
 * parent: gradFill children: gs
 */
class GsLst implements OOXMLElement
{

	private static final long serialVersionUID = 6576320251327916221L;
	private ArrayList<Gs> gs;

	public GsLst( ArrayList<Gs> g )
	{
		this.gs = g;
	}

	public GsLst( GsLst gl )
	{
		this.gs = gl.gs;
	}

	public static GsLst parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		ArrayList<Gs> g = new ArrayList<Gs>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gs" ) )
					{
						lastTag.push( tnm );
						g.add( (Gs) Gs.parseOOXML( xpp, lastTag, bk ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "gsLst" ) )
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
			Logger.logErr( "gsLst.parseOOXML: " + e.toString() );
		}
		GsLst gl = new GsLst( g );
		return gl;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:gsLst" );
		if( !gs.isEmpty() )
		{
			ooxml.append( ">" );
			for( int i = 0; i < gs.size(); i++ )
			{
				ooxml.append( ((Gs) gs.get( i )).getOOXML() );
			}
			ooxml.append( "</a:gsLst>" );
		}
		else
		{
			ooxml.append( "/>" );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GsLst( this );
	}
}

/**
 * gs (Gradient stops) This element defines a gradient stop. A gradient stop
 * consists of a position where the stop appears in the color band.
 * <p/>
 * parent: gsLst children: COLORCHOICE
 */
class Gs implements OOXMLElement
{

	private static final long serialVersionUID = 7626866241477598159L;
	private String pos = null;
	private ColorChoice colorChoice = null;

	public Gs( String pos, ColorChoice c )
	{
		this.pos = pos;
		this.colorChoice = c;
	}

	public Gs( Gs g )
	{
		this.pos = g.pos;
		this.colorChoice = g.colorChoice;
	}

	public static Gs parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String pos = null;
		ColorChoice c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gs" ) )
					{ // get attributes
						pos = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "hslClr" ) || tnm.equals( "prstClr" ) || tnm.equals( "schemeClr" ) || tnm.equals( "scrgbClr" ) || tnm
							.equals( "srgbClr" ) || tnm.equals( "sysClr" ) )
					{ // get attributes
						lastTag.push( tnm );
						c = (ColorChoice) ColorChoice.parseOOXML( xpp, lastTag, bk );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "gs" ) )
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
			Logger.logErr( "gs.parseOOXML: " + e.toString() );
		}
		Gs g = new Gs( pos, c );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:gs pos=\"" + this.pos + "\">" );
		ooxml.append( colorChoice.getOOXML() );
		ooxml.append( "</a:gs>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Gs( this );
	}

}
