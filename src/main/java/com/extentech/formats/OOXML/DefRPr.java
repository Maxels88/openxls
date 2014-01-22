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

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * defRPr (Default Text Run Properties)
 * <p/>
 * This element contains all default run level text properties for the text runs within a containing paragraph. These
 * properties are to be used when overriding properties have not been defined within the rPr element
 * <p/>
 * parent: many, including  pPr
 * children:  ln, FILLS, EFFECTS, highlight, TEXTUNDERLINE, TEXTUNDERLINEFILL, latin, ea, cs, sym, hlinkClick, hlinkMouseOver
 * many attributes
 */
// TODO: FINISH CHILD ELEMENTS highlight TEXTUNDERLINE, TEXTUNDERLINEFILL, sym, hlinkClick, hlinkMouseOver
public class DefRPr implements OOXMLElement
{

	private static final long serialVersionUID = 6764149567499222506L;
	private FillGroup fillGroup = null;
	private EffectPropsGroup effect = null;
	private Ln line = null;
	private HashMap<String, String> attrs = null;
	private String latin = null, ea = null, cs = null;    // really children but only have 1 attribute and no children

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		FillGroup fill = null;
		EffectPropsGroup effect = null;
		Ln l = null;
		String latin = null, ea = null, cs = null;
		HashMap<String, String> attrs = new HashMap<String, String>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "defRPr" ) )
					{        // default text properties
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
					if( endTag.equals( "defRPr" ) )
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
			Logger.logErr( "defPr.parseOOXML: " + e.toString() );
		}
		DefRPr dp = new DefRPr( fill, effect, l, attrs, latin, ea, cs );
		return dp;
	}

	/**
	 * return all the text properties in hashmap from
	 *
	 * @return
	 */
	public HashMap<String, String> getTextProperties()
	{
		HashMap<String, String> textprops = new HashMap<String, String>();
		textprops.putAll( attrs );
		textprops.put( "latin_typeface", latin );
		textprops.put( "ea_typeface", ea );
		textprops.put( "cs_typeface", cs );
		// TODO: Fill, line ...
		return textprops;
	}

	/**
	 * create a default default Run Properties
	 * for BIFF8 compatibility
	 * qaa
	 */
	public DefRPr()
	{
		this.attrs = new HashMap<String, String>();
		this.attrs.put( "sz", "900" );
		this.attrs.put( "b", "1" );
		this.attrs.put( "i", "0" );
		this.attrs.put( "u", "none" );
		this.attrs.put( "strike", "noStrike" );
		this.attrs.put( "baseline", "0" );
		this.fillGroup = new FillGroup( null, null, null, null, new SolidFill() );
		this.latin = "Arial";
		this.ea = "Arial";
		this.cs = "Arial";
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
	public DefRPr( String fontFace, int sz, boolean b, boolean i, String u, String strike, String clr )
	{
		this.attrs = new HashMap<String, String>();
		this.attrs.put( "sz", String.valueOf( sz ) );
		this.attrs.put( "b", (b ? "1" : "0") );
		this.attrs.put( "i", (i ? "1" : "0") );
		this.attrs.put( "u", u );
		this.attrs.put( "strike", strike );
		this.attrs.put( "baseline", "0" );
		this.fillGroup = new FillGroup( null, null, null, null, new SolidFill( clr ) );
		this.latin = fontFace;
		this.ea = fontFace;
		this.cs = fontFace;
	}

	public DefRPr( FillGroup fill, EffectPropsGroup effect, Ln l, HashMap<String, String> attrs, String latin, String ea, String cs )
	{
		this.fillGroup = fill;
		this.effect = effect;
		this.line = l;
		this.latin = latin;
		this.ea = ea;
		this.cs = cs;
		this.attrs = attrs;
	}

	public DefRPr( DefRPr dp )
	{
		this.fillGroup = dp.fillGroup;
		this.effect = dp.effect;
		this.line = dp.line;
		this.latin = dp.latin;
		this.ea = dp.ea;
		this.cs = dp.cs;
		this.attrs = dp.attrs;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:defRPr" );
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( line != null )
		{
			ooxml.append( line.getOOXML() );
		}
		if( fillGroup != null )
		{
			ooxml.append( fillGroup.getOOXML() );        // group fill
		}
		if( effect != null )
		{
			ooxml.append( effect.getOOXML() );  // group effect
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
		// hLinkClick
		// hLinkMouseOver
		ooxml.append( "</a:defRPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new DefRPr( this );
	}
}

