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

import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * EffectPropsGroup Effect Properties either effectDag or effectLst
 */
//TODO: FINISH CHILD ELEMENTS for both effectDag and effectLst
public class EffectPropsGroup implements OOXMLElement
{

	private static final long serialVersionUID = 8250236905326475833L;

	private EffectDag effectDag;
	private EffectLst effectLst;

	public EffectPropsGroup( EffectDag ed, EffectLst el )
	{
		this.effectDag = ed;
		this.effectLst = el;
	}

	public EffectPropsGroup( EffectPropsGroup e )
	{
		this.effectDag = e.effectDag;
		this.effectLst = e.effectLst;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		EffectDag ed = null;
		EffectLst el = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "effectDag" ) )
					{
						lastTag.push( tnm );
						ed = (EffectDag) EffectDag.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "effectLst" ) )
					{
						lastTag.push( tnm );
						el = (EffectLst) EffectLst.parseOOXML( xpp, lastTag );
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
			Logger.logErr( "EffectPropsGroup.parseOOXML: " + e.toString() );
		}
		EffectPropsGroup e = new EffectPropsGroup( ed, el );
		return e;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( effectDag != null )
		{
			ooxml.append( effectDag.getOOXML() );
		}
		if( effectLst != null )
		{
			ooxml.append( effectLst.getOOXML() );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new EffectPropsGroup( this );
	}
}

/**
 * effectDag (Effect Container)
 * This element specifies a list of effects. Effects are applied in the order specified by the container type (sibling or
 * tree).
 * <p/>
 * parent: many
 * children: MANY (EFFECT)
 */ // TODO: FINISH CHILD ELEMENTS
class EffectDag implements OOXMLElement
{

	private static final long serialVersionUID = 4786440439664356745L;
	private HashMap<String, String> attrs = null;

	public EffectDag( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public EffectDag( EffectDag e )
	{
		this.attrs = e.attrs;
	}

	public static EffectDag parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "effectDag" ) )
					{
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "CHILDELEMENT" ) )
					{
						lastTag.push( tnm );
						//layout = (layout) layout.parseOOXML(xpp, lastTag).clone();

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "effectDag" ) )
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
			Logger.logErr( "effectDag.parseOOXML: " + e.toString() );
		}
		EffectDag e = new EffectDag( attrs );
		return e;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:effectDag" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
//    	if (CHILD!=null) { ooxml.append(CHILD.getOOXML());
		ooxml.append( "</a:effectDag>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new EffectDag( this );
	}
}

/**
 * effectLst (Effect Container)
 * This element specifies a list of effects. Effects in an effectLst are applied in the default order by the rendering
 * engine. The following diagrams illustrate the order in which effects are to be applied, both for shapes and for
 * group shapes.
 * <p/>
 * parent: many
 * children: MANY (EFFECT)
 */ // TODO: FINISH CHILD ELEMENTS
class EffectLst implements OOXMLElement
{

	private static final long serialVersionUID = -6164888373165090983L;

	//	public effectLst() { 	}
	public EffectLst()
	{
	}

	public EffectLst( EffectLst e )
	{
	}

	public static EffectLst parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "CHILDELEMENT" ) )
					{
						//lastTag.push(tnm);
						//layout = (layout) layout.parseOOXML(xpp, lastTag).clone();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "effectLst" ) )
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
			Logger.logErr( "effectLst.parseOOXML: " + e.toString() );
		}
		EffectLst e = new EffectLst();
		return e;
	}

	@Override
	public String getOOXML()
	{
		//StringBuffer ooxml= new StringBuffer();	
//    	ooxml.append("<a:effectLst");
//    	if (CHILD!=null) { ooxml.append(CHILD.getOOXML());
//    	ooxml.append("</a:effectLst>");
		return "<a:effectLst/>";    // TODO: FINISH CHILD ELEMENTS
//    	return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new EffectLst( this );
	}
}

