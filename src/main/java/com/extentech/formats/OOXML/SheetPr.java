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

/**
 * sheetPr (Sheet Properties)
 * <p/>
 * Sheet-level properties.
 * <p/>
 * parent:  worksheet, dialogsheet
 * children:  tabColor, outlinePr, pageSetUpPr
 */
// TODO: Finish pageSetUpPr  + input into 2003 sheet settings **
public class SheetPr implements OOXMLElement
{

	private static final long serialVersionUID = 1781567781060400234L;
	private HashMap<String, String> attrs;
	private TabColor tab;
	private OutlinePr outlinePr;
	private PageSetupPr pageSetupPr;

	public SheetPr( HashMap<String, String> attrs, TabColor tab, OutlinePr op, PageSetupPr pr )
	{
		this.attrs = attrs;
		this.tab = tab;
		this.outlinePr = op;
		this.pageSetupPr = pr;
	}

	public SheetPr( SheetPr sp )
	{
		this.attrs = sp.attrs;
		this.tab = sp.tab;
		this.outlinePr = sp.outlinePr;
		this.pageSetupPr = sp.pageSetupPr;
	}

	/**
	 * outlinePr
	 * pageSetUpPr
	 * tabColor
	 * codeName
	 * enableFormatConditionsCalculation
	 * filterMode
	 * published
	 * syncHorizontal
	 * syncRef
	 * syncVertical
	 * transitionEntry
	 * transitionEvaluation
	 *
	 * @param xpp
	 * @return
	 */
	public static SheetPr parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		TabColor tab = null;
		OutlinePr op = null;
		PageSetupPr pr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "sheetPr" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "pageSetUpPr" ) )
					{ // Page setup properties of the worksheet
						pr = (PageSetupPr) PageSetupPr.parseOOXML( xpp );
					}
					else if( tnm.equals( "outlinePr" ) )
					{
						op = (OutlinePr) OutlinePr.parseOOXML( xpp );
					}
					else if( tnm.equals( "tabColor" ) )
					{
						tab = (TabColor) TabColor.parseOOXML( xpp );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "sheetPr" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "sheetPr.parseOOXML: " + e.toString() );
		}
		SheetPr sp = new SheetPr( attrs, tab, op, pr );
		return sp;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<sheetPr" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}

		ooxml.append( ">" );
		// ordered sequence:
		if( tab != null )
		{
			ooxml.append( tab.getOOXML() );
		}
		if( outlinePr != null )
		{
			ooxml.append( outlinePr.getOOXML() );
		}
		if( pageSetupPr != null )
		{
			ooxml.append( pageSetupPr.getOOXML() );
		}
		ooxml.append( "</sheetPr>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new SheetPr( this );
	}

	/**
	 * return the codename used to link this sheet to vba code
	 *
	 * @return
	 */
	public String getCodename()
	{
		if( attrs != null )
		{
			return (String) attrs.get( "codeName" );
		}
		return null;
	}

	/**
	 * set the codename used to link this sheet to vba code
	 *
	 * @param codename
	 */
	public void setCodename( String codename )
	{
		if( attrs == null )
		{
			attrs = new HashMap<String, String>();
		}
		attrs.put( "codeName", codename );
	}
}

/**
 * (Sheet Tab Color)
 * Background color of the sheet tab.
 * <p/>
 * parent: 	sheetPr
 * children: none
 */
class TabColor implements OOXMLElement
{

	private static final long serialVersionUID = -2862996863147633555L;
	private HashMap<String, String> attrs;

	public TabColor( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public TabColor( TabColor t )
	{
		this.attrs = t.attrs;
	}

	public static TabColor parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "tabColor" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "tabColor" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "tabColor.parseOOXML: " + e.toString() );
		}
		TabColor t = new TabColor( attrs );
		return t;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<tabColor" );
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

	public OOXMLElement cloneElement()
	{
		return new TabColor( this );
	}
}

/**
 * outlinePr
 * Sheet Group Outline Settings
 * all attributes are optional, default values:
 * applyStyles="0"	showOutlineSymbols="1"	summaryBelow="1"	summaryRight="1"
 */
class OutlinePr implements OOXMLElement
{

	private static final long serialVersionUID = 3030511803286369045L;
	private HashMap<String, String> attrs;

	public OutlinePr( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public OutlinePr( OutlinePr op )
	{
		this.attrs = op.attrs;
	}

	public static OutlinePr parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "outlinePr" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "outlinePr" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "OutlinePr.parseOOXML: " + e.toString() );
		}
		OutlinePr op = new OutlinePr( attrs );
		return op;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<outlinePr" );
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

	public OOXMLElement cloneElement()
	{
		return new OutlinePr( this );
	}
}

/**
 * pageSetUpPr (Page Setup Properties)
 * Page setup properties of the worksheet
 * attributes:	autoPageBreaks (default=true), fitToPage (default=false)
 */
class PageSetupPr implements OOXMLElement
{

	private static final long serialVersionUID = 3030511803286369045L;
	private boolean autoPageBreaks = true, fitToPage = false;

	public PageSetupPr( boolean autoPageBreaks, boolean fitToPage )
	{
		this.autoPageBreaks = autoPageBreaks;
		this.fitToPage = fitToPage;
	}

	public PageSetupPr( PageSetupPr pr )
	{
		this.autoPageBreaks = pr.autoPageBreaks;
		this.fitToPage = pr.fitToPage;
	}

	public static PageSetupPr parseOOXML( XmlPullParser xpp )
	{
		boolean autoPageBreaks = true, fitToPage = false;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pageSetUpPr" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "fitToPage" ) )
							{
								fitToPage = (xpp.getAttributeValue( i ).equals( "1" ));
							}
							else if( nm.equals( "autoPageBreaks" ) )
							{
								autoPageBreaks = (xpp.getAttributeValue( i ).equals( "1" ));
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "pageSetUpPr" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "PageSetupPr: " + e.toString() );
		}
		PageSetupPr pr = new PageSetupPr( autoPageBreaks, fitToPage );
		return pr;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<pageSetUpPr" );
		// attributes
		if( !autoPageBreaks ) // if not default
		{
			ooxml.append( " autoPageBreaks=\"0\"" );
		}
		if( fitToPage )    // if not default
		{
			ooxml.append( " fitToPage=\"1\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new PageSetupPr( this );
	}
}


