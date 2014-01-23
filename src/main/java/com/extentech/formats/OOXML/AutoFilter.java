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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * autoFilter (AutoFilter Settings)
 * AutoFilter temporarily hides rows based on a filter criteria, which is applied column by column to a table of data
 * in the worksheet. This collection expresses AutoFilter settings.
 * <p/>
 * parent: 		worksheet, table, filter, customSheetView
 * children: 	filterColumn (0+), sortState
 * attributes:	ref
 */
// TODO: finish sortState
// TODO: finish filterColumn children filters->filter, dataGroupItem
public class AutoFilter implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( AutoFilter.class );
	private static final long serialVersionUID = 7111401348177004218L;
	private String ref = null;
	private ArrayList<FilterColumn> filterColumns = null;

	public AutoFilter( String ref, ArrayList<FilterColumn> f )
	{
		this.ref = ref;
		this.filterColumns = f;
	}

	public AutoFilter( AutoFilter a )
	{
		this.ref = a.ref;
		this.filterColumns = a.filterColumns;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp )
	{
		String ref = null;
		ArrayList<FilterColumn> f = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "autoFilter" ) )
					{        // get ref attribute
						if( xpp.getAttributeCount() == 1 )
						{
							ref = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "sortState" ) )
					{
					}
					else if( tnm.equals( "filterColumn" ) )
					{
						if( f == null )
						{
							f = new ArrayList<>();
						}
						f.add( FilterColumn.parseOOXML( xpp ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "autoFilter" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "autoFilter.parseOOXML: " + e.toString() );
		}
		AutoFilter a = new AutoFilter( ref, f );
		return a;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<autoFilter" );
		if( ref != null )
		{
			ooxml.append( " ref=\"" + ref + "\"" );
		}
		ooxml.append( ">" );
		if( filterColumns != null )
		{
			for( FilterColumn filterColumn : filterColumns )
			{
				ooxml.append( filterColumn.getOOXML() );
			}
		}
		ooxml.append( "</autoFilter>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new AutoFilter( this );
	}
}

/**
 * filterColumn (AutoFilter Column)
 * The filterColumn collection identifies a particular column in the AutoFilter range and specifies filter information
 * that has been applied to this column. If a column in the AutoFilter range has no criteria specified, then there is
 * no corresponding filterColumn collection expressed for that column
 * <p/>
 * parent: 		autoFilter
 * children:	CHOICE OF: colorFilter, customFilters, dynamicFilter, filters, iconFilter, top10
 * attributes:	colId REQ, hiddenButton, showButton
 */
class FilterColumn implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( FilterColumn.class );
	private static final long serialVersionUID = 5005589034415840928L;
	private HashMap<String, String> attrs = null;
	private Object filter = null;        // CHOICE of filter

	public FilterColumn( HashMap<String, String> attrs, Object filter )
	{
		this.attrs = attrs;
		this.filter = filter;
	}

	public FilterColumn( FilterColumn f )
	{
		this.attrs = f.attrs;
		this.filter = f.filter;
	}

	public static FilterColumn parseOOXML( XmlPullParser xpp )
	{
		HashMap<String, String> attrs = new HashMap<>();
		Object filter = null;        // CHOICE of filter
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "filterColumn" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "colorFilter" ) )
					{
						filter = ColorFilter.parseOOXML( xpp );
					}
					else if( tnm.equals( "customFilters" ) )
					{
					}
					else if( tnm.equals( "dynamicFilter" ) )
					{
					}
					else if( tnm.equals( "filters" ) )
					{
					}
					else if( tnm.equals( "iconFilter" ) )
					{
					}
					else if( tnm.equals( "top10" ) )
					{
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "filterColumn" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "filterColumn.parseOOXML: " + e.toString() );
		}
		FilterColumn f = new FilterColumn( attrs, filter );
		return f;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<filterColumn " );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( filter instanceof ColorFilter )
		{
			ooxml.append( ((ColorFilter) filter).getOOXML() );
		}
		if( filter instanceof CustomFilters )
		{
			ooxml.append( ((CustomFilters) filter).getOOXML() );
		}
		if( filter instanceof DynamicFilter )
		{
			ooxml.append( ((DynamicFilter) filter).getOOXML() );
		}
		if( filter instanceof Filters )
		{
			ooxml.append( ((Filters) filter).getOOXML() );
		}
		if( filter instanceof IconFilter )
		{
			ooxml.append( ((IconFilter) filter).getOOXML() );
		}
		if( filter instanceof Top10 )
		{
			ooxml.append( ((Top10) filter).getOOXML() );
		}
		ooxml.append( "</filterColumn>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new FilterColumn( this );
	}
}

/**
 * colorFilter (Color Filter Criteria)
 * This element specifies the color to filter by and whether to use the cell's fill or font color in the filter criteria. If
 * the cell's font or fill color does not match the color specified in the criteria, the rows corresponding to those cells
 * are hidden from view.
 * <p/>
 * parent: 	filterColumn
 * children: none
 */
class ColorFilter implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( ColorFilter.class );
	private static final long serialVersionUID = 7077951504723033275L;
	private HashMap<String, String> attrs = null;

	public ColorFilter( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public ColorFilter( ColorFilter c )
	{
		this.attrs = c.attrs;
	}

	public static ColorFilter parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "colorFilter" ) )
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
					if( endTag.equals( "colorFilter" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "colorFilter.parseOOXML: " + e.toString() );
		}
		ColorFilter oe = new ColorFilter( attrs );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<colorFilter" );
		// attributes
		Iterator<?> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ColorFilter( this );
	}
}

/**
 * dynamicFilter (Dynamic Filter)
 * This collection specifies dynamic filter criteria. These criteria are considered dynamic because they can change,
 * either with the data itself (e.g., "above average") or with the current system date (e.g., show values for "today").
 * For any cells whose values do not meet the specified criteria, the corresponding rows shall be hidden from view
 * when the filter is applied.
 * <p/>
 * parent: 	filterColumn
 * children: none
 */
class DynamicFilter implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( DynamicFilter.class );
	private static final long serialVersionUID = -473171074711686551L;
	private HashMap<String, String> attrs = null;

	public DynamicFilter( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public DynamicFilter( DynamicFilter d )
	{
		this.attrs = d.attrs;
	}

	public static DynamicFilter parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "dynamicFilter" ) )
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
					if( endTag.equals( "dynamicFilter" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "dynamicFilter.parseOOXML: " + e.toString() );
		}
		DynamicFilter d = new DynamicFilter( attrs );
		return d;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<dynamicFilter" );
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
		return new DynamicFilter( this );
	}
}

/**
 * iconFilter (Icon Filter)
 * <p/>
 * This element specifies the icon set and particular icon within that set to filter by. For any cells whose icon does
 * not match the specified criteria, the corresponding rows shall be hidden from view when the filter is applied.
 * <p/>
 * parent: filterColumn
 * children: none
 */
class IconFilter implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( IconFilter.class );
	private static final long serialVersionUID = -5897037678209125965L;
	private HashMap<String, String> attrs = null;

	public IconFilter( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public IconFilter( IconFilter i )
	{
		this.attrs = i.attrs;
	}

	public static IconFilter parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "iconFilter" ) )
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
					if( endTag.equals( "iconFilter" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "iconFilter.parseOOXML: " + e.toString() );
		}
		IconFilter i = new IconFilter( attrs );
		return i;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<iconFilter" );
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
		return new IconFilter( this );
	}
}

/**
 * customFilters (Custom Filters)
 * When there is more than one custom filter criteria to apply (an 'and' or 'or' joining two criteria), then this
 * element groups the customFilter elements together.
 */
class CustomFilters implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CustomFilters.class );
	private static final long serialVersionUID = -2491942158519963335L;
	private boolean and = false;
	private CustomFilter[] custfilter = null;

	public CustomFilters( boolean and, CustomFilter[] custfilter )
	{
		this.and = and;
		this.custfilter = custfilter;
	}

	public CustomFilters( CustomFilters c )
	{
		this.and = c.and;
		this.custfilter = c.custfilter;
	}

	public static CustomFilters parseOOXML( XmlPullParser xpp )
	{
		boolean and = false;
		CustomFilter[] custfilter = new CustomFilter[2];
		int idx = 0;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "customFilters" ) )
					{        // get attributes
						if( xpp.getAttributeCount() == 1 )
						{
							and = (xpp.getAttributeValue( 0 ).equals( "1" ));
						}
					}
					else if( tnm.equals( "customFilter" ) )
					{    // 1-2
						custfilter[idx++] = CustomFilter.parseOOXML( xpp );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "customFilters" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "customFilters.parseOOXML: " + e.toString() );
		}
		CustomFilters c = new CustomFilters( and, custfilter );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<customFilters" );
		if( this.and )
		{
			ooxml.append( " and=\"1\"" );
		}
		ooxml.append( ">" );
		if( custfilter != null )
		{    // shouln't be!
			if( custfilter[0] != null )// shouldn't be!
			{
				ooxml.append( custfilter[0].getOOXML() );
			}
			if( custfilter[1] != null )
			{
				ooxml.append( custfilter[1].getOOXML() );
			}
		}
		ooxml.append( "</customFilters>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CustomFilters( this );
	}
}

/**
 * customFilter (Custom Filter Criteria)
 * A custom AutoFilter specifies an operator and a value. There can be at most two customFilters specified, and in
 * that case the parent element specifies whether the two conditions are joined by 'and' or 'or'. For any cells
 * whose values do not meet the specified criteria, the corresponding rows shall be hidden from view when the
 * filter is applied.
 * <p/>
 * parent: customFilters
 * children: none
 */
class CustomFilter implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( CustomFilter.class );
	private static final long serialVersionUID = 7995078604042667255L;
	private HashMap<String, String> attrs = null;

	public CustomFilter( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public CustomFilter( CustomFilter c )
	{
		this.attrs = c.attrs;
	}

	public static CustomFilter parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "customFilter" ) )
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
					if( endTag.equals( "customFilter" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "customFilter.parseOOXML: " + e.toString() );
		}
		CustomFilter c = new CustomFilter( attrs );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<customFilter" );
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
		return new CustomFilter( this );
	}
}

/**
 * top10 (Top 10)
 * This element specifies the top N (percent or number of items) to filter by.
 * <p/>
 * parent: filterColumn
 * children: none
 */
class Top10 implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Top10.class );
	private static final long serialVersionUID = 77735498689922082L;
	private HashMap<String, String> attrs = null;

	public Top10( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Top10( Top10 t )
	{
		this.attrs = t.attrs;
	}

	public static Top10 parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "top10" ) )
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
					if( endTag.equals( "top10" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "top10.parseOOXML: " + e.toString() );
		}
		Top10 t = new Top10( attrs );
		return t;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<top10" );
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
		return new Top10( this );
	}
}

/**
 * filters (Filter Criteria)
 * When multiple values are chosen to filter by, or when a group of date values are chosen to filter by, this element
 * groups those criteria together.
 * <p/>
 * parent:  filterColumn
 * children: filter (0+), dateGroupItem (0+)
 */
// TODO: finish children
class Filters implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( Filters.class );
	private static final long serialVersionUID = 921424089049938924L;
	private HashMap<String, String> attrs = null;

	public Filters( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Filters( Filters f )
	{
		this.attrs = f.attrs;
	}

	public static Filters parseOOXML( XmlPullParser xpp )
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
					if( tnm.equals( "filters" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "filter" ) )
					{
					}
					else if( tnm.equals( "dateGroupItem" ) )
					{
						//layout = (layout) layout.parseOOXML(xpp, lastTag).clone();
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "filters" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "filters.parseOOXML: " + e.toString() );
		}
		Filters oe = new Filters( attrs );
		return oe;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<filters" );
		// attributes
		Iterator<?> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		// TODO: one or more filter elements
		// TODO: one or more dateGroupItem elements
		ooxml.append( "</filters>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Filters( this );
	}
}

	
