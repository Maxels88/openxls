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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;

/**
 * Geometry Group: one of prstGeom (Preset shapes) or custGeom (custom shapes)
 */
//TODO: Finish custGeom child elements ahLst, cxnLst, rect (+ finish path element)
public class GeomGroup implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7202561792070909825L;
	private PrstGeom p = null;
	private CustGeom c = null;

	public GeomGroup()
	{
	}

	public GeomGroup( PrstGeom p, CustGeom c )
	{
		this.p = p;
		this.c = c;
	}

	public GeomGroup( GeomGroup g )
	{
		this.p = g.p;
		this.c = g.c;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		PrstGeom p = null;
		CustGeom c = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "prstGeom" ) )
					{
						lastTag.push( tnm );
						p = PrstGeom.parseOOXML( xpp, lastTag );
						lastTag.pop();
						break;
					}
					if( tnm.equals( "custGeom" ) )
					{
						lastTag.push( tnm );
						c = CustGeom.parseOOXML( xpp, lastTag );
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
			Logger.logErr( "GeomGroup.parseOOXML: " + e.toString() );
		}
		GeomGroup g = new GeomGroup( p, c );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( this.p != null )
		{
			ooxml.append( p.getOOXML() );
		}
		else if( this.c != null )
		{
			ooxml.append( c.getOOXML() );
		}
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GeomGroup( this );
	}

}

class PrstGeom implements OOXMLElement
{

	private static final long serialVersionUID = 8327708502983472577L;
	private String prst = null;
	private AvLst avLst = null;

	public PrstGeom()
	{    // no-param constructor, set up common defaults
		prst = "rect";
		avLst = new AvLst();
	}

	public PrstGeom( String prst, AvLst a )
	{
		this.prst = prst;
		this.avLst = a;
	}

	public PrstGeom( PrstGeom p )
	{
		this.prst = p.prst;
		this.avLst = p.avLst;
	}

	public static PrstGeom parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String prst = null;
		AvLst a = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "prstGeom" ) )
					{        // get attributes
						prst = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "avLst" ) )
					{
						lastTag.push( tnm );
						a = (AvLst) AvLst.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "prstGeom" ) )
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
			Logger.logErr( "prstGeom.parseOOXML: " + e.toString() );
		}
		PrstGeom p = new PrstGeom( prst, a );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:prstGeom prst=\"" + prst + "\">" );
		if( avLst != null )
		{
			ooxml.append( avLst.getOOXML() );
		}
		ooxml.append( "</a:prstGeom>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PrstGeom( this );
	}

}

/**
 * custGeom (Custom Geometry)
 * This element specifies the existence of a custom geometric shape. This shape will consist of a series of lines and
 * curves described within a creation path. In addition to this there may also be adjust values, guides, adjust
 * handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
 * <p/>
 * parent:  soPr
 * children:  avLst, gdLst, ahLst, cxnLst, rect, pathLst (REQ)
 */
// TODO: Finish child elements ahLst, cxnLst
class CustGeom implements OOXMLElement
{

	private static final long serialVersionUID = 4036207867619551810L;
	private PathLst pathLst;
	private GdLst gdLst;
	private AvLst avLst;
	private CxnLst cxnLst;
	private Rect rect;

	public CustGeom( PathLst p, GdLst g, AvLst a, CxnLst cx, Rect r )
	{
		this.pathLst = p;
		this.gdLst = g;
		this.avLst = a;
		this.cxnLst = cx;
		this.rect = r;
	}

	public CustGeom( CustGeom c )
	{
		this.pathLst = c.pathLst;
		this.gdLst = c.gdLst;
		this.avLst = c.avLst;
		this.cxnLst = c.cxnLst;
		this.rect = c.rect;
	}

	public static CustGeom parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		PathLst p = null;
		GdLst g = null;
		AvLst a = null;
		CxnLst cx = null;
		Rect r = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pathLst" ) )
					{        // REQ
						lastTag.push( tnm );
						p = PathLst.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "gdLst" ) )
					{
						lastTag.push( tnm );
						g = GdLst.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "avLst" ) )
					{
						lastTag.push( tnm );
						a = (AvLst) AvLst.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "cxnLst" ) )
					{
						lastTag.push( tnm );
						cx = CxnLst.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "rect" ) )
					{
						lastTag.push( tnm );
						r = Rect.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "custGeom" ) )
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
			Logger.logErr( "custGeom.parseOOXML: " + e.toString() );
		}
		CustGeom c = new CustGeom( p, g, a, cx, r );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:custGeom>" );
		if( avLst != null )
		{
			ooxml.append( avLst.getOOXML() );    // avLst
		}
		if( gdLst != null )
		{
			ooxml.append( gdLst.getOOXML() );    // gdLst
		}
		// TODO: ahLst
		ooxml.append( "<a:ahLst/>" );
		if( cxnLst != null )
		{
			ooxml.append( cxnLst.getOOXML() );    // cxnLst
		}
		if( rect != null )
		{
			ooxml.append( rect.getOOXML() );    // rect
		}
		if( pathLst != null )
		{
			ooxml.append( pathLst.getOOXML() );    // pathLst
		}
		ooxml.append( "</a:custGeom>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CustGeom( this );
	}
}

/**
 * pathLst (List of Shape Paths)
 * This element specifies the entire path that is to make up a single geometric shape. The pathLst can consist of
 * many individual paths within it.
 * <p/>
 * parent:  custGeom
 * children path (multiple)
 */
class PathLst implements OOXMLElement
{

	private static final long serialVersionUID = -1996347204024728000L;
	private ArrayList<Path> path;
	private HashMap<String, String> attrs = null;

	public PathLst( HashMap<String, String> attrs, ArrayList<Path> p )
	{
		this.attrs = attrs;
		this.path = p;
	}

	public PathLst( PathLst pl )
	{
		this.attrs = pl.attrs;
		this.path = pl.path;
	}

	public static PathLst parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		ArrayList<Path> p = new ArrayList<Path>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pathLst" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "path" ) )
					{        // get one or more path children
						lastTag.push( tnm );
						p.add( Path.parseOOXML( xpp, lastTag ) );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "pathLst" ) )
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
			Logger.logErr( "pathLst.parseOOXML: " + e.toString() );
		}
		PathLst pl = new PathLst( attrs, p );
		return pl;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:pathLst" );
		// attributes
		Iterator<String> ii = attrs.keySet().iterator();
		while( ii.hasNext() )
		{
			String key = ii.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( path != null )
		{
			for( Path aPath : path )
			{
				ooxml.append( aPath.getOOXML() );
			}
		}
		ooxml.append( "</a:pathLst>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new PathLst( this );
	}
}

/**
 *
 *
 */
class Path implements OOXMLElement
{

	private static final long serialVersionUID = 6906237439620322589L;
	private HashMap<String, String> attrs = null;

	public Path( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Path( Path p )
	{
		this.attrs = p.attrs;
	}

	public static Path parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "path" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "moveTo" ) )
					{
					}
					else if( tnm.equals( "lnTo" ) )
					{
					}
					else if( tnm.equals( "close" ) )
					{
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "path" ) )
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
			Logger.logErr( "path.parseOOXML: " + e.toString() );
		}
		Path p = new Path( attrs );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:path" );
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
		return new Path( this );
	}
}

/**
 * gdLst (List of Shape Guides)
 * This element specifies all the guides that will be used for this shape. A guide is specified by the gd element and
 * defines a calculated value that may be used for the construction of the corresponding shape.
 */
class GdLst implements OOXMLElement
{

	private static final long serialVersionUID = -7852193131141462744L;
	private ArrayList<Gd> gds;

	public GdLst( ArrayList<Gd> gds )
	{
		this.gds = gds;
	}

	public GdLst( GdLst g )
	{
		this.gds = g.gds;
	}

	public static GdLst parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		ArrayList<Gd> gds = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "gd" ) )
					{
						lastTag.push( tnm );
						if( gds == null )
						{
							gds = new ArrayList<Gd>();
						}
						gds.add( (Gd) Gd.parseOOXML( xpp, lastTag ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "gdLst" ) )
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
			Logger.logErr( "gdLst.parseOOXML: " + e.toString() );
		}
		GdLst g = new GdLst( gds );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:gdLst>" );
		if( gds != null )
		{
			for( Gd gd : gds )
			{
				ooxml.append( gd.getOOXML() );
			}
		}
		ooxml.append( "</a:gdLst>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GdLst( this );
	}
}

/**
 ahLst (List of Shape Adjust Handles)
 * This element specifies the adjust handles that will be applied to a custom geometry. These adjust handles will
 * specify points within the geometric shape that can be used to perform certain transform operations on the
 * shape.
 * [Example: Consider the scenario where a custom geometry, an arrow in this case, has been drawn and adjust
 * handles have been placed at the top left corner of both the arrow head and arrow body. The user interface can
 * then be made to transform only certain parts of the shape by using the corresponding adjust handle
 *
 * parent: custGeom
 * children:  one or more of [ahXY, ahPolar (both REQ)]
 */

/**
 * rect (Shape Text Rectangle)
 * <p/>
 * This element specifies the rectangular bounding box for text within a custGeom shape. The default for this
 * rectangle is the bounding box for the shape. This can be modified using this elements four attributes to inset or
 * extend the text bounding box.
 * <p/>
 * parent: custGeom
 * children: none
 */
class Rect implements OOXMLElement
{

	private static final long serialVersionUID = 2790708601254975676L;
	private HashMap<String, String> attrs;

	public Rect( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Rect( Rect r )
	{
		this.attrs = r.attrs;
	}

	public static Rect parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "rect" ) )
					{        // get attributes
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
					if( endTag.equals( "rect" ) )
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
			Logger.logErr( "rect.parseOOXML: " + e.toString() );
		}
		Rect r = new Rect( attrs );
		return r;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:rect" );
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
		return new Rect( this );
	}
}

/**
 * cxnLst (List of Shape Connection Sites)
 * This element specifies all the connection sites that will be used for this shape. A connection site is specified by
 * defining a point within the shape bounding box that can have a cxnSp element attached to it. These connection
 * sites are specified using the shape coordinate system that is specified within the ext transform element.
 * <p/>
 * parents:		custGeom
 * children:	cxn	(0 or more)
 */
class CxnLst implements OOXMLElement
{

	private static final long serialVersionUID = -562847539163221621L;
	private ArrayList<Cxn> cxns;

	public CxnLst( ArrayList<Cxn> cxns )
	{
		this.cxns = cxns;
	}

	public CxnLst( CxnLst c )
	{
		this.cxns = c.cxns;
	}

	public static CxnLst parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		ArrayList<Cxn> cxns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cxn" ) )
					{
						lastTag.push( tnm );
						if( cxns == null )
						{
							cxns = new ArrayList<Cxn>();
						}
						cxns.add( Cxn.parseOOXML( xpp, lastTag ) );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cxnLst" ) )
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
			Logger.logErr( "cxnLst.parseOOXML: " + e.toString() );
		}
		CxnLst c = new CxnLst( cxns );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:cxnLst>" );
		if( cxns != null )
		{
			for( Cxn cxn : cxns )
			{
				ooxml.append( cxn.getOOXML() );
			}
		}
		ooxml.append( "</a:cxnLst>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CxnLst( this );
	}
}

/**
 * cxn (Shape Connection Site)
 * This element specifies the existence of a connection site on a custom shape. A connection site allows a cxnSp to
 * be attached to this shape. This connection will be maintined when the shape is repositioned within the
 * document. It should be noted that this connection is placed within the shape bounding box using the transform
 * coordinate system which is also called the shape coordinate system, as it encompasses theentire shape. The
 * width and height for this coordinate system are specified within the ext transform element.
 * <p/>
 * parents: 	cxnLst
 * children:    pos	REQ
 */
class Cxn implements OOXMLElement
{

	private static final long serialVersionUID = -4193511102420582252L;
	private HashMap<String, String> attrs;
	private Pos pos = null;

	public Cxn( HashMap<String, String> attrs, Pos p )
	{
		this.attrs = attrs;
		this.pos = p;
	}

	public Cxn( Cxn c )
	{
		this.attrs = c.attrs;
		this.pos = c.pos;
	}

	public static Cxn parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		Pos p = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cxn" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );    // ang REQ
						}
					}
					else if( tnm.equals( "pos" ) )
					{
						lastTag.push( tnm );
						p = Pos.parseOOXML( xpp, lastTag );

					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cxn" ) )
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
			Logger.logErr( "cxn.parseOOXML: " + e.toString() );
		}
		Cxn c = new Cxn( attrs, p );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:cxn" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = i.next();
			String val = attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( pos != null )
		{
			ooxml.append( pos.getOOXML() );
		}
		ooxml.append( "</a:cxn>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Cxn( this );
	}
}

/**
 * pos (Shape Position Coordinate)
 * Specifies a position coordinate within the shape bounding box. It should be noted that this coordinate is placed
 * within the shape bounding box using the transform coordinate system which is also called the shape coordinate
 * system, as it encompasses the entire shape. The width and height for this coordinate system are specified within
 * the ext transform element.
 * <p/>
 * parents:  cxn, ahPolar, ahXY
 * children: none
 */
class Pos implements OOXMLElement
{

	private static final long serialVersionUID = 5500991309750603125L;
	private HashMap<String, String> attrs;

	public Pos( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public Pos( Pos p )
	{
		this.attrs = p.attrs;
	}

	public static Pos parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "pos" ) )
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
					if( endTag.equals( "pos" ) )
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
			Logger.logErr( "pos.parseOOXML: " + e.toString() );
		}
		Pos p = new Pos( attrs );
		return p;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:pos" );
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
		return new Pos( this );
	}
}

