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
 * ln (Outline)
 * <p/>
 * This OOXML/DrawingML element specifies an outline style that can be applied to a number of different objects such as shapes and
 * text. The line allows for the specifying of many different types of outlines including even line dashes and bevels
 * <p/>
 * parent:  many
 * children: bevel, custDash, gradFill, headEnd, miter, noFill, pattFill, prstDash, round, solidFill, tailEnd
 * attributes:  w (line width) cap (line ending cap) cmpd (compound line type) algn (stroke alignment)
 */
// TODO: Finish custDash
public class Ln implements OOXMLElement
{

	private static final long serialVersionUID = -161619607936083688L;
	private HashMap<String, String> attrs = null;
	private FillGroup fill = null;
	private JoinGroup join = null;
	private DashGroup dash = null;
	private HeadEnd h = null;
	private TailEnd t = null;

	/**
	 * create a new solid line in width w with color clr
	 *
	 * @param w   int specifies the width of a line in EMUs. 1 pt = 12700 EMUs
	 * @param clr String hex color without the #
	 */
	public Ln( int w, String clr )
	{
		this.fill = new FillGroup( null, null, null, null, new SolidFill( clr ) );
		this.dash = new DashGroup( new PrstDash( "solid" ) );
		this.setWidth( w );
	}

	public Ln()
	{    // no-param constructor, set up common defaults
		this.fill = new FillGroup( null, null, null, null, new SolidFill() );
		this.join = new JoinGroup( "800000", true, false, false );
		this.h = new HeadEnd();
		this.t = new TailEnd();
	}

	public Ln( HashMap<String, String> attrs, FillGroup fill, JoinGroup join, DashGroup dash, HeadEnd h, TailEnd t )
	{
		this.attrs = attrs;
		this.fill = fill;
		this.join = join;
		this.dash = dash;
		this.t = t;
		this.h = h;
	}

	public Ln( Ln l )
	{
		this.attrs = l.attrs;
		this.fill = l.fill;
		this.join = l.join;
		this.dash = l.dash;
		this.t = l.t;
		this.h = l.h;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		HashMap<String, String> attrs = new HashMap<String, String>();
		FillGroup fill = null;
		JoinGroup join = null;
		DashGroup dash = null;
		HeadEnd h = null;
		TailEnd t = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "ln" ) )
					{        // get ln attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
					else if( tnm.equals( "noFill" ) ||
							tnm.equals( "solidFill" ) ||
							tnm.equals( "pattFill" ) ||
							tnm.equals( "gradFill" ) )
					{
						lastTag.push( tnm );
						fill = (FillGroup) FillGroup.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "bevel" ) ||
							tnm.equals( "round" ) ||
							tnm.equals( "miter" ) )
					{
						lastTag.push( tnm );
						join = (JoinGroup) JoinGroup.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "prstDash" ) || tnm.equals( "custDash" ) )
					{
						lastTag.push( tnm );
						dash = (DashGroup) DashGroup.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "headEnd" ) )
					{
						lastTag.push( tnm );
						h = (HeadEnd) HeadEnd.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "tailEnd" ) )
					{
						lastTag.push( tnm );
						t = (TailEnd) TailEnd.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "ln" ) )
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
			Logger.logErr( "ln.parseOOXML: " + e.toString() );
		}
		Ln l = new Ln( attrs, fill, join, dash, h, t );
		return l;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:ln" );
		// attributes
		Iterator<String> i = attrs.keySet().iterator();
		while( i.hasNext() )
		{
			String key = (String) i.next();
			String val = (String) attrs.get( key );
			ooxml.append( " " + key + "=\"" + val + "\"" );
		}
		ooxml.append( ">" );
		if( fill != null )
		{
			ooxml.append( fill.getOOXML() );
		}
		if( dash != null )
		{
			ooxml.append( dash.getOOXML() );
		}
		if( join != null )
		{
			ooxml.append( join.getOOXML() );
		}
		if( h != null )
		{
			ooxml.append( h.getOOXML() );
		}
		if( t != null )
		{
			ooxml.append( t.getOOXML() );
		}
		ooxml.append( "</a:ln>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new Ln( this );
	}

	/**
	 * set the line width
	 *
	 * @param w
	 */
	public void setWidth( int w )
	{
		if( attrs == null )
		{
			attrs = new HashMap<String, String>();
		}
		attrs.put( "w", String.valueOf( w ) );

	}

	/**
	 * return line width in 2003-style units, 0 for none
	 *
	 * @return
	 */
	public int getWidth()
	{
		if( attrs != null && attrs.get( "w" ) != null )
		// Specifies the width to be used for the underline stroke.
		// If this attribute is omitted, then a value of 0 is assumed.
		{
			return Integer.valueOf( (String) attrs.get( "w" ) ) / 12700; // 1 pt = 12700 EMUs.
		}
		return 0;    // default
	}

	public int getColor()
	{
		if( fill == null )
		{
			return -1;
		}
		return fill.getColor();
	}

	public void setColor( String clr )
	{
		if( fill == null )
		{
			this.fill = new FillGroup( null, null, null, null, new SolidFill() );
		}
		fill.setColor( clr );
	}

	/**
	 * convert 2007 preset dasing scheme to 2003 line stype
	 * <br>0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none,
	 * 6= dk gray pattern, 7= med. gray, 8= light gray
	 *
	 * @return
	 */
	public int getLineStyle()
	{
		if( dash != null )
		{
			String style = dash.getPresetDashingScheme();
			if( style.equals( "solid" ) )
			{
				return 0;    // solid
			}
			else if( style.equals( "dash" ) ||
					style.equals( "sysDash" ) ||
					style.equals( "lgDash" ) )
			{
				return 1;
			}
			else if( style.equals( "sysDot" ) )
			{
				return 2;
			}
			else if( style.equals( "dashDot" ) ||
					style.equals( "sysDashDot" ) ||
					style.equals( "lgDashDot" ) )
			{
				return 3;
			}
			else if( style.equals( "sysDashDashDot" ) || style.equals( "lgDashDotDot" ) )
			{
				return 4;
			}
		}
		// ? none = 5 ???
		return 0;    // solid
	}
}

/**
 * line dash group:  prstDash, custDash
 */
// TODO: Finish custDash
class DashGroup implements OOXMLElement
{

	private static final long serialVersionUID = -6892326040716070609L;
	private PrstDash prstDash;

	public DashGroup( PrstDash p )
	{
		this.prstDash = p;
	}

	public DashGroup( DashGroup d )
	{
		this.prstDash = d.prstDash;
	}

	public static DashGroup parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		PrstDash p = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if(/*tnm.equals("custDash") ||*/
							tnm.equals( "prstDash" ) )
					{
						p = (PrstDash) PrstDash.parseOOXML( xpp, lastTag );
						break;
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "custDash" ) || endTag.equals( "prstDash" ) )
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
			Logger.logErr( "custDash.parseOOXML: " + e.toString() );
		}
		DashGroup dg = new DashGroup( p );
		return dg;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( prstDash != null )
		{
			ooxml.append( prstDash.getOOXML() );
		}
		// TODO: finish custDash
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new DashGroup( this );
	}

	/**
	 * returns the preset dashing scheme for the line, if any
	 * <br>One of:
	 * <br>dash, dashDot, lgDash, lgDashDot, lgDashDotDot, solid, sysDash, sysDashDot,
	 * sysDashDotDot, sysDot
	 *
	 * @return
	 */
	public String getPresetDashingScheme()
	{
		if( prstDash != null )
		{
			return prstDash.getPresetDashingScheme();
		}
		return null;
	}
}

/**
 * choice of miter, bevel or round
 * <p/>
 * parent: ln
 */
/* since each child element is so simple, we will just store which element it is rather than create separate object */
class JoinGroup implements OOXMLElement
{

	private static final long serialVersionUID = -6107424300366896696L;
	private boolean miter;
	private boolean round;
	private boolean bevel;
	private String miterVal;

	public JoinGroup( String a, boolean m, boolean r, boolean b )
	{
		this.miterVal = a;
		this.miter = m;
		this.round = r;
		this.bevel = b;
	}

	public JoinGroup( JoinGroup j )
	{
		this.miterVal = j.miterVal;
		this.miter = j.miter;
		this.round = j.round;
		this.bevel = j.bevel;
	}

	public static JoinGroup parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		boolean miter = false;
		boolean round = false;
		boolean bevel = false;
		String a = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "bevel" ) )
					{
						bevel = true;
					}
					else if( tnm.equals( "round" ) )
					{
						round = true;
					}
					else if( tnm.equals( "miter" ) )
					{
						miter = true;
						a = xpp.getAttributeValue( 0 );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "bevel" ) ||
							endTag.equals( "round" ) ||
							endTag.equals( "miter" ) )
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
			Logger.logErr( "JoinGroup.parseOOXML: " + e.toString() );
		}
		JoinGroup jg = new JoinGroup( a, miter, round, bevel );
		return jg;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		if( miter )
		{
			ooxml.append( "<a:miter lim=\"" + miterVal + "\"/>" );
		}
		else if( round )
		{
			ooxml.append( "<a:round/>" );
		}
		else if( bevel )
		{
			ooxml.append( "<a:bevel/>" );
		}
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new JoinGroup( this );
	}
}

/**
 * headEnd (Line Head/End Style)
 * <p/>
 * This element specifies decorations which can be added to the head of a line.
 */
class HeadEnd implements OOXMLElement
{
	private static final long serialVersionUID = -6744308104003922477L;
	private String len = null;
	private String type = null;
	private String w = null;

	public HeadEnd()
	{

	}

	public HeadEnd( String len, String type, String w )
	{
		this.len = len;
		this.type = type;
		this.w = w;
	}

	public HeadEnd( HeadEnd te )
	{
		this.len = te.len;
		this.type = te.type;
		this.w = te.w;
	}

	public static HeadEnd parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String len = null;
		String type = null;
		String w = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "headEnd" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "len" ) )
							{
								len = xpp.getAttributeValue( i );
							}
							else if( n.equals( "type" ) )
							{
								type = xpp.getAttributeValue( i );
							}
							else if( n.equals( "w" ) )
							{
								w = xpp.getAttributeValue( i );
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "headEnd" ) )
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
			Logger.logErr( "headEnd.parseOOXML: " + e.toString() );
		}
		HeadEnd te = new HeadEnd( len, type, w );
		return te;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:headEnd" );
		// attributes
		if( len != null )
		{
			ooxml.append( " len=\"" + len + "\"" );
		}
		if( type != null )
		{
			ooxml.append( " type=\"" + type + "\"" );
		}
		if( w != null )
		{
			ooxml.append( " w=\"" + w + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new HeadEnd( this );
	}
}

/**
 * tailEnd (Tail line end style)
 * <p/>
 * This element specifies decorations which can be added to the tail of a line.
 */
class TailEnd implements OOXMLElement
{
	private static final long serialVersionUID = -5587427916156543370L;
	private String len = null;
	private String type = null;
	private String w = null;

	public TailEnd()
	{
	}

	public TailEnd( String len, String type, String w )
	{
		this.len = len;
		this.type = type;
		this.w = w;
	}

	public TailEnd( TailEnd te )
	{
		this.len = te.len;
		this.type = te.type;
		this.w = te.w;
	}

	public static TailEnd parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String len = null;
		String type = null;
		String w = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "tailEnd" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "len" ) )
							{
								len = xpp.getAttributeValue( i );
							}
							else if( n.equals( "type" ) )
							{
								type = xpp.getAttributeValue( i );
							}
							else if( n.equals( "w" ) )
							{
								w = xpp.getAttributeValue( i );
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "tailEnd" ) )
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
			Logger.logErr( "tailEnd.parseOOXML: " + e.toString() );
		}
		TailEnd te = new TailEnd( len, type, w );
		return te;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:tailEnd" );
		// attributes
		if( len != null )
		{
			ooxml.append( " len=\"" + len + "\"" );
		}
		if( type != null )
		{
			ooxml.append( " type=\"" + type + "\"" );
		}
		if( w != null )
		{
			ooxml.append( " w=\"" + w + "\"" );
		}
		ooxml.append( "/>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new TailEnd( this );
	}
}

class PrstDash implements OOXMLElement
{

	private static final long serialVersionUID = -4645986946936173151L;
	private String val = null;

	public PrstDash()
	{
	}

	public PrstDash( String val )
	{
		this.val = val;
	}

	public PrstDash( PrstDash p )
	{
		this.val = p.val;
	}

	public static PrstDash parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		String val = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "prstDash" ) )
					{        // get val attribute
						val = xpp.getAttributeValue( 0 );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "prstDash" ) )
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
			Logger.logErr( "prstDash.parseOOXML: " + e.toString() );
		}
		PrstDash p = new PrstDash( val );
		return p;
	}

	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:prstDash val=\"" + val + "\"/>" );
		return ooxml.toString();
	}

	public OOXMLElement cloneElement()
	{
		return new PrstDash( this );
	}

	/**
	 * returns the preset dashing scheme  for the line, if any
	 * <br>One of:
	 * <br>dash, dashDot, lgDash, lgDashDot, lgDashDotDot, solid, sysDash, sysDashDot,
	 * sysDashDotDot, sysDot
	 *
	 * @return
	 */
	public String getPresetDashingScheme()
	{
		return val;
	}
}
