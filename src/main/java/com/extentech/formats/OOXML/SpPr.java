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

import java.util.Stack;

/**
 * spPr
 * This OOXML/DrawingML element specifies the visual shape properties that can be applied to a shape. These properties include the
 * shape fill, outline, geometry, effects, and 3D orientation.
 * <p/>
 * parents:  cxnSp, pic, sp
 * children:
 * blipFill (Picture Fill) §5.1.10.14
 * custGeom (Custom Geometry) §5.1.11.8
 * effectDag (Effect Container) §5.1.10.25
 * effectLst (Effect Container) §5.1.10.26
 * extLst (Extension List) §5.1.2.1.15
 * gradFill (Gradient Fill) §5.1.10.33
 * grpFill (Group Fill) §5.1.10.35
 * ln (Outline) §5.1.2.1.24
 * noFill (No Fill) §5.1.10.44
 * pattFill (Pattern Fill) §5.1.10.47
 * prstGeom (Preset geometry) §5.1.11.18
 * scene3d (3D Scene Properties) §5.1.4.1.26
 * solidFill (Solid Fill) §5.1.10.54
 * sp3d (Apply 3D shape properties) §5.1.7.12
 * xfrm (2D Transform)
 */
public class SpPr implements OOXMLElement
{

	private static final long serialVersionUID = 4542844402486023785L;
	private Xfrm x = null;
	private GeomGroup geom = null;
	private FillGroup fill = null;
	private Ln l = null;
	private EffectPropsGroup effect = null;
	// scene3d, sp3d
	String bwMode = null;
	// namespace
	private String ns = null;

	public SpPr( String ns )
	{    // no-param constructor, set up common defaults
		x = new Xfrm();
		x.setNS( "a" );
		this.ns = ns;
		geom = new GeomGroup( new PrstGeom(), null );
		bwMode = "auto";
	}

	public SpPr( Xfrm x, GeomGroup geom, FillGroup fill, Ln l, EffectPropsGroup effect, String bwMode, String ns )
	{
		this.x = x;
		this.geom = geom;
		this.fill = fill;
		this.l = l;
		this.effect = effect;
		this.bwMode = bwMode;
		this.ns = ns;
	}

	public SpPr( SpPr clone )
	{
		this.x = clone.x;
		this.geom = clone.geom;
		this.fill = clone.fill;
		this.l = clone.l;
		this.effect = clone.effect;
		this.bwMode = clone.bwMode;
		this.ns = clone.ns;
	}

	/**
	 * create a default shape property with a solid fill and a line
	 *
	 * @param solidfill
	 * @param w
	 * @param lnClr
	 */
	public SpPr( String ns, String solidfill, int w, String lnClr )
	{
		this.ns = ns;
		if( solidfill != null )
		{
			this.fill = new FillGroup( null, null, null, null, new SolidFill( solidfill ) );
		}
		this.l = new Ln( w, lnClr );
	}

	/**
	 * set the namespace for spPr element
	 *
	 * @param ns
	 */
	public void setNS( String ns )
	{
		this.ns = ns;
	}

	/**
	 * parse shape OOXML element spPr
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spPr object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		Xfrm x = null;
		FillGroup fill = null;
		Ln l = null;
		EffectPropsGroup effect = null;
		GeomGroup geom = null;
		// scene3d, sp3d
		String bwMode = null;
		String ns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "spPr" ) )
					{
						ns = xpp.getPrefix();
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "bwMode" ) )
							{
								bwMode = xpp.getAttributeValue( i );
							}
						}
					}
					else if( tnm.equals( "xfrm" ) )
					{
						lastTag.push( tnm );
						x = (Xfrm) Xfrm.parseOOXML( xpp, lastTag );
						//x.setNS("a");
					}
					else if( tnm.equals( "ln" ) )
					{
						lastTag.push( tnm );
						l = (Ln) Ln.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "prstGeom" ) || tnm.equals( "custGeom" ) )
					{        // GEOMETRY GROUP
						lastTag.push( tnm );
						geom = (GeomGroup) GeomGroup.parseOOXML( xpp, lastTag );
					}
					else if(                                // FILL GROUP
							tnm.equals( "solidFill" ) ||
									tnm.equals( "noFill" ) ||
									tnm.equals( "gradFill" ) ||
									tnm.equals( "grpFill" ) ||
									tnm.equals( "pattFill" ) ||
									tnm.equals( "blipFill" ) )
					{
						lastTag.push( tnm );
						fill = (FillGroup) FillGroup.parseOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "extLst" ) )
					{
						lastTag.push( tnm );
						ExtLst.parseOOXML( xpp, lastTag ); // ignore for now TODO: FINISH
					}
					else if(                                // EFFECT GROUP
							tnm.equals( "effectLst" ) || tnm.equals( "effectDag" ) )
					{
						lastTag.push( tnm );
						effect = (EffectPropsGroup) EffectPropsGroup.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "spPr" ) )
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
			Logger.logErr( "spPr.parseOOXML: " + e.toString() );
		}
		SpPr sp = new SpPr( x, geom, fill, l, effect, bwMode, ns );
		return sp;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<" + this.ns + ":spPr" );
		if( bwMode != null )
		{
			ooxml.append( " bwMode=\"" + bwMode + "\">" );
		}
		else
		{
			ooxml.append( ">" );
		}
		if( x != null )
		{
			ooxml.append( x.getOOXML() );    // must pass namespace to xfrm
		}
		if( geom != null )
		{
			ooxml.append( geom.getOOXML() );    // geometry choice
		}
		if( fill != null )
		{
			ooxml.append( fill.getOOXML() );    // fill choice
		}
		if( l != null )
		{
			ooxml.append( l.getOOXML() );        // ln element
		}
		if( effect != null )
		{
			ooxml.append( effect.getOOXML() );    // effect properties choice
		}
		// scene3d, sp3d
		ooxml.append( "</" + ns + ":spPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new SpPr( this );
	}

	/**
	 * return the id for the embedded picture (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		if( fill != null )
		{
			return fill.getEmbed();
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
		if( fill != null )
		{
			return fill.getLink();
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
		if( fill != null )
		{
			fill.setEmbed( embed );
		}
	}

	/**
	 * set the link attribute for this blip (the id for the linked picture)
	 *
	 * @param embed
	 */
	public void setLink( String link )
	{
		if( fill != null )
		{
			fill.setLink( link );
		}
	}

	/**
	 * returns the ln element of this shape property set, if any
	 *
	 * @return
	 */
	public Ln getLn()
	{
		return l;
	}

	/**
	 * add a line for this shape property
	 *
	 * @param w   line width
	 * @param clr html color string
	 */
	public void setLine( int w, String clr )
	{
		l = new Ln();
		l.setWidth( w );
		l.setColor( clr );
	}

	/**
	 * returns the width of the line of this shape proeprty,
	 * or -1 if no line is present
	 *
	 * @return
	 */
	public int getLineWidth()
	{
		if( l != null )
		{
			return l.getWidth();
		}
		return -1;
	}

	/**
	 * returns the color of the line of this shape property, or -1 if no line is present
	 *
	 * @return
	 */
	public int getLineColor()
	{
		if( l != null )
		{
			return l.getColor();
		}
		return -1;
	}

	/**
	 * returns the line style of the line of this shape property, or -1 if no line is present
	 *
	 * @return
	 */
	public int getLineStyle()
	{
		if( l != null )
		{
			return l.getLineStyle();
		}
		return -1;
	}

	/**
	 * return the fill color
	 *
	 * @return
	 */
	public int getColor()
	{
		if( fill != null )
		{
			return fill.getColor();
		}
		return 1;
	}

	/**
	 * returns the fill of this shape property set, if any
	 *
	 * @return
	 */
	public FillGroup getFill()
	{
		return fill;
	}

	/**
	 * remove the line, if any, for this shape property
	 */
	public void removeLine()
	{
		l = null;
	}

	/**
	 * return true if this shape properties contains a line
	 */
	public boolean hasLine()
	{
		return (l != null);
	}
}

