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
 * grpSp (Group Shape)
 * This element specifies a group shape that represents many shapes grouped together.
 * Within a group shape 1 each of the shapes that make up the group are
 * specified just as they normally would. The idea behind grouping elements however is that a
 * single transform can apply to many shapes at the same time.
 * <p/>
 * parents: absoluteAnchor, grpSp, oneCellAnchor, twoCellAnchor
 * children: nvGrpSpPr (req), grpSpPr (req), choice of: (sp, grpSp, graphicFrame, cxnSp, pic) 0 to unbounded
 */
public class GrpSp implements OOXMLElement
{

	private static final long serialVersionUID = -3276180769601314853L;
	private NvGrpSpPr nvpr = null;
	private GrpSpPr sppr = null;
	private ArrayList<OOXMLElement> choice = null;

	public GrpSp( NvGrpSpPr nvpr, GrpSpPr sppr, ArrayList<OOXMLElement> choice )
	{
		this.nvpr = nvpr;
		this.sppr = sppr;
		this.choice = choice;
	}

	public GrpSp( GrpSp g )
	{
		this.nvpr = g.nvpr;
		this.sppr = g.sppr;
		this.choice = g.choice;
	}

	/**
	 * parse grpSp element OOXML
	 *
	 * @param xpp
	 * @param lastTag
	 * @return
	 */
	public static GrpSp parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		NvGrpSpPr nvpr = null;
		GrpSpPr sppr = null;
		ArrayList<OOXMLElement> choice = new ArrayList<OOXMLElement>();
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "nvGrpSpPr" ) )
					{
						lastTag.push( tnm );
						nvpr = NvGrpSpPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "grpSpPr" ) )
					{
						lastTag.push( tnm );
						sppr = GrpSpPr.parseOOXML( xpp, lastTag, bk );
						// choice of: 0 or more below:
					}
					else if( tnm.equals( "sp" ) )
					{
						lastTag.push( tnm );
						choice.add( Sp.parseOOXML( xpp, lastTag, bk ) );
					}
					else if( tnm.equals( "grpSp" ) )
					{
						if( nvpr != null )
						{ // if not the initial start attribute, this is a child
							lastTag.push( tnm );
							choice.add( GrpSp.parseOOXML( xpp, lastTag, bk ) );
						}
					}
					else if( tnm.equals( "graphicFrame" ) )
					{
						lastTag.push( tnm );
						choice.add( GraphicFrame.parseOOXML( xpp, lastTag ) );
					}
					else if( tnm.equals( "cxnSp" ) )
					{
						lastTag.push( tnm );
						choice.add( CxnSp.parseOOXML( xpp, lastTag, bk ) );
					}
					else if( tnm.equals( "pic" ) )
					{
						lastTag.push( tnm );
						choice.add( Pic.parseOOXML( xpp, lastTag, bk ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "grpSp" ) )
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
			Logger.logErr( "GrpSp.parseOOXML: " + e.toString() );
		}
		GrpSp gf = new GrpSp( nvpr, sppr, choice );
		return gf;
	}

	/**
	 * return grpSp element OOXML
	 */
	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:grpSp>" );
		ooxml.append( nvpr.getOOXML() );
		ooxml.append( sppr.getOOXML() );
		if( choice != null )
		{
			for( OOXMLElement aChoice : choice )
			{
				ooxml.append( aChoice.getOOXML() );
			}
		}
		ooxml.append( "</xdr:grpSp>" );
		return ooxml.toString();
	}

	static int SP = 0;
	static int PIC = 1;
	static int CXN = 2;
	static int GRAPHICFRAME = 3;

	private OOXMLElement getObject( int type )
	{
		if( choice != null )
		{
			for( OOXMLElement oe : choice )
			{
				if( (oe instanceof Sp) && (type == SP) )
				{
					return oe;
				}
				if( (oe instanceof Pic) && (type == PIC) )
				{
					return oe;
				}
				if( (oe instanceof CxnSp) && (type == CXN) )
				{
					return oe;
				}
				if( (oe instanceof GraphicFrame) && (type == GRAPHICFRAME) )
				{
					return oe;
				}
				if( oe instanceof GrpSp )
				{
					return ((GrpSp) oe).getObject( type );
				}

			}
		}
		return null;
	}

	/**
	 * get name attribute of the group shape
	 *
	 * @return
	 */
	public String getName()
	{
		return nvpr.getName();
	}

	/**
	 * set name attribute of the group shape
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		nvpr.setName( name );    // set the group name
	}

	/**
	 * get macro attribute of the group shape
	 *
	 * @return
	 */
	public String getMacro()
	{
		String macro = null;
		// Have a shape with an embed?
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			macro = ((Sp) oe).getMacro();
		}
		// how's about a picture?
		if( macro == null )
		{
			oe = getObject( PIC );
		}
		if( (oe != null) && (macro == null) )
		{
			macro = ((Pic) oe).getMacro();
		}
		// or a connection shape?
		if( macro == null )
		{
			oe = getObject( CXN );
		}
		if( (oe != null) && (macro == null) )
		{
			macro = ((CxnSp) oe).getMacro();
		}
		return macro;
	}

	/**
	 * set macro attribute of the group shape
	 *
	 * @param name
	 */
	public void setMacro( String macro )
	{
		// Have a shape with a macro?
		String m = null;
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			m = ((Sp) oe).getMacro();
		}
		if( m != null )
		{
			((Sp) oe).setMacro( macro );
			return;
		}
		// how's about a picture?
		oe = getObject( PIC );
		if( oe != null )
		{
			m = ((Pic) oe).getMacro();
		}
		if( m != null )
		{
			((Pic) oe).setMacro( macro );
			return;
		}
		// or a connection shape?
		oe = getObject( CXN );
		if( oe != null )
		{
			m = ((CxnSp) oe).getMacro();
		}
		if( m != null )
		{
			((CxnSp) oe).setMacro( macro );
		}
	}

	/**
	 * get cthe descr attribute of the group shape
	 *
	 * @return
	 */
	public String getDescr()
	{
		return nvpr.getDescr();
	}

	/**
	 * set description attribute of the group shape
	 * sometimes associated with shape name
	 *
	 * @param descr
	 */
	public void setDescr( String descr )
	{
		nvpr.setDescr( descr );
	}

	/**
	 * return the rid for the embedded object (picture or chart or picture shape) (i.e. resides within the file)
	 *
	 * @return
	 */
	public String getEmbed()
	{
		String embed = null;
		// Have a shape with an embed?
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			embed = ((Sp) oe).getEmbed();
		}
		// how's about a picture?
		if( embed == null )
		{
			oe = getObject( PIC );
		}
		if( (oe != null) && (embed == null) )
		{
			embed = ((Pic) oe).getEmbed();
		}
		// or a connection shape?
		if( embed == null )
		{
			oe = getObject( CXN );
		}
		if( (oe != null) && (embed == null) )
		{
			embed = ((CxnSp) oe).getEmbed();
		}
		return embed;
	}

	/**
	 * set the rid for this object (picture or chart or picture shape) (resides within the file)
	 *
	 * @param embed
	 */
	public void setEmbed( String embed )
	{
		// Have a shape with an embed?
		String e = null;
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			e = ((Sp) oe).getEmbed();
		}
		if( e != null )
		{
			((Sp) oe).setEmbed( embed );
			return;
		}
		// how's about a picture?
		oe = getObject( PIC );
		if( oe != null )
		{
			e = ((Pic) oe).getEmbed();
		}
		if( e != null )
		{
			((Pic) oe).setEmbed( embed );
			return;
		}
		// or a connection shape?
		oe = getObject( CXN );
		if( oe != null )
		{
			e = ((CxnSp) oe).getEmbed();
		}
		if( e != null )
		{
			((CxnSp) oe).setEmbed( embed );
		}
	}

	/**
	 * return the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public String getLink()
	{
		String link = null;
		// Have a shape with an embed?
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			link = ((Sp) oe).getLink();
		}
		// how's about a picture?
		if( link == null )
		{
			oe = getObject( PIC );
		}
		if( (oe != null) && (link == null) )
		{
			link = ((Pic) oe).getLink();
		}
		// or a connection shape?
		if( link == null )
		{
			oe = getObject( CXN );
		}
		if( (oe != null) && (link == null) )
		{
			link = ((CxnSp) oe).getLink();
		}
		return link;
	}

	/**
	 * rset the id for the linked picture (i.e. doesn't reside in the file)
	 *
	 * @return
	 */
	public void setLink( String link )
	{
		String l = null;
		// Have a shape with an embed?
		OOXMLElement oe = getObject( SP );
		if( oe != null )
		{
			l = ((Sp) oe).getLink();
		}
		if( l != null )
		{
			((Sp) oe).setLink( link );
			return;
		}
		// how's about a picture?
		oe = getObject( PIC );
		if( oe != null )
		{
			l = ((Pic) oe).getLink();
		}
		if( l != null )
		{
			((Pic) oe).setLink( link );
			return;
		}
		// or a connection shape?
		oe = getObject( CXN );
		if( oe != null )
		{
			l = ((CxnSp) oe).getLink();
		}
		if( l != null )
		{
			((CxnSp) oe).setLink( link );
		}
	}

	/**
	 * return the rid of the chart element, if exists
	 *
	 * @return
	 */
	public String getChartRId()
	{
		OOXMLElement oe = getObject( GRAPHICFRAME );
		if( oe != null )
		{
			return ((GraphicFrame) oe).getChartRId();
		}
		return null;
	}

	/**
	 * set the URI associated with this graphic data
	 *
	 * @param uri
	 */
	public void setURI( String uri )
	{
		OOXMLElement oe = getObject( GRAPHICFRAME );
		((GraphicFrame) oe).setURI( uri );
	}

	/**
	 * utility to return the shape properties element for this picture
	 * should be depreciated when OOXML is completely distinct from BIFF8
	 *
	 * @return
	 */
	public SpPr getSppr()
	{
		OOXMLElement oe = getObject( PIC );
		if( oe != null )
		{
			return ((Pic) oe).getSppr();
		}
		return null;

	}

	/**
	 * return if this Group refers to a shape, as opposed a chart or an image
	 *
	 * @return
	 */
	public boolean hasShape()
	{
		return ((getObject( SP ) != null) || (getObject( CXN ) != null));
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{

		if( choice != null )
		{
			// get the first??
			OOXMLElement oe = choice.get( 0 );
			if( oe instanceof Sp )
			{
				return ((Sp) oe).getId();
			}
			if( oe instanceof Pic )
			{
				return ((Pic) oe).getId();
			}
			if( oe instanceof GraphicFrame )
			{
				return ((GraphicFrame) oe).getId();
			}
			if( oe instanceof CxnSp )
			{
				return ((CxnSp) oe).getId();
			}
			if( oe instanceof GrpSp )
			{
				return ((GrpSp) oe).getId();
			}
		}
		return -1;
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( choice != null )
		{
			// get the first??
			OOXMLElement oe = choice.get( 0 );
			if( oe instanceof Sp )
			{
				((Sp) oe).setId( id );
			}
			else if( oe instanceof Pic )
			{
				((Pic) oe).setId( id );
			}
			else if( oe instanceof GraphicFrame )
			{
				((GraphicFrame) oe).setId( id );
			}
			else if( oe instanceof CxnSp )
			{
				((CxnSp) oe).setId( id );
			}
			else if( oe instanceof GrpSp )
			{
				((GrpSp) oe).setId( id );
			}
		}
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GrpSp( this );
	}
}

/**
 * nvGrpSpPr 1 (Non-Visual Properties for a Group Shape)
 */
class NvGrpSpPr implements OOXMLElement
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4404072268706949318L;
	CNvPr cp = null;
	CNvGrpSpPr cgrpsppr = null;

	public NvGrpSpPr( CNvPr cp, CNvGrpSpPr cgrpsppr )
	{
		this.cp = cp;
		this.cgrpsppr = cgrpsppr;
	}

	public NvGrpSpPr( NvGrpSpPr g )
	{
		this.cp = g.cp;
		this.cgrpsppr = g.cgrpsppr;
	}

	public static NvGrpSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		CNvPr cp = null;
		CNvGrpSpPr cgrpsppr = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "cNvPr" ) )
					{
						lastTag.push( tnm );
						cp = (CNvPr) CNvPr.parseOOXML( xpp, lastTag );
					}
					else if( tnm.equals( "cNvGrpSpPr" ) )
					{
						lastTag.push( tnm );
						cgrpsppr = CNvGrpSpPr.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "nvGrpSpPr" ) )
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
			Logger.logErr( "NvGrpSpPr.parseOOXML: " + e.toString() );
		}
		NvGrpSpPr grpsppr = new NvGrpSpPr( cp, cgrpsppr );
		return grpsppr;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:nvGrpSpPr>" );
		ooxml.append( cp.getOOXML() );
		ooxml.append( cgrpsppr.getOOXML() );
		ooxml.append( "</xdr:nvGrpSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new NvGrpSpPr( this );
	}

	/**
	 * get cNvPr name attribute
	 *
	 * @return
	 */
	public String getName()
	{
		if( cp != null )
		{
			return cp.getName();
		}
		return null;
	}

	/**
	 * set cNvPr name attribute
	 *
	 * @param name
	 */
	public void setName( String name )
	{
		if( cp != null )
		{
			cp.setName( name );
		}
	}

	/**
	 * get cNvPr descr attribute
	 *
	 * @return
	 */
	public String getDescr()
	{
		if( cp != null )
		{
			return cp.getDescr();
		}
		return null;
	}

	/**
	 * set cNvPr desc attribute
	 *
	 * @param name
	 */
	public void setDescr( String descr )
	{
		if( cp != null )
		{
			cp.setDescr( descr );
		}
	}

	/**
	 * set the cNvPr id for this element
	 *
	 * @param id
	 */
	public void setId( int id )
	{
		if( cp != null )
		{
			cp.setId( id );
		}
	}

	/**
	 * return the cNvPr id for this element
	 *
	 * @return
	 */
	public int getId()
	{
		if( cp != null )
		{
			return cp.getId();
		}
		return -1;
	}

}

/**
 * grpSpPr (Group Shape Properties)
 * This element specifies the properties that are to be common across all of the shapes
 * within the corresponding group. If there are any conflicting properties within the
 * group shape properties and the individual shape properties then the individual shape properties
 * should take precedence.
 * <p/>
 * children:  xfrm, fillproperties (group), effectproperties (group), scene3d, extLst (all optional)
 */
// TODO: FINISH scene3d, extLst
class GrpSpPr implements OOXMLElement
{

	private static final long serialVersionUID = 7464871024304781512L;
	private Xfrm xf = null;
	private String bwmode = null;
	private FillGroup fill = null;
	private EffectPropsGroup effect = null;

	public GrpSpPr( Xfrm xf, String bwmode, FillGroup fill, EffectPropsGroup effect )
	{
		this.xf = xf;
		this.bwmode = bwmode;
		this.fill = fill;
		this.effect = effect;
	}

	public GrpSpPr( GrpSpPr g )
	{
		this.xf = g.xf;
		this.bwmode = g.bwmode;
		this.fill = g.fill;
		this.effect = g.effect;
	}

	public static GrpSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		Xfrm xf = null;
		String bwmode = null;
		FillGroup fill = null;
		EffectPropsGroup effect = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "grpSpPr" ) )
					{        // get attributes
						if( xpp.getAttributeCount() == 1 )
						{
							bwmode = xpp.getAttributeValue( 0 );
						}
					}
					else if( tnm.equals( "xfrm" ) )
					{
						lastTag.push( tnm );
						xf = (Xfrm) Xfrm.parseOOXML( xpp, lastTag );
						// scene3d, extLst- finish
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
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "grpSpPr" ) )
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
			Logger.logErr( "GrpSpPr.parseOOXML: " + e.toString() );
		}
		GrpSpPr g = new GrpSpPr( xf, bwmode, fill, effect );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:grpSpPr" );
		if( bwmode != null )
		{
			// TODO: Finish
			ooxml.append( " bwMode=\"" + bwmode + "\"" );
		}
		ooxml.append( ">" );
		if( xf != null )
		{
			ooxml.append( xf.getOOXML() );
		}
		if( fill != null )
		{
			ooxml.append( fill.getOOXML() );
		}
		if( effect != null )
		{
			ooxml.append( effect.getOOXML() );
		}
		ooxml.append( "</xdr:grpSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new GrpSpPr( this );
	}
}

/**
 * cNvGrpSpPr (Non-Visual Group Shape Drawing Properties)
 * <p/>
 * parent:		nvGrpSpPr
 * children: 	grpSpLocks, extLst
 */
class CNvGrpSpPr implements OOXMLElement
{

	private static final long serialVersionUID = -1106010927060582127L;
	private GrpSpLocks gsl = null;

	public CNvGrpSpPr( GrpSpLocks gsl )
	{
		this.gsl = gsl;
	}

	public CNvGrpSpPr( CNvGrpSpPr g )
	{
		this.gsl = g.gsl;
	}

	public static CNvGrpSpPr parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		GrpSpLocks gsl = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "grpSpLocks" ) )
					{
						lastTag.push( tnm );
						gsl = GrpSpLocks.parseOOXML( xpp, lastTag );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "cNvGrpSpPr" ) )
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
			Logger.logErr( "CNvGrpSpPr.parseOOXML: " + e.toString() );
		}
		CNvGrpSpPr c = new CNvGrpSpPr( gsl );
		return c;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:cNvGrpSpPr>" );
		if( gsl != null )
		{
			ooxml.append( gsl.getOOXML() );
		}
		ooxml.append( "</xdr:cNvGrpSpPr>" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new CNvGrpSpPr( this );
	}
}

/**
 * grpSpLocks (Group Shape Locks)
 * This element specifies all locking properties for a connection shape. These properties inform the generating
 * application about specific properties that have been previously locked and thus should not be changed.
 * <p/>
 * parent:		cNvGrpSpPr
 * attributes:  many, all optional
 * children: 	extLst
 */
class GrpSpLocks implements OOXMLElement
{

	private static final long serialVersionUID = -2592038952923879415L;
	private HashMap<String, String> attrs;

	public GrpSpLocks( HashMap<String, String> attrs )
	{
		this.attrs = attrs;
	}

	public GrpSpLocks( GrpSpLocks g )
	{
		this.attrs = g.attrs;
	}

	public static GrpSpLocks parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
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
					if( tnm.equals( "grpSpLocks" ) )
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
					if( endTag.equals( "grpSpLocks" ) )
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
			Logger.logErr( "CNvGrpSpPr.parseOOXML: " + e.toString() );
		}
		GrpSpLocks g = new GrpSpLocks( attrs );
		return g;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<xdr:grpSpLocks>" );
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
		return new GrpSpLocks( this );
	}
}


