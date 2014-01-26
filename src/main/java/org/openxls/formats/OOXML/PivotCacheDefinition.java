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
package org.openxls.formats.OOXML;

import org.openxls.ExtenXLS.WorkBookHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserFactory;

import java.io.InputStream;

public class PivotCacheDefinition implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( PivotCacheDefinition.class );
	private static final long serialVersionUID = -5070227633357072878L;
	private String ref = null;
	private String sheet = null;
	private int icache;

	@Override
	public OOXMLElement cloneElement()
	{
		return null;
	}

	public PivotCacheDefinition( String ref, String sheet, int icache )
	{
		this.ref = ref;
		this.sheet = sheet;
		this.icache = icache;
	}

	public static PivotCacheDefinition parseOOXML( WorkBookHandle bk, String cacheid, InputStream ii )
	{
		String ref = null;
		String sheet = null;
		int icache = 1;
		try
		{
			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();
			xpp.setInput( ii, null ); // using XML 1.0 specification
			int eventType = xpp.getEventType();

			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "pivotCacheDefinition" ) )
					{ // get attributes
						//  r:id="rId1" refreshedBy="Kaia" refreshedDate="41038.467970833335" createdVersion="1" refreshedVersion="3" recordCount="4" upgradeOnRefresh="1">
					}
					else if( tnm.equals( "cacheSource" ) )
					{
						for( int z = 0; z < xpp.getAttributeCount(); z++ )
						{
							String nm = xpp.getAttributeName( z );
							String v = xpp.getAttributeValue( z );
							if( nm.equals( "type" ) )
							{
								if( !v.equals( "worksheet" ) )
								{
									// consolidation, external, scenario --
									log.warn( "PivotCacheDefinition: Data Souce " + v + " Not Supported" );
									return null;
								}
							}
						}
					}
					else if( tnm.equals( "worksheetSource" ) )
					{
						// ref, sheet, id (sheet rid), name (range)
						for( int z = 0; z < xpp.getAttributeCount(); z++ )
						{
							String nm = xpp.getAttributeName( z );
							String v = xpp.getAttributeValue( z );
							if( nm.equals( "ref" ) )
							{
								ref = v;
							}
							else if( nm.equals( "sheet" ) )
							{
								sheet = v;
							}
							else if( nm.equals( "name" ) )
							{
								ref = v;
							}
							else if( nm.equals( "id" ) )
							{

							}
						}
					}
					else if( tnm.equals( "cacheFields" ) )
					{
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{ // go to end of file
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "PivotCacheDefinition.parseOOXML: " + e.toString() );
		}

		if( cacheid != null )
		{
// KSC: TESTING!!!			icache= Integer.valueOf((String)cacheid)+1;
			icache = 1;
		}
		icache = bk.getWorkBook().addPivotStream( ref, sheet, icache );
		return new PivotCacheDefinition( ref, sheet, icache );
	}

	/**
	 * return the pivot cache id
	 */
	public int getICache()
	{
		return icache;
	}

	/**
	 * returns the data source reference
	 *
	 * @return
	 */
	public String getRef()
	{
		return ref;
	}

	/**
	 * return the sheet the data reference is on
	 *
	 * @return
	 */
	public String getSheet()
	{
		return sheet;
	}

	@Override
	public String getOOXML()
	{
		// TODO: Finish
		return null;
	}

}
