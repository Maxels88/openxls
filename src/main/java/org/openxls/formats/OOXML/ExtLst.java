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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.HashMap;
import java.util.Stack;

// TODO: FINISH
public class ExtLst implements OOXMLElement
{
	private static final Logger log = LoggerFactory.getLogger( ExtLst.class );
	private static final long serialVersionUID = -4122012942547055359L;
	private HashMap<String, String> attrs = null;
	private String nameSpace = null;

	public ExtLst()
	{
		attrs = new HashMap<>();
		attrs.put( "cx", new String( "0" ) );
		attrs.put( "cy", new String( "0" ) );
	}

	public ExtLst( HashMap<String, String> attrs, String ns )
	{
		this.attrs = attrs;
		nameSpace = ns;
	}

	public ExtLst( ExtLst e )
	{
		attrs = e.attrs;
		nameSpace = e.nameSpace;
	}

	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		HashMap<String, String> attrs = new HashMap<>();
		String ns = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "extLst" ) )
					{        // get attributes
						ns = xpp.getPrefix();
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							attrs.put( xpp.getAttributeName( i ), xpp.getAttributeValue( i ) );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "extLst" ) )
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
			log.error( "extLst.parseOOXML: " + e.toString() );
		}
		Ext e = new Ext( attrs, ns );
		return e;
	}

	/**
	 * set the namespace for ext element
	 *
	 * @param ns
	 */
	public void setNamespace( String ns )
	{
		nameSpace = ns;
	}

	@Override
	public String getOOXML()
	{
		//TODO: FINISH
		return "";
		/*
		StringBuffer ooxml= new StringBuffer();	
    	ooxml.append("<" + this.ns + ":ext");
    	// attributes
    	Iterator i= attrs.keySet().iterator();
    	while (i.hasNext()) {
    		String key= (String) i.next();
    		String val= (String) attrs.get(key);
    		ooxml.append(" " + key + "=\"" + val + "\"");
    	}
    	ooxml.append("/>");
    	return ooxml.toString();
    	*/
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new ExtLst( this );
	}
}
	
