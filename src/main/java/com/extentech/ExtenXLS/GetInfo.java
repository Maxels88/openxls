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
package com.extentech.ExtenXLS;

import java.io.IOException;
import java.util.Properties;

/**
 * Prints various bits of information about this build.
 */
public class GetInfo
{

	private static Properties buildProps;

	private static final void loadBuildProps()
	{
		if( null == buildProps )
		{
			buildProps = new Properties();
			try
			{
				buildProps.load( GetInfo.class.getResourceAsStream( "build.properties" ) );
			}
			catch( IOException caught )
			{
				// don't care, it'll just be empty
			}
		}
	}

	private static String version;

	/**
	 * Return the version number of this installation of ExtenXLS.
	 */
	public static final String getVersion()
	{
		if( null == version )
		{
			loadBuildProps();
			version = buildProps.getProperty( "version", "unknown" );

			if( version.endsWith( "-SNAPSHOT" ) )
			{
				version = version.substring( 0, version.length() - 9 );

				String timestamp = buildProps.getProperty( "timestamp" );
				String commit = buildProps.getProperty( "commit" );
				if( "UNKNOWN".equals( commit ) )
				{
					commit = null;
				}

				if( null != timestamp )
				{
					version += "-" + timestamp;
				}
				else if( null != commit )
				{
					version += "-unknown";
				}

				if( null != commit )
				{
					version += "-" + commit.substring( 0, 10 );
				}
			}
		}

		return version;
	}

	public static void main( String[] args )
	{
		// Product name and version
		System.out.println( "This is ExtenXLS " + getVersion() );
	}
}
