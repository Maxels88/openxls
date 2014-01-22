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
package docs.samples;

import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;

import java.io.File;
import java.lang.reflect.Method;

/**
 * Convenience program to run ExtenXLS tests
 */
public class RunSamples
{

	/**
	 * Run all tests in the test Suite
	 * <p/>
	 * Jan 18, 2010
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{

		String inf = System.getProperty( "user.dir" ) + "/docs/samples/";
		try
		{
			Logger.logInfo( "RunSamples Begin..." );
			RunSamples.execDir( new File( inf ) );
			Logger.logInfo( "RunSamples Complete." );
		}
		catch( Exception e )
		{
			Logger.logErr( "RunSamples failed.", e );
		}

	}

	/**
	 * Recursively executes all of the main methods in all of the classes
	 * in a directory and subdirectories.
	 * <p/>
	 * Jan 18, 2010
	 *
	 * @param f
	 */
	private static void execDir( File f )
	{
		if( f.isDirectory() )
		{
			String[] children = f.list();
			for( int i = 0; i < children.length; i++ )
			{
				execDir( new File( f, children[i] ) );
			}
		}
		else if( f.getName().indexOf( "RunSamples" ) == -1 )
		{
			String fn = f.getName();
			String packagename = f.getPath();
			packagename = StringTool.replaceChars( "\\", packagename, "." ); // win
			packagename = StringTool.replaceChars( "/", packagename, "." ); // unix

			packagename = packagename.substring( packagename.indexOf( "docs." ) );
			if( fn.indexOf( ".class" ) > -1 )
			{
				try
				{
					packagename = packagename.substring( 0, packagename.indexOf( ".class" ) );
					Class cx = Class.forName( packagename );
					Method[] m = cx.getMethods();
					for( int t = 0; t < m.length; t++ )
					{
						if( m[t].getName().equals( "main" ) )
						{
							String[] args = new String[1];
							try
							{
								Logger.logInfo( "Running " + packagename );
								m[t].invoke( cx, args );
							}
							catch( Throwable tx )
							{
								Logger.logErr( "Problem Running " + packagename + " main method.", tx );
							}
						}
					}
				}
				catch( Exception e )
				{
					Logger.logErr( "Cannot run " + packagename + " main method.", e );
				}
			}
		}
	}
}