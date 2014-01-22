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
package com.extentech.toolkit;

/* Based on Bare Bones Browser Launcher by Dem Pilafian,
 * which is in the public domain. You may find it at:
 * http://www.centerkey.com/java/browser/
 */

/**
 * Launches the user's default browser to display a web page.
 *
 * @author Dem Pilafian
 * @author Sam Hanes
 */
public class BrowserLauncher
{
	/**
	 * List of potential browsers on systems without a default mechanism.
	 */
	public static final String[] browsers = {
			"google-chrome", "firefox", "opera", "konqueror", "epiphany", "seamonkey", "galeon", "kazehakase", "mozilla"
	};

	/**
	 * The browser that was last successfully run.
	 */
	private static String browser = null;

	/**
	 * Opens the specified web page in the user's default browser.
	 *
	 * @param url the URL of the page to be opened
	 * @throws Exception if an error occurred attempting to launch the browser.
	 *                   If the browser is successfully started but later fails for
	 *                   some reason, no exception will be thrown.
	 */
	public static void open( String url ) throws Exception
	{
		// Attempt to use the Desktop class from JDK 1.6+ (even if on 1.5)
		// This uses reflection to mimic the call:
		// java.awt.Desktop.getDesktop().browse( java.net.URI.create(url) );
		try
		{
			Class desktop = Class.forName( "java.awt.Desktop" );
			desktop.getDeclaredMethod( "browse", new Class[]{ java.net.URI.class } ).invoke( desktop.getDeclaredMethod( "getDesktop",
			                                                                                                            (Class[]) null )
			                                                                                        .invoke( null, (Object[]) null ),
			                                                                                 new Object[]{ java.net.URI.create( url ) } );

			// If that didn't throw an exception, we're done
			return;
		}
		catch( ClassNotFoundException e )
		{
			// Intentionally empty, falls back to platform-dependent code
		}
		catch( NoSuchMethodException e )
		{
			// Intentionally empty, falls back to platform-dependent code
		}
		catch( Exception e )
		{
			throw new Exception( "failed to launch browser", e );
		}

		String osName = System.getProperty( "os.name" );
		try
		{
			// If this is OS X, use the FileManager class
			if( osName.startsWith( "Mac OS" ) )
			{
				Class.forName( "com.apple.eio.FileManager" ).getDeclaredMethod( "openURL", new Class[]{ String.class } ).invoke( null,
				                                                                                                                 new Object[]{
						                                                                                                                 url
				                                                                                                                 } );
			}

			// If this is Windows, call the FileProtocolHandler via rundll
			else if( osName.startsWith( "Windows" ) )
			{
				Runtime.getRuntime().exec( "rundll32 url.dll,FileProtocolHandler " + url );
			}

			// Otherwise, assume this is a POSIX-like system and
			// start trying possible browser commands
			else
			{
				// If we haven't found a browser yet, try some possible ones
				if( browser == null )
				{
					for( String browser1 : browsers )
					{
						if( Runtime.getRuntime().exec( new String[]{ "which", browser1 } ).waitFor() == 0 )
						{
							browser = browser1;
						}
					}

					// If we couldn't find one, throw an exception
					if( browser == null )
					{
						throw new Exception( "no browser found" );
					}
				}

				// Call the browser with the URL
				Runtime.getRuntime().exec( new String[]{ browser, url } );
			}
		}
		catch( Exception e )
		{
			throw new Exception( "failed to launch browser", e );
		}
	}
}
