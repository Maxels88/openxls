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
package org.openxls.toolkit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExtenClassLoader extends java.lang.ClassLoader
{
	private static final Logger log = LoggerFactory.getLogger( ExtenClassLoader.class );
	private String targetClassName;

	public ExtenClassLoader( String target )
	{
		targetClassName = target;
	}

	public ExtenClassLoader()
	{
	}

	private String wd = "";

	public void setDirectory( String _wd )
	{
		wd = _wd;
	}

	protected byte[] loadClassFromFile( String name )
	{
		byte[] classBytes = null;
		try
		{
			File file = null;
			FileInputStream stream = null;
			//  name = name.substring(name.indexOf(wd)+wd.length()); // strip the working directory
			name = wd + "/" + name;
			name = StringTool.replaceChars( ".", name, "/" );
			file = new File( name + ".class" );
			classBytes = new byte[(int) file.length()];
			stream = new FileInputStream( file );
			stream.read( classBytes );
			stream.close();
		}
		catch( IOException io )
		{
		}
		return classBytes;
	}

	@Override
	public synchronized Class loadClass( String name ) throws ClassNotFoundException
	{
		return loadClass( name, false );
	}

	@Override
	public synchronized Class loadClass( String name, boolean resolve ) throws ClassNotFoundException
	{
		Class loadedClass;
		byte[] bytes;
		if( !name.equals( targetClassName ) )
		{
			try
			{
				loadedClass = super.findSystemClass( name );
				return loadedClass;
			}
			catch( ClassNotFoundException e )
			{
			}
		}
		bytes = loadClassFromFile( name );
		if( bytes == null )
		{
			throw new ClassNotFoundException();
		}
		loadedClass = defineClass( name, bytes, 0, bytes.length );
		if( loadedClass == null )
		{
			log.error( "Class cannot be loaded: " + name );
			throw new ClassFormatError();
		}
		if( resolve )
		{
			resolveClass( loadedClass );
		}

		return loadedClass;
	}
}