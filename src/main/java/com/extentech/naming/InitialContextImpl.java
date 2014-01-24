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
package com.extentech.naming;

import javax.naming.Context;
import javax.naming.Name;
import javax.naming.NameParser;
import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import java.util.Hashtable;

/**
 * A basic JNDI Context which holds a flat lookup of names
 */
public class InitialContextImpl implements javax.naming.Context
{

	NameParser nameParser = new NameParserImpl();
	protected Hashtable env;

	// provide persistence between instantiations
	public static String CONTEXT_ID = "com.extentech.naming.InitialContextImpl_instance";
	public static String LOAD_CONTEXT = "com.extentech.naming.load_context";

	public InitialContextImpl()
	{
		if( System.getProperties().get( CONTEXT_ID ) != null )
		{
			env = (Hashtable) System.getProperties().get( CONTEXT_ID );
		}
		else
		{
			String loadme = System.getProperty( LOAD_CONTEXT );
			env = new Hashtable();    // 20070518 KSC: Moved so gets init even if no LOAD_CONTEXT
			if( loadme != null )
			{
				if( loadme.equals( "true" ) )
				{
					//env = new Hashtable(); KSC: See above
					// this breaks properties
					System.getProperties().put( CONTEXT_ID, env );
				}
			}
		}
	}

	// check return... -jm
	@Override
	public Object addToEnvironment( String propName, Object propVal ) throws NamingException
	{
		if( env.contains( propVal ) )
		{
			throw new NamingException( "Object " + propName + " already exists in NamingContext." );
		}
		env.put( propName, propVal );
		return propVal;
	}

	// we use string to bind -- is that bad?
	@Override
	public void bind( Name name, Object obj ) throws NamingException
	{
		String str = name.toString();
		bind( str, obj );
	}

	@Override
	public void bind( String name, Object obj ) throws NamingException
	{
		try
		{
			addToEnvironment( name, obj );
		}
		catch( NamingException e )
		{
			env.remove( obj );
			env.put( name, obj ); // override
		}
	}

	private boolean closed = false;

	@Override
	public void close() throws NamingException
	{
		closed = true;
	}

	// ?
	@Override
	public Name composeName( Name name, Name prefix ) throws NamingException
	{
		NameImpl retval = new NameImpl();
		retval.addAll( prefix );
		retval.addAll( name );
		return retval;
	}

	@Override
	public String composeName( String name, String prefix ) throws NamingException
	{
		StringBuffer sb = new StringBuffer();
		sb.append( name );
		sb.append( prefix );
		return sb.toString();
	}

	@Override
	public Hashtable getEnvironment() throws NamingException
	{
		return env;
	}

	@Override
	public NameParser getNameParser( String name ) throws NamingException
	{
		nameParser.parse( name );
		return nameParser;
	}

	@Override
	public NameParser getNameParser( Name name ) throws NamingException
	{
		return nameParser;
	}

	@Override
	public Object lookup( Name name ) throws NamingException
	{
		return env.get( name );
	}

	@Override
	public Object lookup( String name ) throws NamingException
	{
		return env.get( name );
	}

	@Override
	public Object lookupLink( Name name ) throws NamingException
	{
		return env.get( name );
	}

	@Override
	public Object lookupLink( String name ) throws NamingException
	{
		return env.get( name );
	}

	@Override
	public void rebind( Name name, Object obj ) throws NamingException
	{
		bind( name, obj );
	}

	@Override
	public void rebind( String name, Object obj ) throws NamingException
	{
		bind( name, obj );
	}

	@Override
	public Object removeFromEnvironment( String propName ) throws NamingException
	{
		return env.remove( propName );
	}

	@Override
	public void rename( String oldName, String newName ) throws NamingException
	{
		Object ob = env.get( oldName );
		env.remove( oldName );
		env.put( newName, ob );
	}

	@Override
	public void rename( Name oldName, Name newName ) throws NamingException
	{
		Object ob = env.get( oldName );
		env.remove( oldName );
		env.put( newName, ob );
	}

	@Override
	public void unbind( String name ) throws NamingException
	{
		try
		{
			env.remove( env.get( name ) );
		}
		catch( Exception e )
		{
			throw new NamingException( e.toString() );
		}
	}

	@Override
	public void unbind( Name name ) throws NamingException
	{
		try
		{
			env.remove( env.get( name ) );
		}
		catch( Exception e )
		{
			throw new NamingException( e.toString() );
		}
	}

// TODO: Implement the following mehods -jm 9/27/2004

	@Override
	public NamingEnumeration list( String name ) throws NamingException
	{
		return null;
	}

	@Override
	public NamingEnumeration list( Name name ) throws NamingException
	{
		return null;
	}

	@Override
	public NamingEnumeration listBindings( Name name ) throws NamingException
	{
		return null;
	}

	@Override
	public NamingEnumeration listBindings( String name ) throws NamingException
	{
		return null;
	}

	@Override
	public Context createSubcontext( Name name ) throws NamingException
	{
		// This method is derived from interface javax.naming.Context
		// to do: code goes here
		return null;
	}

	@Override
	public Context createSubcontext( String name ) throws NamingException
	{
		// This method is derived from interface javax.naming.Context
		// to do: code goes here
		return null;
	}

	@Override
	public void destroySubcontext( String name ) throws NamingException
	{
		// This method is derived from interface javax.naming.Context
		// to do: code goes here
	}

	@Override
	public void destroySubcontext( Name name ) throws NamingException
	{
		// This method is derived from interface javax.naming.Context
		// to do: code goes here
	}

	@Override
	public String getNameInNamespace() throws NamingException
	{
		// This method is derived from interface javax.naming.Context
		// to do: code goes here
		return null;
	}
}