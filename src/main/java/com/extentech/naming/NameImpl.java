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

import com.extentech.toolkit.CompatibleVector;

import javax.naming.InvalidNameException;
import javax.naming.Name;
import java.util.Enumeration;

/*	
	Name add(int posn, String comp) 
			  Adds a single component at a specified position within this name. 
	 Name add(String comp) 
			  Adds a single component to the end of this name. 
	 Name addAll(int posn, Name n) 
			  
	 Name addAll(Name suffix) 
			  Adds the components of a name -- in order -- to the end of this name. 
	 Object clone() 
			  Generates a new copy of this name. 
	 int compareTo(Object obj) 
			  Compares this name with another name for order. 
	 boolean endsWith(Name n) 
			  Determines whether this name ends with a specified suffix. 
	 String get(int posn) 
			  Retrieves a component of this name. 
	 Enumeration getAll() 
			  Retrieves the components of this name as an enumeration of strings. 
	 boolean isEmpty() 
			  Determines whether this name is empty. 
	 Object remove(int posn) 
			  Removes a component from this name. 
	 int size() 
			  Returns the number of components in this name. 
	 boolean startsWith(Name n) 
			  Determines whether this name starts with a specified prefix. 
*/

public class NameImpl implements Name
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4387233472850688497L;
	CompatibleVector vals = new CompatibleVector();	/* (non-Javadoc)
	
	  
	 * * @see javax.naming.Name#clone()
	 */

	@Override
	public Object clone()
	{
		NameImpl nimple = new NameImpl();
		CompatibleVector newvals = new CompatibleVector();
		newvals.addAll( vals );
		nimple.vals = newvals;
		return nimple;
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#remove(int)
	 */
	@Override
	public Object remove( int arg0 ) throws InvalidNameException
	{
		return vals.remove( arg0 );
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#get(int)
	 */
	@Override
	public String get( int arg0 )
	{
		return vals.get( arg0 ).toString();
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#getAll()
	 */
	@Override
	public Enumeration getAll()
	{
		return vals.elements();
	}

	/* Creates a name whose components consist of a prefix of the components of this name. 
	 * @see javax.naming.Name#getPrefix(int)
	 */
	@Override
	public Name getPrefix( int arg0 )
	{
		return null;
	}

	/* Creates a name whose components consist of a suffix of the components in this name. 
	 * @see javax.naming.Name#getSuffix(int)
	 */
	@Override
	public Name getSuffix( int arg0 )
	{
		// TODO Auto-generated method stub
		return null;
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#add(java.lang.String)
	 */
	@Override
	public Name add( String arg0 ) throws InvalidNameException
	{
		return null;
	}

	/* Adds the components of a name -- in order -- at a specified position within this name. 
	 * @see javax.naming.Name#addAll(int, javax.naming.Name)
	 */
	@Override
	public Name addAll( int arg0, Name arg1 ) throws InvalidNameException
	{
		this.vals.addAll( arg0, ((NameImpl) arg1).getVals() );
		return this;
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#addAll(javax.naming.Name)
	 */
	@Override
	public Name addAll( Name arg0 ) throws InvalidNameException
	{
		this.vals.addAll( ((NameImpl) arg0).getVals() );
		return this;
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#size()
	 */
	@Override
	public int size()
	{
		return vals.size();
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#isEmpty()
	 */
	@Override
	public boolean isEmpty()
	{
		return vals.size() > 0;
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#compareTo(java.lang.Object)
	 */
	@Override
	public int compareTo( Object arg0 )
	{
		return this.compareTo( arg0 );
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#endsWith(javax.naming.Name)
	 */
	@Override
	public boolean endsWith( Name arg0 )
	{
		Object ob1 = arg0.get( arg0.size() - 1 );
		Object ob2 = this.get( this.size() - 1 );
		return ob1.equals( ob2 );
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#startsWith(javax.naming.Name)
	 */
	@Override
	public boolean startsWith( Name arg0 )
	{
		Object ob1 = arg0.get( 0 );
		Object ob2 = this.get( 0 );
		return ob1.equals( ob2 );
	}

	/* (non-Javadoc)
	 * @see javax.naming.Name#add(int, java.lang.String)
	 */
	@Override
	public Name add( int arg0, String arg1 ) throws InvalidNameException
	{
		vals.set( arg0, arg1 );
		return this;
	}

	/**
	 * @return
	 */
	CompatibleVector getVals()
	{
		return vals;
	}

	/**
	 * @param vector
	 */
	void setVals( CompatibleVector vector )
	{
		vals = vector;
	}

}