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

import java.util.Enumeration;

/**
 * a Vector class designed to provide forwards compatibility
 * for JDK1.1 programs.
 * <p/>
 * <p/>
 * //	  						add; 		toArray; 	iterator; 	insert; 		get; 				indexOf; 	remove
 * //	  	TreeList =		1260	7360;		3080;  		160;   		170;				3400;  		170;
 * //	 	ArrayList =  	220		1480;		1760; 		6870;    	50;				1540; 		7200;
 * //		LinkedList =  	270		7360;		3350;		55860;		290720;		2910;		55200;
 */
public class FastAddVector extends SpecialArrayList implements java.io.Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5615615290731997512L;

	public FastAddVector()
	{
		super();
	}

	public FastAddVector( int i )
	{
		// super(i);
		super();
	}

	int hits = 0, misses = 0;

	/*
	 * If passed in a Double, it stores it as a Double in the vector
	 * in logical order (1,2,3)
	 * Returns true if all values can be represented by double and the element is inserted
	 */
	public boolean addOrderedDouble( Double obj )
	{
		double d = obj.doubleValue();
		try
		{
			for( int i = 0; i < super.size(); i++ )
			{
				Double dd = (Double) super.get( i );
				if( dd.doubleValue() > d )
				{
					super.add( i, obj );
					return true;
				}
			}
		}
		catch( Exception e )
		{
			return false;
		}
		super.add( obj );
		return true;
	}

	/*
	 * Returns the last object in the collection
	 * 
	 */
	public Object last()
	{
		return super.get( super.size() - 1 );
	}

	public void add( int idx, CompatibleVectorHints obj )
	{
		obj.setRecordIndexHint( idx );
		super.add( idx, obj );
	}

	@Override
	public boolean remove( Object obj )
	{
		if( super.remove( obj ) )
		{
			return true;
		}
		return false;
	}

	@Override
	public void clear()
	{
		super.clear();
	}

	@Override
	public Object[] toArray()
	{
		if( true )
		{
			return elementData;
		}
		Object[] obj = new Object[super.size()];
		for( int i = 0; i < super.size(); i++ )
		{
			obj[i] = super.get( i );
		}
		return obj;
	}

	public void removeAllElements()
	{
		super.clear();
	}

	/**
	 * @param obar
	 */
	public void copyInto( Object[] obar )
	{
		for( int x = 0; x < obar.length; x++ )
		{
			super.add( obar[x] );
		}
	}

	public void insertElementAt( Object ob, int i )
	{
		super.add( i, ob );
	}

	public Object lastElement()
	{
		return super.get( super.size() - 1 );
	}

	public Enumeration elements()
	{
		FastAddVectorEnumerator cve = new FastAddVectorEnumerator( this );
		return cve;
	}

	public Object elementAt( int t )
	{
		return super.get( t );
	}

	@Override
	public Object[] toArray( Object[] obj )
	{
		for( int i = 0; i < super.size(); i++ )
		{
			obj[i] = super.get( i );
		}
		return obj;
	}

	final class FastAddVectorEnumerator implements Enumeration
	{

		private FastAddVector it = null;
		int x = 0;

		FastAddVectorEnumerator( FastAddVector itx )
		{
			it = itx;
		}

		@Override
		public Object nextElement()
		{
			return it.elementAt( x++ );
		}

		@Override
		public boolean hasMoreElements()
		{
			return (x < it.size());
		}

	}

}


