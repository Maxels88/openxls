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

import java.util.ArrayList;
import java.util.Enumeration;

/**
 * a Vector class designed to provide forwards compatibility
 * for JDK1.1 programs.
 */
public class FastGetVector extends ArrayList
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6701901995748359720L;
	private int change_offset = 0;
	private int reindex_change_size = 1000;

	/**
	 * reset the hints for all vector elements
	 * expense is linear to size but will increase
	 * accuracy of subsequent 'indexOf' calls.
	 */
	public void resetHints( boolean ignore_records )
	{
		// ExcelTools.benchmark("Re-indexing CompatibleVector" + reindex_change_size++);
		if( !ignore_records )
		{
			for( int t = 0; t < this.size(); t++ )
			{
				try
				{
					((CompatibleVectorHints) this.get( t )).setRecordIndexHint( t );
				}
				catch( Exception e )
				{
					return;
				}
			}
		}
		change_offset = 0;
	}

	public FastGetVector()
	{
		super();
	}

	public FastGetVector( int i )
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
		double d = obj;
		try
		{
			for( int i = 0; i < super.size(); i++ )
			{
				Double dd = (Double) super.get( i );
				if( dd > d )
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
		this.change_offset++;// 
		obj.setRecordIndexHint( idx );
		super.add( idx, obj );
	}

	/*
		public void addAll(CompatibleVector cv){
			for(int i=0;i<cv.size();i++){
				Object b = cv.get(i);
				if(b instanceof CompatibleVectorHints){
					this.add((CompatibleVectorHints)b);
				}else{
					this.add((Object)b);
				}
			}
		}
	*/
	@Override
	public boolean remove( Object obj )
	{
		this.change_offset--;
		if( super.remove( obj ) )
		{
			return true;
		}
		return false;
	}

	@Override
	public void clear()
	{
		this.change_offset = 0;
		super.clear();
	}

	@Override
	public Object[] toArray()
	{
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

	public void copyInto( Object[] obar )
	{
		for( int x = 0; x < obar.length; x++ )
		{
			super.add( obar );
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
		FastGetVectorEnumerator cve = new FastGetVectorEnumerator( this );
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

}

final class FastGetVectorEnumerator implements Enumeration
{

	private FastGetVector it = null;
	int x = 0;

	FastGetVectorEnumerator( FastGetVector itx )
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
