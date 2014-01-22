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

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Stack;
import java.util.Vector;

/**
 * A recycling cache, items are checked at intervals
 */
public abstract class GenericRecycleBin extends java.lang.Thread implements Map, com.extentech.toolkit.RecycleBin
{
	protected Map map = new java.util.HashMap();
	protected Vector active = new Vector();
	protected Stack spares = new Stack();

	/**
	 * add an item
	 */
	@Override
	public void addItem( Recyclable r ) throws RecycleBinFullException
	{
		if( (MAXITEMS == -1) || (map.size() < MAXITEMS) )
		{
			addItem( Integer.valueOf( map.size() ), r );
		}
		else
		{
			throw new RecycleBinFullException();
		}
	}

	/**
	 * returns number of items in cache
	 *
	 * @return
	 */
	public int getNumItems()
	{
		return active.size();
	}

	@Override
	public void addItem( Object key, Recyclable r ) throws RecycleBinFullException
	{
		// recycle();
		if( (MAXITEMS == -1) || (map.size() < MAXITEMS) )
		{
			active.add( r );
			map.put( key, r );

		}
		else
		{
			throw new RecycleBinFullException();
		}
	}

	/**
	 * iterate all active items and try to recycle
	 */
	public synchronized void recycle()
	{
		Recyclable[] rs = new Recyclable[active.size()];
		active.copyInto( rs );
		for( int t = 0; t < rs.length; t++ )
		{
			try
			{
				Recyclable rb = rs[t];
				if( !rb.inUse() )
				{
					// recycle it
					rb.recycle();

					// remove from active and lookup
					active.remove( rb );
					map.remove( rb );

					// put in spares
					spares.push( rb );

				}
			}
			catch( Exception ex )
			{
				Logger.logErr( "recycle failed", ex );
			}
		}

	}

	@Override
	public void empty()
	{
		map.clear();
		active.clear();
	}

	@Override
	public synchronized List getAll()
	{
		return active;
	}

	/**
	 * returns a new or recycled item from the spares pool
	 *
	 * @see com.extentech.toolkit.RecycleBin#getItem()
	 */
	@Override
	public synchronized Recyclable getItem() throws RecycleBinFullException
	{
		Recyclable active = null;
		// spares contains the recycled
		if( spares.size() > 0 )
		{
			active = (Recyclable) spares.pop();
			addItem( active );
			return active;
		}
		recycle();

		// technically infinite loop until exception thrown
		return getItem();
	}

	protected int MAXITEMS = -1; // no limit is default

	/**
	 * max number of items to be put
	 * in this bin.
	 */
	@Override
	public void setMaxItems( int i )
	{
		MAXITEMS = i;
	}

	public int getMaxItems()
	{
		return MAXITEMS;
	}

	public int getSpareCount()
	{
		return spares.size();
	}

	public GenericRecycleBin()
	{
	}

	@Override
	public void clear()
	{
		map.clear();
		active.clear();
	}

	@Override
	public boolean containsKey( Object key )
	{
		return map.containsKey( key );

	}

	@Override
	public boolean containsValue( Object value )
	{
		return map.containsValue( value );
	}

	@Override
	public Set entrySet()
	{
		return map.entrySet();
	}

	public boolean equals( Object o )
	{
		return map.equals( o );
	}

	@Override
	public Object get( Object key )
	{
		return map.get( key );
	}

	public int hashCode()
	{
		return map.hashCode();
	}

	@Override
	public boolean isEmpty()
	{
		return map.isEmpty();
	}

	@Override
	public Set keySet()
	{
		return map.keySet();
	}

	@Override
	public Object put( Object arg0, Object arg1 )
	{
		active.add( arg1 );
		return map.put( arg0, arg1 );
	}

	@Override
	public void putAll( Map arg0 )
	{
		active.addAll( arg0.entrySet() );
		map.putAll( arg0 );
	}

	@Override
	public Object remove( Object key )
	{
		active.remove( map.get( key ) );
		return map.remove( key );
	}

	@Override
	public int size()
	{
		return map.size();
	}

	public String toString()
	{
		return map.toString();
	}

	@Override
	public Collection values()
	{
		return map.values();
	}

	public java.util.Map getMap()
	{
		return map;
	}

	public void setMap( java.util.HashMap _map )
	{
		map = _map;
	}

}