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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Iterator;
import java.util.NoSuchElementException;
import java.util.Vector;

/**
 * a Vector class designed to provide forwards compatibility
 * for JDK1.1 programs.
 */
public class CompatibleVector extends Vector
{
	private static final Logger log = LoggerFactory.getLogger( CompatibleVector.class );
	private static final long serialVersionUID = 6805047965683753637L;
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

	public CompatibleVector()
	{
		super();
	}

	public CompatibleVector( int i )
	{
		super( i );
	}

	@Override
	public Iterator iterator()
	{
		return new Itr();
	}

	int hits = 0, misses = 0;

	/**
	 * if the object being checked implements index
	 * hints, the lookup can be performed much faster
	 * <p/>
	 * speed of lookups is affected by 'shuffling' positions
	 * of vector elements.
	 */
	public int indexOf( CompatibleVectorHints r )
	{
		//  return super.indexOf(r);

		int x = r.getRecordIndexHint();
		if( (x > 0) && (x < super.size()) )
		{
			if( super.elementAt( x ) != null )
			{
				if( super.elementAt( x ).equals( r ) )
				{
					//Logger.logInfo("hit/miss="+ hits++ + ":"+ misses);
					//   hits++;
					return x;
				}
			}
		}
		x -= change_offset;
		if( (x > 0) && (x < super.size()) )
		{
			if( super.elementAt( x ).equals( r ) )
			{

				//Logger.logInfo("hit/miss="+ hits++ + ":"+ misses);
				// hits++;
				return x;
			}
		}
		int t = -1;
		if( change_offset > reindex_change_size )
		{
			this.resetHints( false );
		}
		// if(x>0)t = super.indexOf(r,x);
		if( x > 0 )
		{
			t = super.indexOf( r );
		}
		if( t < 0 )
		{
			t = super.indexOf( r );
		}
		r.setRecordIndexHint( t );
		// Logger.logInfo("hit/miss="+ hits++ + ":"+ misses++);
		return t;
	   /**/
	   /*
        int idx = r.getRecordIndexHint();
        int retval = -1;
        int recsz = this.size();
        if(idx >= recsz)idx = recsz-1;
        if(idx  < 0){
            idx = this.indexOf((Object)r);
            r.setRecordIndexHint(idx);
            return idx;
        }
        if(this.get(idx).equals(r))return idx;
        else{
            boolean found = false, checkhi = true, checklo = true;
              int hi =idx, lo = idx;
          // int hi = idx + change_offset, lo = idx + change_offset;
            //Logger.logInfo(change_offset);
            if(hi < 0)hi = recsz/2;
            if(lo > 0)lo = recsz/2;
            while(!found && (checkhi || checklo)){
                if(++hi >= recsz)checkhi = false;
                if(--lo < 0)checklo = false;
                if(checkhi){
                    Object b = this.get(hi);
                    if(b.equals(r)){
                        found = true;
                        retval = hi;
                      // Logger.logInfo(" hi:" + (idx - hi));
                    }else if(checklo){
                        Object c = this.get(lo);
                        if(c.equals(r)){
                            found = true;
                            retval = lo;
                     // Logger.logInfo("lo:" + (idx - lo));
                        }
                    }
                }
            }
        }
        Logger.logInfo(" hi:" + (idx - hi));
        Logger.logInfo("lo:" + (idx - lo));
        if((retval == -1)&&(change_offset > 0)){
            this.resetHints(false);
           // Logger.logInfo("Loopy!! " + r);
            retval = this.indexOf(r);
        }
        r.setRecordIndexHint(retval);
        return retval;
        */
	}
    
/* */
	/**
	 * Index of element to be returned by subsequent call to next.
	 */
	int cursor = 0;

	/**
	 * Index of element returned by most recent call to next or
	 * previous.  Reset to -1 if this element is deleted by a call
	 * to remove.
	 */
	int lastRet = -1;

	/**
	 * overriding AbstractList so we can have concurrent mods...
	 */
	public Object next()
	{
		//       checkForComodification();
		try
		{
			Object next = get( cursor );
			lastRet = cursor++;
			return next;
		}
		catch( IndexOutOfBoundsException e )
		{
			throw new NoSuchElementException();
		}
	}

	@Override
	public Object get( int idx )
	{
		return super.elementAt( idx );
	}

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
				Double dd = (Double) super.elementAt( i );
				if( dd > d )
				{
					super.insertElementAt( obj, i );
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
		return super.elementAt( super.size() - 1 );
	}

	public boolean add( CompatibleVectorHints obj )
	{
		this.change_offset++; //
		int idx = super.size();
		if( obj != null )
		{
			obj.setRecordIndexHint( idx );
		}
		try
		{
			super.insertElementAt( obj, idx );
			return true;
		}
		catch( Exception e )
		{
			return false;
		}
	}

	public void add( int idx, CompatibleVectorHints obj )
	{
		this.change_offset++;//
		obj.setRecordIndexHint( idx );
		super.insertElementAt( obj, idx );
	}

	public void addAll( CompatibleVector cv )
	{
		for( Object b : cv )
		{
			if( b instanceof CompatibleVectorHints )
			{
				this.add( (CompatibleVectorHints) b );
			}
			else
			{
				this.add( b );
			}
		}
	}

	@Override
	public boolean remove( Object obj )
	{
		if( super.remove( obj ) )
		{
			this.change_offset--;
			return true;
		}
		return false;
	}

	@Override
	public void clear()
	{
		this.change_offset = 0;
		super.removeAllElements();
	}

	@Override
	public Object[] toArray()
	{
		Object[] obj = new Object[super.size()];
		for( int i = 0; i < super.size(); i++ )
		{
			obj[i] = super.elementAt( i );
		}
		return obj;
	}

	@Override
	public Object[] toArray( Object[] obj )
	{
		for( int i = 0; i < super.size(); i++ )
		{
			try
			{
				obj[i] = super.elementAt( i );
			}
			catch( Exception e )
			{
				log.error( "CompatibleVector.toArray() failed.", e );
			}
		}
		return obj;
	}

	private class Itr implements Iterator
	{
		/**
		 * Index of element to be returned by subsequent call to next.
		 */
		int cursor = 0;

		/**
		 * Index of element returned by most recent call to next or
		 * previous.  Reset to -1 if this element is deleted by a call
		 * to remove.
		 */
		int lastRet = -1;

		/**
		 * The modCount value that the iterator believes that the backing
		 * List should have.  If this expectation is violated, the iterator
		 * has detected concurrent modification.
		 */
		int expectedModCount = modCount;

		@Override
		public boolean hasNext()
		{
			return cursor != size();
		}

		@Override
		public Object next()
		{
			Object next = get( cursor );
			lastRet = cursor++;
			return next;
		}

		@Override
		public void remove()
		{
			if( lastRet == -1 )
			{
				throw new IllegalStateException();
			}

			CompatibleVector.this.remove( lastRet );
			if( lastRet < cursor )
			{
				cursor--;
			}
			lastRet = -1;
			expectedModCount = modCount;
		}

	}
}