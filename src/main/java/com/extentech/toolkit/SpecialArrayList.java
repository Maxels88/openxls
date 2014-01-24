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

import java.util.AbstractList;
import java.util.Collection;
import java.util.Collections;
import java.util.ConcurrentModificationException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.RandomAccess;
import java.util.Vector;

/**
 * Resizable-array implementation of the <tt>List</tt> interface.  Implements
 * all optional list operations, and permits all elements, including
 * <tt>null</tt>.  In addition to implementing the <tt>List</tt> interface,
 * this class provides methods to manipulate the size of the array that is
 * used internally to store the list.  (This class is roughly equivalent to
 * <tt>Vector</tt>, except that it is unsynchronized.)<p>
 * <p/>
 * The <tt>size</tt>, <tt>isEmpty</tt>, <tt>get</tt>, <tt>set</tt>,
 * <tt>iterator</tt>, and <tt>listIterator</tt> operations run in constant
 * time.  The <tt>add</tt> operation runs in <i>amortized constant time</i>,
 * that is, adding n elements requires O(n) time.  All of the other operations
 * run in linear time (roughly speaking).  The constant factor is low compared
 * to that for the <tt>LinkedList</tt> implementation.<p>
 * <p/>
 * Each <tt>ArrayList</tt> instance has a <i>capacity</i>.  The capacity is
 * the size of the array used to store the elements in the list.  It is always
 * at least as large as the list size.  As elements are added to an ArrayList,
 * its capacity grows automatically.  The details of the growth policy are not
 * specified beyond the fact that adding an element has constant amortized
 * time cost.<p>
 * <p/>
 * An application can increase the capacity of an <tt>ArrayList</tt> instance
 * before adding a large number of elements using the <tt>ensureCapacity</tt>
 * operation.  This may reduce the amount of incremental reallocation.<p>
 * <p/>
 * <strong>Note that this implementation is not synchronized.</strong> If
 * multiple threads access an <tt>ArrayList</tt> instance concurrently, and at
 * least one of the threads modifies the list structurally, it <i>must</i> be
 * synchronized externally.  (A structural modification is any operation that
 * adds or deletes one or more elements, or explicitly resizes the backing
 * array; merely setting the value of an element is not a structural
 * modification.)  This is typically accomplished by synchronizing on some
 * object that naturally encapsulates the list.  If no such object exists, the
 * list should be "wrapped" using the <tt>Collections.synchronizedList</tt>
 * method.  This is best done at creation time, to prevent accidental
 * unsynchronized access to the list:
 * <pre>
 *  List list = Collections.synchronizedList(new ArrayList(...));
 * </pre><p>
 *
 * The iterators returned by this class's <tt>iterator</tt> and
 * <tt>listIterator</tt> methods are <i>fail-fast</i>: if list is structurally
 * modified at any time after the iterator is created, in any way except
 * through the iterator's own remove or add methods, the iterator will throw a
 * ConcurrentModificationException.  Thus, in the face of concurrent
 * modification, the iterator fails quickly and cleanly, rather than risking
 * arbitrary, non-deterministic behavior at an undetermined time in the
 * future.<p>
 *
 * Note that the fail-fast behavior of an iterator cannot be guaranteed
 * as it is, generally speaking, impossible to make any hard guarantees in the
 * presence of unsynchronized concurrent modification.  Fail-fast iterators
 * throw <tt>ConcurrentModificationException</tt> on a best-effort basis.
 * Therefore, it would be wrong to write a program that depended on this
 * exception for its correctness: <i>the fail-fast behavior of iterators
 * should be used only to detect bugs.</i><p>
 *
 * This class is a member of the
 * <a href="{@docRoot}/../guide/collections/index.html">
 * Java Collections Framework</a>.
 *
 * @see Collection
 * @see List
 * @see LinkedList
 * @see Vector
 * @see Collections#synchronizedList(List)
 */

class SpecialArrayList extends AbstractList implements List, RandomAccess, Cloneable, java.io.Serializable
{
	private static final long serialVersionUID = 8683452581122892189L;

	/**
	 * Returns an iterator over the elements in this list in proper
	 * sequence. <p>
	 * <p/>
	 * This implementation returns a straightforward implementation of the
	 * iterator interface, relying on the backing list's <tt>size()</tt>,
	 * <tt>get(int)</tt>, and <tt>remove(int)</tt> methods.<p>
	 * <p/>
	 * Note that the iterator returned by this method will throw an
	 * <tt>UnsupportedOperationException</tt> in response to its
	 * <tt>remove</tt> method unless the list's <tt>remove(int)</tt> method is
	 * overridden.<p>
	 * <p/>
	 * This implementation can be made to throw runtime exceptions in the face
	 * of concurrent modification, as described in the specification for the
	 * (protected) <tt>modCount</tt> field.
	 *
	 * @return an iterator over the elements in this list in proper sequence.
	 * @see #modCount
	 */
	@Override
	public Iterator iterator()
	{
		return new Itr();
	}

	/**
	 * The array buffer into which the elements of the ArrayList are stored.
	 * The capacity of the ArrayList is the length of this array buffer.
	 */
	Object[] elementData;

	/**
	 * The size of the ArrayList (the number of elements it contains).
	 *
	 * @serial
	 */
	private int size;

	/**
	 * Constructs an empty list with the specified initial capacity.
	 *
	 * @param initialCapacity the initial capacity of the list.
	 * @throws IllegalArgumentException if the specified initial capacity
	 *                                  is negative
	 */
	public SpecialArrayList( int initialCapacity )
	{
		super();
		if( initialCapacity < 0 )
		{
			throw new IllegalArgumentException( "Illegal Capacity: " + initialCapacity );
		}
		elementData = new Object[initialCapacity];
	}

	/**
	 * Constructs an empty list with an initial capacity of ten.
	 */
	public SpecialArrayList()
	{
		this( 10 );
	}

	/**
	 * Constructs a list containing the elements of the specified
	 * collection, in the order they are returned by the collection's
	 * iterator.  The <tt>ArrayList</tt> instance has an initial capacity of
	 * 110% the size of the specified collection.
	 *
	 * @param c the collection whose elements are to be placed into this list.
	 * @throws NullPointerException if the specified collection is null.
	 */
	public SpecialArrayList( Collection c )
	{
		size = c.size();
		// Allow 10% room for growth
		elementData = new Object[(int) Math.min( (size * 110L) / 100, Integer.MAX_VALUE )];
		c.toArray( elementData );
	}

	/**
	 * Trims the capacity of this <tt>ArrayList</tt> instance to be the
	 * list's current size.  An application can use this operation to minimize
	 * the storage of an <tt>ArrayList</tt> instance.
	 */
	public void trimToSize()
	{
		modCount++;
		int oldCapacity = elementData.length;
		if( size < oldCapacity )
		{
			Object[] oldData = elementData;
			elementData = new Object[size];
			System.arraycopy( oldData, 0, elementData, 0, size );
		}
	}

	/**
	 * Increases the capacity of this <tt>ArrayList</tt> instance, if
	 * necessary, to ensure  that it can hold at least the number of elements
	 * specified by the minimum capacity argument.
	 *
	 * @param minCapacity the desired minimum capacity.
	 */
	public void ensureCapacity( int minCapacity )
	{
		modCount++;
		int oldCapacity = elementData.length;
		if( minCapacity > oldCapacity )
		{
			Object[] oldData = elementData;
			int newCapacity = ((oldCapacity * 3) / 2) + 1;
			if( newCapacity < minCapacity )
			{
				newCapacity = minCapacity;
			}
			elementData = new Object[newCapacity];
			System.arraycopy( oldData, 0, elementData, 0, size );
		}
	}

	/**
	 * Returns the number of elements in this list.
	 *
	 * @return the number of elements in this list.
	 */
	@Override
	public int size()
	{
		return size;
	}

	/**
	 * Tests if this list has no elements.
	 *
	 * @return <tt>true</tt> if this list has no elements;
	 * <tt>false</tt> otherwise.
	 */
	@Override
	public boolean isEmpty()
	{
		return size == 0;
	}

	/**
	 * Returns <tt>true</tt> if this list contains the specified element.
	 *
	 * @param elem element whose presence in this List is to be tested.
	 * @return <code>true</code> if the specified element is present;
	 * <code>false</code> otherwise.
	 */
	@Override
	public boolean contains( Object elem )
	{
		return indexOf( elem ) >= 0;
	}

	/**
	 * Searches for the first occurence of the given argument, testing
	 * for equality using the <tt>equals</tt> method.
	 *
	 * @param elem an object.
	 * @return the index of the first occurrence of the argument in this
	 * list; returns <tt>-1</tt> if the object is not found.
	 * @see Object#equals(Object)
	 */
	@Override
	public int indexOf( Object elem )
	{
		if( elem == null )
		{
			for( int i = 0; i < size; i++ )
			{
				if( elementData[i] == null )
				{
					return i;
				}
			}
		}
		else
		{
			for( int i = 0; i < size; i++ )
			{
				if( elem.equals( elementData[i] ) )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * Returns the index of the last occurrence of the specified object in
	 * this list.
	 *
	 * @param elem the desired element.
	 * @return the index of the last occurrence of the specified object in
	 * this list; returns -1 if the object is not found.
	 */
	@Override
	public int lastIndexOf( Object elem )
	{
		if( elem == null )
		{
			for( int i = size - 1; i >= 0; i-- )
			{
				if( elementData[i] == null )
				{
					return i;
				}
			}
		}
		else
		{
			for( int i = size - 1; i >= 0; i-- )
			{
				if( elem.equals( elementData[i] ) )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * Returns a shallow copy of this <tt>ArrayList</tt> instance.  (The
	 * elements themselves are not copied.)
	 *
	 * @return a clone of this <tt>ArrayList</tt> instance.
	 */
	@Override
	public Object clone()
	{
		try
		{
			SpecialArrayList v = (SpecialArrayList) super.clone();
			v.elementData = new Object[size];
			System.arraycopy( elementData, 0, v.elementData, 0, size );
			v.modCount = 0;
			return v;
		}
		catch( CloneNotSupportedException e )
		{
			// this shouldn't happen, since we are Cloneable
			throw new InternalError();
		}
	}

	/**
	 * Returns an array containing all of the elements in this list
	 * in the correct order.
	 *
	 * @return an array containing all of the elements in this list
	 * in the correct order.
	 */
	@Override
	public Object[] toArray()
	{
		Object[] result = new Object[size];
		System.arraycopy( elementData, 0, result, 0, size );
		return result;
	}

	/**
	 * Returns an array containing all of the elements in this list in the
	 * correct order; the runtime type of the returned array is that of the
	 * specified array.  If the list fits in the specified array, it is
	 * returned therein.  Otherwise, a new array is allocated with the runtime
	 * type of the specified array and the size of this list.<p>
	 * <p/>
	 * If the list fits in the specified array with room to spare (i.e., the
	 * array has more elements than the list), the element in the array
	 * immediately following the end of the collection is set to
	 * <tt>null</tt>.  This is useful in determining the length of the list
	 * <i>only</i> if the caller knows that the list does not contain any
	 * <tt>null</tt> elements.
	 *
	 * @param a the array into which the elements of the list are to
	 *          be stored, if it is big enough; otherwise, a new array of the
	 *          same runtime type is allocated for this purpose.
	 * @return an array containing the elements of the list.
	 * @throws ArrayStoreException if the runtime type of a is not a supertype
	 *                             of the runtime type of every element in this list.
	 */
	@Override
	public Object[] toArray( Object[] a )
	{
		if( a.length < size )
		{
			a = (Object[]) java.lang.reflect.Array.newInstance( a.getClass().getComponentType(), size );
		}

		System.arraycopy( elementData, 0, a, 0, size );

		if( a.length > size )
		{
			a[size] = null;
		}

		return a;
	}

	// Positional Access Operations

	/**
	 * Returns the element at the specified position in this list.
	 *
	 * @param index index of element to return.
	 * @return the element at the specified position in this list.
	 * @throws IndexOutOfBoundsException if index is out of range <tt>(index
	 *                                   &lt; 0 || index &gt;= size())</tt>.
	 */
	@Override
	public Object get( int index )
	{
		RangeCheck( index );
		return elementData[index];
	}

	/**
	 * Replaces the element at the specified position in this list with
	 * the specified element.
	 *
	 * @param index   index of element to replace.
	 * @param element element to be stored at the specified position.
	 * @return the element previously at the specified position.
	 * @throws IndexOutOfBoundsException if index out of range
	 *                                   <tt>(index &lt; 0 || index &gt;= size())</tt>.
	 */
	@Override
	public Object set( int index, Object element )
	{
		RangeCheck( index );

		Object oldValue = elementData[index];
		elementData[index] = element;
		return oldValue;
	}

	/**
	 * Appends the specified element to the end of this list.
	 *
	 * @param o element to be appended to this list.
	 * @return <tt>true</tt> (as per the general contract of Collection.add).
	 */
	@Override
	public boolean add( Object o )
	{
		ensureCapacity( size + 1 );  // Increments modCount!!
		elementData[size++] = o;
		return true;
	}

	/**
	 * Inserts the specified element at the specified position in this
	 * list. Shifts the element currently at that position (if any) and
	 * any subsequent elements to the right (adds one to their indices).
	 *
	 * @param index   index at which the specified element is to be inserted.
	 * @param element element to be inserted.
	 * @throws IndexOutOfBoundsException if index is out of range
	 *                                   <tt>(index &lt; 0 || index &gt; size())</tt>.
	 */
	@Override
	public void add( int index, Object element )
	{
		if( (index > size) || (index < 0) )
		{
			throw new IndexOutOfBoundsException( "Index: " + index + ", Size: " + size );
		}

		ensureCapacity( size + 1 );  // Increments modCount!!
		System.arraycopy( elementData, index, elementData, index + 1, size - index );
		elementData[index] = element;
		size++;
	}

	/**
	 * Removes the element at the specified position in this list.
	 * Shifts any subsequent elements to the left (subtracts one from their
	 * indices).
	 *
	 * @param index the index of the element to removed.
	 * @return the element that was removed from the list.
	 * @throws IndexOutOfBoundsException if index out of range <tt>(index
	 *                                   &lt; 0 || index &gt;= size())</tt>.
	 */
	@Override
	public Object remove( int index )
	{
		RangeCheck( index );

		modCount++;
		Object oldValue = elementData[index];

		int numMoved = size - index - 1;
		if( numMoved > 0 )
		{
			System.arraycopy( elementData, index + 1, elementData, index, numMoved );
		}
		elementData[--size] = null; // Let gc do its work

		return oldValue;
	}

	/**
	 * Removes all of the elements from this list.  The list will
	 * be empty after this call returns.
	 */
	@Override
	public void clear()
	{
		modCount++;

		// Let gc do its work
		for( int i = 0; i < size; i++ )
		{
			elementData[i] = null;
		}

		size = 0;
	}

	/**
	 * Appends all of the elements in the specified Collection to the end of
	 * this list, in the order that they are returned by the
	 * specified Collection's Iterator.  The behavior of this operation is
	 * undefined if the specified Collection is modified while the operation
	 * is in progress.  (This implies that the behavior of this call is
	 * undefined if the specified Collection is this list, and this
	 * list is nonempty.)
	 *
	 * @param c the elements to be inserted into this list.
	 * @return <tt>true</tt> if this list changed as a result of the call.
	 * @throws NullPointerException if the specified collection is null.
	 */
	@Override
	public boolean addAll( Collection c )
	{
		Object[] a = c.toArray();
		// 20080124 KSC: replace with arraycopy
		//this.elementData = a;
		int numNew = a.length;
		elementData = new Object[numNew];
		System.arraycopy( a, 0, elementData, 0, numNew );
		//   ensureCapacity(size + numNew);  // Increments modCount
		size += c.size();
		modCount += size;
		modCount++;
		return numNew != 0;
	}

	/**
	 * Inserts all of the elements in the specified Collection into this
	 * list, starting at the specified position.  Shifts the element
	 * currently at that position (if any) and any subsequent elements to
	 * the right (increases their indices).  The new elements will appear
	 * in the list in the order that they are returned by the
	 * specified Collection's iterator.
	 *
	 * @param index index at which to insert first element
	 *              from the specified collection.
	 * @param c     elements to be inserted into this list.
	 * @return <tt>true</tt> if this list changed as a result of the call.
	 * @throws IndexOutOfBoundsException if index out of range <tt>(index
	 *                                   &lt; 0 || index &gt; size())</tt>.
	 * @throws NullPointerException      if the specified Collection is null.
	 */
	@Override
	public boolean addAll( int index, Collection c )
	{
		if( (index > size) || (index < 0) )
		{
			throw new IndexOutOfBoundsException( "Index: " + index + ", Size: " + size );
		}

		Object[] a = c.toArray();
		int numNew = c.size();    // 20080121 KSC: a.length;
		ensureCapacity( size + numNew );  // Increments modCount

		int numMoved = size - index;
		if( numMoved > 0 )
		{
			System.arraycopy( elementData, index, elementData, index + numNew, numMoved );
		}

		System.arraycopy( a, 0, elementData, index, numNew );
		size += numNew;
		return numNew != 0;
	}

	/**
	 * Removes from this List all of the elements whose index is between
	 * fromIndex, inclusive and toIndex, exclusive.  Shifts any succeeding
	 * elements to the left (reduces their index).
	 * This call shortens the list by <tt>(toIndex - fromIndex)</tt> elements.
	 * (If <tt>toIndex==fromIndex</tt>, this operation has no effect.)
	 *
	 * @param fromIndex index of first element to be removed.
	 * @param toIndex   index after last element to be removed.
	 */
	@Override
	protected void removeRange( int fromIndex, int toIndex )
	{
		modCount++;
		int numMoved = size - toIndex;
		System.arraycopy( elementData, toIndex, elementData, fromIndex, numMoved );

		// Let gc do its work
		int newSize = size - (toIndex - fromIndex);
		while( size != newSize )
		{
			elementData[--size] = null;
		}
	}

	/**
	 * Check if the given index is in range.  If not, throw an appropriate
	 * runtime exception.  This method does *not* check if the index is
	 * negative: It is always used immediately prior to an array access,
	 * which throws an ArrayIndexOutOfBoundsException if index is negative.
	 */
	private void RangeCheck( int index )
	{
		if( index >= size )
		{
			throw new IndexOutOfBoundsException( "Index: " + index + ", Size: " + size );
		}
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
			return cursor != size;
		}

		@Override
		public Object next()
		{
			//    checkForComodification();
			try
			{
				Object next = elementData[cursor];
				lastRet = cursor++;
				return next;
			}
			catch( IndexOutOfBoundsException e )
			{
				// checkForComodification();
				throw new NoSuchElementException();
			}
		}

		@Override
		public void remove()
		{
			if( lastRet == -1 )
			{
				throw new IllegalStateException();
			}
			//checkForComodification();

			try
			{
				SpecialArrayList.this.remove( lastRet );
				if( lastRet < cursor )
				{
					cursor--;
				}
				lastRet = -1;
				expectedModCount = modCount;
			}
			catch( IndexOutOfBoundsException e )
			{
				throw new ConcurrentModificationException();
			}
		}

		final void checkForComodification()
		{
			// NOT!
		}
	}
}