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

import java.io.IOException;
import java.io.OutputStream;
import java.util.concurrent.locks.ReentrantReadWriteLock;

/**
 * An <code>IndirectOutputStream</code> forwards all requests unmodified to
 * another output stream which may be changed at runtime. By default an
 * <code>IOException</code> will be thrown if no sink is configured when a
 * request is received. The stream may be configured to drop such requests
 * instead.
 */
public class IndirectOutputStream extends OutputStream
{
	private OutputStream sink;
	private boolean discardOnNull;
	private final ReentrantReadWriteLock lock = new ReentrantReadWriteLock();

	/**
	 * Creates a new <code>IndirectOutputStream</code> with no sink which
	 * fails when no sink is present.
	 */
	public IndirectOutputStream()
	{
		this( null );
	}

	/**
	 * Creates a new <code>IndirectOutputStream</code> with the given sink
	 * which fails when no sink is present.
	 *
	 * @param sink the initial sink
	 */
	public IndirectOutputStream( OutputStream sink )
	{
		this( sink, false );
	}

	/**
	 * Creates a new <code>IndirectOutputStream</code> with the given sink
	 * and behavior when no sink is present.
	 *
	 * @param sink    the initial sink
	 * @param discard whether to discard requests when no sink is present
	 */
	public IndirectOutputStream( OutputStream sink, boolean discard )
	{
		this.sink = sink;
		discardOnNull = discard;
	}

	/**
	 * Gets the currently configured sink.
	 *
	 * @return the stream to which requests are currently being forwarded
	 * or <code>null</code> if no sink present
	 */
	public OutputStream getSink()
	{
		// don't bother with a read lock, a single read is atomic
		return sink;
	}

	/**
	 * Sets the stream to which requests are forwarded.
	 *
	 * @param sink the stream to which requests should be forwarded
	 *             or <code>null</code> to remove the current sink
	 */
	public void setSink( OutputStream sink )
	{
		lock.writeLock().lock();
		try
		{
			this.sink = sink;
		}
		finally
		{
			lock.writeLock().unlock();
		}
	}

	/**
	 * Gets the current behavior when no wink is present.
	 *
	 * @return whether requests will be discarded when no sink is present
	 */
	public boolean discardOnNoSink()
	{
		// don't bother with a read lock, a single read is atomic
		return discardOnNull;
	}

	/**
	 * Sets the behavior when no sink is present.
	 *
	 * @param discard whether to discard requests when no sink is present
	 */
	public void discardOnNoSink( boolean discard )
	{
		// don't bother with a write lock, a single write is atomic
		// as are all uses of this field 
		discardOnNull = discard;
	}

	private boolean checkSink() throws IOException
	{
		if( null == sink )
		{
			if( discardOnNull )
			{
				return true;
			}
			throw new IOException( "sink not connected" );
		}
		return false;
	}

	@Override
	public synchronized void write( int b ) throws IOException
	{
		lock.readLock().lock();
		try
		{
			if( checkSink() )
			{
				return;
			}
			sink.write( b );
		}
		finally
		{
			lock.readLock().unlock();
		}
	}

	@Override
	public synchronized void write( byte[] b ) throws IOException
	{
		lock.readLock().lock();
		try
		{
			if( checkSink() )
			{
				return;
			}
			sink.write( b );
		}
		finally
		{
			lock.readLock().unlock();
		}
	}

	@Override
	public synchronized void write( byte[] b, int off, int len ) throws IOException
	{
		lock.readLock().lock();
		try
		{
			if( checkSink() )
			{
				return;
			}
			sink.write( b, off, len );
		}
		finally
		{
			lock.readLock().unlock();
		}
	}

	@Override
	public synchronized void flush() throws IOException
	{
		lock.readLock().lock();
		try
		{
			if( checkSink() )
			{
				return;
			}
			sink.flush();
		}
		finally
		{
			lock.readLock().unlock();
		}
	}

	@Override
	public synchronized void close() throws IOException
	{
		lock.readLock().lock();
		try
		{
			if( checkSink() )
			{
				return;
			}
			sink.close();
		}
		finally
		{
			lock.readLock().unlock();
		}
	}
}
