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

import java.io.IOException;
import java.io.OutputStream;
import java.io.Writer;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.charset.Charset;
import java.nio.charset.CharsetDecoder;
import java.nio.charset.CoderResult;
import java.nio.charset.CodingErrorAction;

/**
 * A <code>WriterOutputStream</code> is a bridge from byte streams to character
 * streams: bytes written to it are decoded into characters using a specified
 * {@linkplain Charset charset}. The charset that it uses may be specified by
 * name or given explicitly, or the system default charset may be used. The
 * decoded characters are written to a provided {@link Appendable}, which will
 * usually be a {@link Writer}.
 * <p/>
 * <p>The input is buffered so that writes don't need to be aligned to
 * character boundaries. Because conversion is only performed when the input
 * buffer is full (or when the stream is flushed) output also behaves as if it
 * were buffered. It is therefore usually unnecessary to externally buffer the
 * input or output.</p>
 * <p/>
 * <p>In some charsets some or all characters are represented by multi-byte
 * sequences. If a byte sequence is encountered that is not valid in the input
 * charset or that cannot be mapped to a valid Unicode character it will be
 * replaced in the output with the value <code>"\uFFFD"</code>. If more control
 * over the decoding process is required use {@link CharsetDecoder}.</p>
 *
 * @see InputStreamReader
 * @see OutputStreamWriter
 */
public class WriterOutputStream extends OutputStream
{
	private static final int BUFFER_SIZE = 8192;

	private Appendable target;
	private CharsetDecoder decoder;
	private float bytesPerChar;

	private ByteBuffer inputBuffer;
	private CharBuffer outputBuffer;

	/**
	 * Creates a <code>WriterOutputStream</code> with the default charset.
	 *
	 * @param target the sink for the decoded characters
	 */
	public WriterOutputStream( Appendable target )
	{
		this( target, Charset.defaultCharset() );
	}

	/**
	 * Creates a <code>WriterOutputStream</code> with the named charset.
	 *
	 * @param target  the sink for the decoded characters
	 * @param charset the character set with which to interpret the input bytes
	 */
	public WriterOutputStream( Appendable target, String charset )
	{
		this( target, Charset.forName( charset ) );
	}

	/**
	 * Creates a <code>WriterOutputStream</code> with the given charset.
	 *
	 * @param target  the sink for the decoded characters
	 * @param charset the character set with which to interpret the input bytes
	 */
	public WriterOutputStream( Appendable target, Charset charset )
	{
		this.target = target;

		bytesPerChar = charset.newEncoder().maxBytesPerChar();
		decoder = charset.newDecoder();
		decoder.onMalformedInput( CodingErrorAction.REPLACE );
		decoder.onUnmappableCharacter( CodingErrorAction.REPLACE );

		inputBuffer = ByteBuffer.allocate( (int) Math.ceil( BUFFER_SIZE * bytesPerChar ) );
		outputBuffer = CharBuffer.allocate( BUFFER_SIZE );
	}

	@Override
	public synchronized void write( byte[] buffer, int offset, int length ) throws IOException
	{
		if( null == decoder )
		{
			throw new IOException( "this stream has been closed" );
		}

		// if the input buffer is too full decode it first
		if( inputBuffer.remaining() < bytesPerChar )
		{
			decodeInputBuffer();
		}

		// Append the input to the buffer if it'll fit. If not and there are
		// bytes left in the buffer fill it anyway so we don't lose them.
		if( (length <= inputBuffer.remaining()) || (inputBuffer.position() > 0) )
		{
			int fill = Math.min( inputBuffer.remaining(), length );
			inputBuffer.put( buffer, offset, fill );

			// if we've buffered the entire input, we're done
			if( fill == length )
			{
				return;
			}

			// otherwise, decode the input buffer
			inputBuffer.flip();
			decode( inputBuffer );

			fill -= inputBuffer.remaining();
			offset += fill;
			length -= fill;
			inputBuffer.clear();
		}

		// if the remaining input won't fit in the buffer decode it directly
		if( length > inputBuffer.remaining() )
		{
			ByteBuffer tempBuffer = ByteBuffer.wrap( buffer, offset, length );
			decode( tempBuffer );

			// if any bytes are left over, put them in the input buffer
			if( tempBuffer.hasRemaining() )
			{
				inputBuffer.put( tempBuffer );
			}
		}

		// otherwise, just append it to the buffer
		else
		{
			inputBuffer.put( buffer, offset, length );
		}
	}

	@Override
	public synchronized void write( int b ) throws IOException
	{
		if( null == decoder )
		{
			throw new IOException( "this stream has been closed" );
		}

		// if the buffer is full, decode it first
		if( !inputBuffer.hasRemaining() )
		{
			decodeInputBuffer();
		}

		// append the input to the buffer
		inputBuffer.put( (byte) b );
	}

	/**
	 * Flushes the input buffer through the decoder.
	 * If the input buffer ends with an incomplete character it will remain in
	 * the buffer. If the underlying character sink is a {@link Writer} its
	 * <code>flush</code> method will be called after all buffered input has
	 * been flushed.
	 */
	@Override
	public void flush() throws IOException
	{
		synchronized(this)
		{
			if( null == decoder )
			{
				throw new IOException( "this stream has been closed" );
			}

			decodeInputBuffer();
		}

		// flush the underlying Writer, if any
		if( target instanceof Writer )
		{
			((Writer) target).flush();
		}
	}

	private void decodeInputBuffer() throws IOException
	{
		inputBuffer.flip();
		decode( inputBuffer );
		inputBuffer.compact();
	}

	private void decode( ByteBuffer bytes ) throws IOException
	{
		CoderResult result;

		do
		{
			outputBuffer.clear();
			result = decoder.decode( bytes, outputBuffer, false );

			outputBuffer.flip();
			target.append( outputBuffer );
		} while( result.equals( CoderResult.OVERFLOW ) );
	}

	/**
	 * Closes the stream, flushing it first.
	 * If any partial characters remain in the input buffer the replacement
	 * value will be output in their place. If the underlying character sink is
	 * a {@link Writer} its <code>close</code> method will be called after all
	 * buffered input has been flushed. Once the stream has been closed further
	 * calls to <code>write</code> or <code>flush</code> will cause an
	 * <code>IOException</code> to be thrown. Closing a previously closed
	 * stream has no effect.
	 */
	@Override
	public void close() throws IOException
	{
		synchronized(this)
		{
			CoderResult result;
			if( null == decoder )
			{
				return;
			}

			// flush the input buffer
			inputBuffer.flip();
			do
			{
				outputBuffer.clear();
				result = decoder.decode( inputBuffer, outputBuffer, true );

				outputBuffer.flip();
				target.append( outputBuffer );
			} while( result.equals( CoderResult.OVERFLOW ) );

			// flush the decoder
			do
			{
				outputBuffer.clear();
				result = decoder.flush( outputBuffer );

				outputBuffer.flip();
				target.append( outputBuffer );
			} while( result.equals( CoderResult.OVERFLOW ) );

			// release the buffers and decoder
			inputBuffer = null;
			outputBuffer = null;
			decoder = null;
		}

		// close the underlying Writer, if any
		if( target instanceof Writer )
		{
			((Writer) target).close();
		}
	}
}
