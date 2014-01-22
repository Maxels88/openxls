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
import java.io.Reader;

/**
 * Stream parser for delimiter-separated values formats.
 * These include comma separated values (CSV) and tab separated values (TSV).
 */
public class DelimitedValuesParser
{
	/**
	 * Represents the type of a token.
	 */
	public enum Token
	{
		VALUE, NEWLINE, EOF
	}

	private Reader source;

	/**
	 * The delimiter used to separate values.
	 */
	private char delimiter = '\t';

	/**
	 * Contains the current value, if any.
	 */
	private StringBuilder value = new StringBuilder();

	/**
	 * The last token returned.
	 */
	private Token current = null;

	/**
	 * The next token to be returned.
	 * This is used when a single character ends a token and is itself a token.
	 */
	private Token next = null;

	public DelimitedValuesParser( Reader source )
	{
		this.source = source;
	}

	public Token next() throws IOException
	{
		// reset the value builder
		value.setLength( 0 );

		// if there's a token waiting, return it
		if( next != null )
		{
			current = next;
			next = null;
			return current;
		}

		while( true )
		{
			int read = source.read();
			if( read == -1 )
			{
				if( value.length() == 0 )
				{
					return current = Token.EOF;
				}
				else
				{
					return current = Token.VALUE;
				}
			}

			if( read == delimiter )
			{
				return current = Token.VALUE;
			}

			if( read == '\n' )
			{
				if( value.length() > 0 )
				{
					if( value.charAt( value.length() - 1 ) == '\r' )
					{
						value.setLength( value.length() - 1 );
					}
					next = Token.NEWLINE;
					return current = Token.VALUE;
				}
				else
				{
					return current = Token.NEWLINE;
				}
			}

			value.append( (char) read );
		}
	}

	public String getValue()
	{
		return value.length() > 0 ? value.toString() : null;
	}
}
