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
/**
 * CompatibleBigDecimal.java
 *
 *
 * CompatibleBigDecimal deals with the java 1.4-1.5 transition error of BigDecimal.toString().
 *
 * Prior to 1.5, BigDecimal.toString would not return Scientific Notation.  JDK1.5 now allows Scientific Notation 
 * to be returned in some cases for this method.  A new method has been created, .toPlainString that mimics the
 * behavior of the old .toString.   As we do not know the runtime JDK, and returning correctly formatted numbers 
 * is critical to the functionality of ExtenXLS, this class mimics 1.4 functionality for toString under either JDK.
 *
 * A static method is used to increase performance for multiple calls.
 *
 *
 */

import java.lang.reflect.Method;
import java.math.BigDecimal;

public class CompatibleBigDecimal extends BigDecimal
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6816994951413033200L;
	private static Method _methodToString = null;

	// Create the static method we can access later.  This could be .toPlainString or .toString.
	static
	{
		try
		{
			_methodToString = BigDecimal.class.getMethod( "toPlainString", (Class[]) null );
		}
		catch( NoSuchMethodException e )
		{
			try
			{
				_methodToString = BigDecimal.class.getMethod( "toString", (Class[]) null );
			}
			catch( NoSuchMethodException ex )
			{
				Logger.logWarn( "Error creating toString method in CompatibleBigDecimal" );
			}
		}
	}

	/**
	 * Constructor
	 */
	public CompatibleBigDecimal( BigDecimal bd )
	{
		super( bd.unscaledValue(), bd.scale() );
	}

	public CompatibleBigDecimal( String num )
	{
		super( num );
	}

	/**
	 * Compatible toString functionality
	 *
	 * @see java.math.BigDecimal#toString()
	 */
	public String toCompatibleString()
	{
		if( _methodToString != null )
		{
			try
			{
				return (String) _methodToString.invoke( this, (Object[]) null );
			}
			catch( Exception e )
			{
				Logger.logWarn( "Error in calling CompatibleBigDecimal.toString" );
			}
		}
		return null;
	}

}
