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

import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import java.util.Enumeration;

public class NamingEnumerationImpl implements NamingEnumeration
{

	private Enumeration e = null;

	void setEnumeration( Enumeration ex )
	{
		e = ex;
	}

	/* (non-Javadoc)
	 * @see javax.naming.NamingEnumeration#close()
	 */
	@Override
	public void close() throws NamingException
	{
		e = null;
	}

	/* (non-Javadoc)
	 * @see javax.naming.NamingEnumeration#hasMore()
	 */
	@Override
	public boolean hasMore() throws NamingException
	{
		return e.hasMoreElements();
	}

	/* (non-Javadoc)
	 * @see javax.naming.NamingEnumeration#next()
	 */
	@Override
	public Object next() throws NamingException
	{
		return e.nextElement();
	}

	/* (non-Javadoc)
	 * @see java.util.Enumeration#hasMoreElements()
	 */
	@Override
	public boolean hasMoreElements()
	{
		return e.hasMoreElements();
	}

	/* (non-Javadoc)
	 * @see java.util.Enumeration#nextElement()
	 */
	@Override
	public Object nextElement()
	{
		return e.nextElement();
	}

}