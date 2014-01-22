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

import java.util.List;

public interface RecycleBin
{

	/**
	 * add a Recyclable to the
	 * bin
	 */
	public void addItem( Recyclable r ) throws RecycleBinFullException;

	/**
	 * add a Recyclable to the
	 * bin
	 */
	public void addItem( Object key, Recyclable r ) throws RecycleBinFullException;

	/**
	 * get an unused Recyclable item from the bin
	 *
	 * @throws RecycleBinFullException
	 */
	public Recyclable getItem() throws RecycleBinFullException;

	/**
	 * get all of the items in this bin
	 */
	public List getAll();

	/**
	 * empty the current contents of the bin
	 */
	public void empty();

	/**
	 * set the maximum number of items for this
	 * bin
	 */
	public void setMaxItems( int i );

}