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
/*
 * Created on Dec 4, 2004
 *
 * To change the template for this generated file go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
package org.openxls.formats.XLS;

import java.io.OutputStream;

/**
 * To change the template for this generated type comment go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
public interface Book
{

	public BiffRec addRecord( BiffRec b, boolean y );

	public ContinueHandler getContinueHandler();

	public ByteStreamer getStreamer();

	public void setFirstBof( Bof b );

	/**
	 * Stream the Book bytes to out
	 *
	 * @param out
	 * @return
	 */
	public int stream( OutputStream out );

	public abstract String toString();

	public abstract String getFileName();

	/**
	 * get a handle to the factory
	 */
	public abstract WorkBookFactory getFactory();

	public void setFactory( WorkBookFactory fact );

	/**
	 * Dec 15, 2010
	 *
	 * @param rec
	 */
	public Boundsheet getSheetFromRec( BiffRec rec, Long l );

	/** set readiness -- is it done parsing? */
	// public boolean isReady();
	// public void setReady(boolean t);

	/** the default session provides initial values
	 * @return
	 */
	//public BookSession getDefaultSession();
	//public void setDefaultSession(BookSession session);

}