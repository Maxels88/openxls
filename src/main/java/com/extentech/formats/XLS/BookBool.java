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
package com.extentech.formats.XLS;

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

/**
 * <b>BOOKBOOL: Additional Workspace Information (DAh)</b><br>
 * <p/>
 * This record stores information about workspace settings
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       grbit          2      Option flags
 * <p/>
 * See book for details of grbit flags, page 425
 * <p/>
 * </p></pre>
 */

public final class BookBool extends com.extentech.formats.XLS.XLSRecord
{

	private static final long serialVersionUID = -4544323710670598072L;
	short grbit;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "BOOKBOOL: " + ((grbit == 0) ? "Save External Links" : "Don't Save External Links") );
		}
	}

}