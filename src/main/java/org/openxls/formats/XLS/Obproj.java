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
package org.openxls.formats.XLS;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <b>Obproj: Visual Basic Project (D3h)</b><br>
 */
public final class Obproj extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Obproj.class );
	private static final long serialVersionUID = 7952338892026147433L;

	@Override
	public void init()
	{
		super.init();
			log.debug( "Visual Basic Project Detected" );
	}
}