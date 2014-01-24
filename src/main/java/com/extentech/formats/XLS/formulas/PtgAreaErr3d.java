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
package com.extentech.formats.XLS.formulas;

import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Externsheet;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * An Erroneous BiffRec range spanning 3rd dimension of WorkSheets.
 * <p/>
 * identical to PtgArea3d
 *
 * @see Ptg
 * @see GenericPtgFunc
 */
public class PtgAreaErr3d extends PtgArea3d implements Ptg
{
	private static final Logger log = LoggerFactory.getLogger( PtgAreaErr3d.class );
	// Excel can handle PtgRefErrors within formulas, as long as they are not the result so...

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -9091097082897614748L;

	@Override
	public boolean getIsRefErr()
	{
		return true;
	}

	@Override
	public String getString()
	{
		if( sheetname == null )
		{
			return "#REF!";
		}
		return sheetname + "!#REF!";
	}

	@Override
	public int getLength()
	{
		return PTG_AREAERR3D_LENGTH;
	}

	/* constructor, takes the array of the ptgRef, including
	the identifier so we do not need to figure it out again later...
	*/
	@Override
	public void init( byte[] b )
	{
		record = b;
		ixti = ByteTools.readShort( record[1], record[2] );
		if( ixti > 0 )
		{
			sheetname = GenericPtg.qualifySheetname( getSheetName() );
		}

	}

	@Override
	public Object getValue()
	{
		if( sheetname == null )
		{
			return "#REF!";
		}
		return sheetname + "!#REF!";
	}

	/**
	 * sets referenced sheet
	 * called from copy worksheet
	 * different from PtgArea3d as PtgAreaErr3d's have not set their  firstPtg and lastPtg
	 */
	@Override
	public void setReferencedSheet( Boundsheet b )
	{
		int boundnum = b.getSheetNum();
		Externsheet xsht = b.getWorkBook().getExternSheet( true );
		//TODO: add handling for multi-sheet reference.  Already handled in externsheet
		try
		{
			int xloc = xsht.insertLocation( boundnum, boundnum );
			ixti = (short) xloc;
			if( ixti > 0 )
			{
				sheetname = GenericPtg.qualifySheetname( getSheetName() );
			}
		}
		catch( WorkSheetNotFoundException e )
		{
			log.error( "Unable to set referenced sheet in PtgRef3d " + e );
		}
	}

	@Override
	public void setLocation( String[] s )
	{
		sheetname = GenericPtg.qualifySheetname( s[0] );
	}

	@Override
	public int[] getRowCol()
	{
		return new int[]{ -1, -1 };
	}
}