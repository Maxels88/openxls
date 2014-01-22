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

import com.extentech.toolkit.ByteTools;

/*
   An Erroneous BiffRec range spanning 3rd dimension of WorkSheets.

	identical to PtgRef3d 
	
 * @see Ptg
 * @see Formula

    
*/
public class PtgRefErr3d extends PtgRef3d implements Ptg
{

	private static final long serialVersionUID = 8691902605148033701L;

	public boolean getIsRefErr()
	{
		return true;
	}

	// IDs: 3C (R) 5C (V) 7C (A)
	public PtgRefErr3d()
	{
		record = new byte[PTG_REFERR3D_LENGTH];
		record[0] = 0x3c;
		// record[1]= index to REF entry in EXTERNSHEET 

	}

	public String getString()
	{
		if( sheetname == null )
		{
			return "#REF!";
		}
		return sheetname + "!#REF!";
	}

	public int getLength()
	{
		return PTG_REFERR3D_LENGTH;
	}

	/*
	 Throw this data into a ptgref's
	 Ixti can reference sheets that don't exist, causing np error.  As we don't perform any functions
	 upon a PTGRef3D error, just swallow
	*/
	public void populateVals()
	{
		ixti = ByteTools.readShort( record[1], record[2] );
		if( ixti > 0 )
		{
			this.sheetname = GenericPtg.qualifySheetname( this.getSheetName() );
		}
	}

	public int[] getRowCol()
	{
		return new int[]{ -1, -1 };
	}

	public Object getValue()
	{
		if( sheetname == null )
		{
			return "#REF!";
		}
		return sheetname + "!#REF!";
	}

	public String getLocation()
	{
		if( sheetname == null )
		{
			return "#REF!";
		}
		return sheetname + "!#REF!";
	}

	public void setLocation( String[] s )
	{
		sheetname = GenericPtg.qualifySheetname( s[0] );
	}

}
