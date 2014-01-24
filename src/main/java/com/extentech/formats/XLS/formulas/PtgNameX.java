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

import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.toolkit.ByteTools;

/*
	This PTG stores an index to a name.  The ilbl field is a 1 based index to the table 
	of NAME records in the workbook
 
	OFFSET      NAME        sIZE        CONTENTS
	---------------------------------------------
	0           ixti        2           index to externsheet
	2           ilbl        2           Index to the NAME table
	4           (reserved)  2   `       Must be 0;
    
 * @see Ptg
 * @see Formula
    
*/
public class PtgNameX extends PtgName implements Ptg, IxtiListener
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1240996941619495505L;
	short ixti;
	int ilbl;

	@Override
	public boolean getIsOperand()
	{

		return true;
	}

	//lookup Name object  in Workbook and return handle
	@Override
	public Name getName()
	{
		WorkBook b = getParentRec().getSheet().getWorkBook();
		// the externsheet reference is negative, there seems to be a problem
		// off the docs.  Just use a placeholder boundsheet, as the PtgRef3D internally will
		// get the value correctly
		//Externsheet x = b.getExternSheet();
		Name n = null;

		try
		{
			n = b.getName( ilbl );
			n.setSheet( getParentRec().getSheet() );
		}
		catch( Exception e )
		{
			; // it's an AddInFormula... -jm
		}
		//Boundsheet[] bound = x.getBoundSheets(ixti);
		return n;
	}

	@Override
	public void addListener()
	{
		try
		{
			getParentRec().getWorkBook().getExternSheet().addPtgListener( this );
		}
		catch( Exception e )
		{
			// no need to output here.  NullPointer occurs when a ref has an invalid ixti, such as when a sheet was removed  Worksheet exception could never really happen.
		}
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		populateVals();
	}

	private void populateVals()
	{
		ixti = ByteTools.readShort( record[1], record[2] );

		ilbl = ByteTools.readShort( record[3], record[4] );
	}

	@Override
	public int getVal()
	{
		return ilbl;
	}

	/*
	 * 
	 * returns the string value of the name
		@see com.extentech.formats.XLS.formulas.Ptg#getValue()
	 */
	@Override
	public Object getValue()
	{

		WorkBook b = getParentRec().getSheet().getWorkBook();
		String externalname = null;
		try
		{
			externalname = b.getExternalName( ilbl );
		}
		catch( Exception e )
		{
		}
		if( externalname != null )
		{
			return externalname;
		}

		Name n = getName();
		return n.getCalculatedValue();
	}

	public String toString()
	{
		if( parent_rec.getSheet() != null )
		{
			return (String) getValue();
		}
		return "Uninitialized PtgNameX";
	}

	@Override
	public String getTextString()
	{
		Object o = getValue();
		if( o == null )
		{
			return "";
		}
		return o.toString();
	}

	@Override
	public int getLength()
	{
		return PTG_NAMEX_LENGTH;
	}

	/**
	 * @return Returns the ixti.
	 */
	@Override
	public short getIxti()
	{
		return ixti;
	}

	/**
	 * @param ixti The ixti to set.
	 */
	@Override
	public void setIxti( short ixti )
	{
		this.ixti = ixti;
	}

	// KSC: Added to handle External names (denoted by PtgNameX records in ExpressionParser)

	/**
	 * For creating a ptg namex from formula parser
	 */
	@Override
	public void setName( String name )
	{
		ptgId = 0x39;    // PtgNameX
		record = new byte[PTG_NAMEX_LENGTH];
		record[0] = ptgId;
		WorkBook b = getParentRec().getSheet().getWorkBook();
		ilbl = b.getExtenalNameNumber( name );
		ixti = (short) b.getExternSheet().getVirtualReference();
		byte[] bb = ByteTools.shortToLEBytes( ixti );
		record[1] = bb[0];
		record[2] = bb[1];
		byte[] bbb = ByteTools.cLongToLEBytes( ilbl );
		record[3] = bbb[0];
		record[4] = bbb[1];
	}
}
    
    