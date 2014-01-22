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

/*
    The style record defines a format or style for a particular cell:
    Two kinds of Style records exist, built in and User defined.
    
    Record Data - Built in Styles
    Offset      Name        Size    Contents
    ===================================================
    4           ixfe        2       Index to the style XF record
    6           istyBuiltIn 1       Built in Style numbers
                                    =0x0 Normal
                                    =0x1 RowLevel_n
                                    =0x2 ColLevel_n
                                    =0x3 Comma
                                    =0x4 Currency
                                    =0x5 Percent
                                    =0x6 Comma[0]
                                    =0x7 Currency
    7           iLevel      1       Level of the outline style RowLevel N or
                                    ColLevel_n
                                        
    Record Data - UserDefined styles
    Offset      Name        Size    Contents
    ===================================================
    4           ixfe        2       Index to the style XF record
                                    (Only the low-order 12 bits are used)
    6           cch         1       Length of the Style Name
    7           rgch        (var)   Style Name
                                    

*/
public class Style extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8634444452254051296L;
	short ixfe;
	byte istyBuiltIn;
	byte iLevel;
	short cch;
	String rgch;

	boolean builtIn = false;

	@Override
	public void init()
	{
		super.init();
		short unstripped = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		ixfe = (short) (unstripped & 0x7FF);
		if( (this.getByteAt( 1 ) & 0x128) == 0x128 )
		{
			builtIn = true;
			istyBuiltIn = this.getByteAt( 2 );
			iLevel = this.getByteAt( 3 );
		}
		else
		{
			cch = this.getByteAt( 2 );
			if( cch < 0 )
			{
				cch *= -1;
			}
			byte[] tmp = this.getBytesAt( 3, cch );
			rgch = new String( tmp );
		}

	}

	public String getTypeName()
	{
		return "Style";
	}

}