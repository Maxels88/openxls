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
package org.openxls.formats.cellformat;

import org.openxls.toolkit.StringTool;

public class CellFormatFactory
{
	private CellFormatFactory()
	{
		// this is a static-only class
		throw new UnsupportedOperationException();
	}

	public static CellFormat fromPatternString( String pattern )
	{
		if( (null == pattern) || "".equals( pattern ) || "General".equalsIgnoreCase( pattern ) )
		{
			return new GeneralCellFormat();
		}

		String[] pats = pattern.split( ";" );

		String tester = StringTool.convertPatternExtractBracketedExpression( pats[0] );
		if( tester.matches( ".*(((y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}).*)+).*" ) )
		{
			String string;
			if( pats.length > 3 )
			{
				string = StringTool.convertPatternFromExcelToStringFormatter( pats[3], false );
			}
			else
			{
				string = "%s";
			}

			return new DateCellFormat( StringTool.convertDatePatternFromExcelToStringFormatter( tester ), string );
		}

		String positive;
		String negative;
		String zero;
		String string;
		positive = StringTool.convertPatternFromExcelToStringFormatter( pats[0], false );

		negative = StringTool.convertPatternFromExcelToStringFormatter( pats[((pats.length > 1) ? 1 : 0)], true );

		if( pats.length > 2 )
		{
			zero = StringTool.convertPatternFromExcelToStringFormatter( pats[2], false );
		}
		else
		{
			zero = positive;
		}

		if( pats.length > 3 )
		{
			string = StringTool.convertPatternFromExcelToStringFormatter( pats[3], false );
		}
		else
		{
			string = "%s";
		}

		return new NumberCellFormat( positive, negative, zero, string );
	}
}
