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
package com.extentech.formats.OOXML;

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.OOXMLReader;
import com.extentech.formats.XLS.OOXMLWriter;
import com.extentech.toolkit.StringTool;
import org.xmlpull.v1.XmlPullParserException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *
 *
 *
 */
public class OOXMLHandle
{

	/**
	 * generates OOXML for a workbook
	 * Creates the ZIP file and writes all files into proper directory structure
	 * Will create either an .xlsx or an .xlsm output, depending upon
	 * whether WorkBookHandle bk contains macros
	 *
	 * @param bk   workbookhandle
	 * @param path output filename and path
	 */
	public static void getOOXML( WorkBookHandle bk, String path ) throws IOException
	{
		try
		{
			if( !com.extentech.formats.XLS.OOXMLAdapter.hasMacros( bk ) )
			{
				path = StringTool.replaceExtension( path, ".xlsx" );
			}
			else    // it's a macro-enabled workbook
			{
				path = StringTool.replaceExtension( path, ".xlsm" );
			}

			java.io.File fout = new java.io.File( path );
			File dirs = fout.getParentFile();
			if( (dirs != null) && !dirs.exists() )
			{
				dirs.mkdirs();
			}
			OOXMLWriter oe = new OOXMLWriter();
			oe.getOOXML( bk, new FileOutputStream( path ) );
		}
		catch( Exception e )
		{
			throw new IOException( "Error parsing OOXML file: " + e.toString() );
		}
	}

	/**
	 * OOXML parseNBind - reads in an OOXML (Excel 7) workbook
	 *
	 * @param bk    WorkBookHandle - workbook to input
	 * @param fName OOXML filename (must be a ZIP file in OPC format)
	 * @throws XmlPullParserException
	 * @throws IOException
	 * @throws CellNotFoundException
	 */
	public static void parseNBind( WorkBookHandle bk, String fName ) throws XmlPullParserException, IOException, CellNotFoundException
	{
		OOXMLReader oe = new OOXMLReader();
		oe.parseNBind( bk, fName );
	}
}
