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

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.ChartHandle;
import com.extentech.ExtenXLS.CommentHandle;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.FormulaHandle;
import com.extentech.ExtenXLS.ImageHandle;
import com.extentech.ExtenXLS.RowHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.OOXML.Border;
import com.extentech.formats.OOXML.Dxf;
import com.extentech.formats.OOXML.Fill;
import com.extentech.formats.OOXML.NumFmt;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.OOXML.OneCellAnchor;
import com.extentech.formats.OOXML.SheetView;
import com.extentech.formats.OOXML.TwoCellAnchor;
import com.extentech.formats.XLS.charts.Chart;
import com.extentech.formats.XLS.charts.OOXMLChart;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipOutputStream;

/**
 * Breaking out functionality for writing out of OOXMLAdapter
 */
public class OOXMLWriter extends OOXMLAdapter implements OOXMLConstants
{
	private static final Logger log = LoggerFactory.getLogger( OOXMLWriter.class );
	/**
	 * generates OOXML for a workbook (see specification above)
	 * Creates the ZIP file and writes all files into proper directory structure re: OPC
	 *
	 * @param workbook
	 * @param out      outputStream used by ZipOutputStream
	 * @throws IOException
	 */
	public void getOOXML( WorkBookHandle bk, OutputStream out ) throws IOException
	{
		// clear out  ArrayLists ContentLists
		mainContentList = new ArrayList();       // main .rels
		wbContentList = new ArrayList();         // workbook.xml.rels
		drContentList = new ArrayList();         // drawingX.xml.rels
		shContentList = new ArrayList();         // sheetX.xml.rels
		sheetsContentList = new ArrayList();     // total contents of each sheet     
		vmlId = 0;                               // reset ordinal id's for external refs
		drawingId = 0;
		commentsId = 0;
		activeXId = 0;
		activeXBinaryId = 0;
		printerSettingsId = 0;
		oleObjectsId = 0;
		chartId = 0;
		imgId = 0;

		// create XLSX zip file from OutputStream 
		zip = new ZipOutputStream( out );

		// Wrap the ZipOutputStream in a Writer to handle character encoding
		// setting encoding is important when input encoding is not utf8; writing to utf8 will convert (for example, format strings in styles.xml ...)
		writer = new OutputStreamWriter( zip, inputEncoding );   //"UTF-8" );

		// retrive external directory used to store passthrough files
		externalDir = getTempDir( bk.getWorkBook().getFactory().getFileName() );
		// writeOOXML files to zip
		writeOOXML( bk );
		// write main .rels file            
		writeRels( mainContentList, "_rels/.rels" );      // TODO: if have doc properties, must add to .rels
		// write [Content_Types].xml
		writeContentType();

		// write out the defferred files to the zip
		writeDeferredFiles();
		writer.flush();
		writer.close();
		if( zip != null )
		{
			zip.flush();
			zip.close();
		}
		if( !bk.getWorkBook().getFactory().getFileName().endsWith( ".tmp" ) )
		{
			deleteDir( new File( externalDir ) );
		}
		zip = null;
	}

	/**
	 * write the deferred files to the zipfile
	 *
	 * @throws IOException
	 */
	protected void writeDeferredFiles() throws IOException
	{
		Iterator its = deferredFiles.keySet().iterator();
		while( its.hasNext() )
		{
			String k = its.next().toString();
			String fx = (String) deferredFiles.get( k );
			writeFileToZIP( fx, k );
			File fdel = new File( fx );
			fdel.deleteOnExit();
			fdel.delete();
		}
	}

	/**
	 * handle Sheet-level External References that are pass-throughs and NOT
	 * recreated on output: control (activeX), printerSettings, oleObjects
	 * Fairly complicated and klugdy ((:
	 *
	 * @param type
	 * @param externalOOXML List is in format of: [0]= type, [1]= pass-through filename,
	 *                      [2] original filename, [3] original rid [, [4]= Extra info, if
	 *                      any, [5]= Embedded files, if any]]
	 * @return
	 */
	private void writeSheetLevelExternalReferenceOOXML( Writer out, String type, List externalOOXML )
	{

		ArrayList refs = getExternalRefType( type, externalOOXML );   // do we have any of the specific type of references? 
		 /*
          * because the following methods all write files to the zip it causes
          * problems with the current zipentry.
          * 
          * for this reason, we must return the file for later writing, after the
          * current zipentry is closed (aka: sheet1.xml)
          */

		if( refs.size() > 0 )
		{ // got something
			StringBuffer ooxml = new StringBuffer();
			int rId = -1;
			try
			{
				if( type.equals( "oleObject" ) )
				{
					ooxml.append( writeExOOMXLElement( "oleObject", refs, false ) );
				}
				else if( type.equals( "activeX" ) )
				{
					ooxml.append( writeExOOMXLElement( "control", refs, false ) );

				}
				else if( type.equalsIgnoreCase( "printerSettings" ) )
				{ // TODO: also: orientation, horizontalDPI
					ooxml.append( writeExOOMXLElement( "pageSetup", refs, true ) );
				}
				else
				{ // TESTING - remove when done!
					// *********************************************************************
					log.warn( "Unknown External Type " + type );
				}
			}
			catch( IOException e )
			{
				log.error( "OOXMLWriter.writeSheetLevelExternalReferenceOOXML: " + e.toString() );
			}
			try
			{
				out.write( ooxml.toString() );
			}
			catch( Exception e )
			{
				;
			}
		}
	}

	/**
	 * search through list of external objects to retrieve those associated with the desired type
	 *
	 * @param type
	 * @param externalOOXML List of previously saved external Objects
	 * @return
	 */
	private ArrayList getExternalRefType( String type, List externalOOXML )
	{
		ArrayList refs = new ArrayList();
		for( Object anExternalOOXML : externalOOXML )
		{
			String[] s = (String[]) anExternalOOXML;
			if( (s != null) && (s.length >= 0) )
			{   // id, dir, filename, rId [, extra info [, embedded file info]]
				if( s[0].equalsIgnoreCase( type ) )
				{ // got one
					refs.add( s );
				}
			}
		}
		return refs;
	}

	/**
	 * most OOXML objects external to workbook are handled here
	 * (NOTE: External objects linked to sheets are handled elsewhere)
	 * Eventually many of these will (docprops, etc) be created by ExtenXLS
	 *
	 * @param bk
	 * @throws IOException
	 */
	private void writeExternalOOXML( WorkBookHandle bk ) throws IOException
	{
		List externalOOXML = bk.getWorkBook().getOOXMLObjects();
		for( Object anExternalOOXML : externalOOXML )
		{
			String[] s = (String[]) anExternalOOXML;
			if( (s != null) && (s.length >= 3) )
			{   // id, dir, filename, rid, [extra info], [embedded information]
				String type = s[EX_TYPE];
				if( type.equalsIgnoreCase( "props" ) ||
						type.equals( "exprops" ) ||
						type.equals( "custprops" ) ||
						type.equals( "connections" ) ||
	                       /* type.equals("calc") || 20081122 KSC: Skip calcChain for now as will error if problems with formulas*/
						type.equals( "externalLink" ) ||
						type.equals( "theme" ) ||
						type.equals( "vba" ) )
				{
					if( type.equals( "props" ) || type.equals( "exprops" ) )
					{
						writeExOOXMLFile( s, mainContentList );
					}
					else
					{
						writeExOOXMLFile( s, wbContentList );
					}
				}
			}
		}
	}

	/**
	 * given an external reference String array, parse, obtaining embedded files if present, then writing correct .rels + master file to ZIP
	 * also adds master file to ContentList for later inclusion in corresponding .rels
	 *
	 * @param String[]  external reference String[] { EX_TYPE, EX_PATH, EX_FNAME, (rid)[ , EX_EXTRAINFO [, EX_EMBEDINFO]]
	 * @param ArrayList contentList
	 * @return rId  int one-based position in contentList for this reference
	 * @throws IOException
	 */
	private int writeExOOXMLFile( String[] s, ArrayList contentList ) throws IOException
	{
		String p = s[EX_PATH];
		String f = s[EX_FNAME];
		int rId = -1;
		File finx = new File( f );
		if( finx.exists() )
		{ // external object hasn't already been input into zip
			String fname = getExOOXMLFileName( f );
			if( s.length > EX_EMBEDINFO )
			{ // then linked to external files must copy and account for
				ArrayList cl = new ArrayList();
				String[] embeds = StringTool.splitString( s[EX_EMBEDINFO].substring( 1, s[EX_EMBEDINFO].length() - 1 ), "," );
				if( embeds != null )
				{  // EMBEDINFO as: type/path/filename these are usually activeXBinary
					for( String embed : embeds )
					{
						String pp = embed.trim();    // original path + filename
						String typ = pp.substring( 0, pp.indexOf( "/" ) );
						pp = pp.substring( pp.indexOf( "/" ) + 1 );
						String pth = pp.substring( 0, pp.lastIndexOf( "/" ) + 1 );
						pp = pp.substring( pp.lastIndexOf( "/" ) + 1 );
						String ff = pp;
						if( !typ.equals( "externalLinkPath" ) )
						{  // retrieve embeds (previously stored in externaldir)
							// ensure correct filename and write out to zip, 
							// storing embed info in content list for for .rels
							ff = getExOOXMLFileName( pp ); // ensure proper ordinal number for filename, if necessary                           
							deferredFiles.put( pth + ff,
							                   externalDir + pp ); // file on disk= externalDir + pp, desired filename in zip= pth+ff                            
							cl.add( new String[]{ "/" + pth + ff, typ } );
							sheetsContentList.addAll( cl );   // most embeds need to be written to main content list 
						} // TODO: externalBooks do not write embeds; other types of external links??? dde, ole ...?
						else
						{  // externalLinkPath - exception to the rule - should only be added to wb content list 
							cl.add( new String[]{ "/" + pth + ff, typ } );
						}
					}
				}
				writeRels( cl, p + "_rels/" + fname + ".rels" );
			}
			deferredFiles.put( p + fname, f );    // file in zip: p+ fname, file on disk= f (in externalDir)

			contentList.add( new String[]{ "/" + p + fname, s[EX_TYPE] } );
			// remove original external filename from disk and map new name on zip to avoid dups
			rId = contentList.size();
		}
		return rId;
	}

	/**
	 * given an ArrayList containing external reference String arrays, parse, obtaining embedded files if present, then writing correct .rels + master file to ZIP
	 * also adds master file to shContentList for later inclusion in sheetX.xml.rels
	 * Also generates proper OOXML for element, linking r:id in SheetX.xml to r:id in Sheet.X.xml.rels
	 *
	 * @param xmlElement String root XMLElement, if !onlyOne, adds a root + "s" e.g. controls or oleObjects
	 * @param refs       ArrayList containing list of external reference String[] { EX_TYPE, EX_PATH, EX_FNAME, (rid)[ , EX_EXTRAINFO [, EX_EMBEDINFO]]
	 * @param onlyOne    if true, doesn't add root "s" element
	 * @return OOXML defining this element
	 * @throws IOException
	 */
	private String writeExOOMXLElement( String xmlElement, ArrayList refs, boolean onlyOne ) throws IOException
	{
		StringBuffer ooxml = new StringBuffer();
		if( !onlyOne )
		{
			ooxml.append( "<" + xmlElement + "s>" );
			ooxml.append( "\r\n" );
		}
		for( Object ref : refs )
		{
			String[] s = (String[]) ref;
			if( (s.length > EX_EXTRAINFO) && (s[EX_EXTRAINFO] != null) )
			{   // add associated info, if any 
				ooxml.append( "<" + xmlElement + " " + s[EX_EXTRAINFO] + " r:id=\"rId" + (shContentList.size() + 1) + "\"/>" );
				ooxml.append( "\r\n" );
			}
			else
			{
				ooxml.append( "<" + xmlElement + " r:id=\"rId" + (shContentList.size() + 1) + "\"/>" );
				ooxml.append( "\r\n" );
			}
			writeExOOXMLFile( s, shContentList );
		}
		if( !onlyOne )
		{
			ooxml.append( "</" + xmlElement + "s>" );
			ooxml.append( "\r\n" );
		}
		return ooxml.toString();
	}

	/**
	 * generic method for creating a .rels file from a content list array cl
	 *
	 * @param cl        ArrayList contentList (type, filename)
	 * @param relsfname
	 * @throws IOException
	 */
	protected void writeRels( ArrayList cl, String relsfname ) throws IOException
	{
		if( (cl == null) || (cl.size() == 0) )
		{
			return;   // don't write a .rels if there are no relationships to track
		}
		StringBuffer rels = new StringBuffer();
		rels.append( xmlHeader );
		rels.append( "\r\n" );
		rels.append( "<Relationships xmlns=\"" + pkgrelns + "\">" );
		rels.append( "\r\n" );

		for( int i = 0; i < cl.size(); i++ )
		{
			rels.append( "<Relationship Id=\"rId" + (i + 1) );
			// TODO: only external types are hyperlink externalLink?????
			String type = ((String[]) cl.get( i ))[1];
			if( !type.equals( "hyperlink" ) && !type.equals( "externalLinkPath" ) )
			{
				rels.append( "\" Type=\"" + getRelationshipType( type ) + "\" Target=\"" + ((String[]) cl.get( i ))[0] + "\"/>" );
			}
			else
			{
				rels.append( "\" Type=\"" + getRelationshipType( type ) + "\" Target=\"" + ((String[]) cl.get( i ))[0] + "\" TargetMode=\"External\"/>" );
			}
			rels.append( "\r\n" );
		}
		rels.append( "</Relationships>" );
		rels.append( "\r\n" );

		// write to tmp
		addDeferredFile( rels, relsfname );
	}

	/**
	 * writes all package contents to [Content_Types].xml
	 * Package contents are contained in global array lists
	 * mainContentList, wbContentList, sheetsContentList, drContentList
	 *
	 * @throws IOException
	 */
	protected void writeContentType() throws IOException
	{
		StringBuffer ct = new StringBuffer();
		ct.append( xmlHeader );
		ct.append( "\r\n" );
		ct.append( "<Types xmlns=\"" + typens + "\">" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"png\" ContentType=\"image/png\"/>" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"emf\" ContentType=\"image/x-emf\"/>" );
		ct.append( "\r\n" );
		ct.append(
				"<Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>" );
		ct.append( "\r\n" );
		ct.append( "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>" );
		ct.append( "\r\n" );
		// write ALL content lists here     
		for( Object aMainContentList : mainContentList )
		{
			ct.append( "<Override PartName=\"" + ((String[]) aMainContentList)[0] + "\" ContentType=\"" + getContentType( ((String[]) aMainContentList)[1] ) + "\"/>" );
			ct.append( "\r\n" );
		}
		for( Object aWbContentList : wbContentList )
		{
			ct.append( "<Override PartName=\"" + ((String[]) aWbContentList)[0] + "\" ContentType=\"" + getContentType( ((String[]) aWbContentList)[1] ) + "\"/>" );
			ct.append( "\r\n" );
		}
         /* printerSettings and vmlDrawing files are not included in Content_Types.xml - rather, they are handled via <Default Extension> element*/
         /* same goes for images */
		for( Object aSheetsContentList : sheetsContentList )
		{
			if( !((((String[]) aSheetsContentList)[1]).equals( "printerSettings" ) ||
					(((String[]) aSheetsContentList)[1]).equals( "vmldrawing" ) ||
					(((String[]) aSheetsContentList)[1]).equals( "hyperlink" ) ||
					(((String[]) aSheetsContentList)[1]).equals( "image" )) )
			{
				ct.append( "<Override PartName=\"" + ((String[]) aSheetsContentList)[0] + "\" ContentType=\"" + getContentType( ((String[]) aSheetsContentList)[1] ) + "\"/>" );
				ct.append( "\r\n" );
			}
		}
		for( Object aDrContentList : drContentList )
		{
			if( !((String[]) aDrContentList)[1].equals( "image" ) )  /* image files not included in Content_Type - rather, they are handled via <Default Extension> element*/
			{
				ct.append( "<Override PartName=\"" + ((String[]) aDrContentList)[0] + "\" ContentType=\"" + getContentType( ((String[]) aDrContentList)[1] ) + "\"/>" );
			}
			ct.append( "\r\n" );
		}

		ct.append( "</Types>" );
		addDeferredFile( ct, "[Content_Types].xml" );
	}

	/**
	 * generates OOXML for a workbook
	 * Creates the ZIP file and writes all files into proper directory structure
	 * Will create either an .xlsx or an .xlsm output, depending upon
	 * whether WorkBookHandle bk contains macros
	 *
	 * @param bk   workbookhandle
	 * @param path output filename and path
	 */
	public void getOOXML( WorkBookHandle bk, String path ) throws Exception
	{
		if( !OOXMLAdapter.hasMacros( bk ) )
		{
			path = StringTool.replaceExtension( path, ".xlsx" );
			format = WorkBookHandle.FORMAT_XLSX;
		}
		else
		{    // it's a macro-enabled workbook
			path = StringTool.replaceExtension( path, ".xlsm" );
			format = WorkBookHandle.FORMAT_XLSM;
		}

		java.io.File fout = new java.io.File( path );
		File dirs = fout.getParentFile();
		if( (dirs != null) && !dirs.exists() )
		{
			dirs.mkdirs();
		}
		getOOXML( bk, new FileOutputStream( path ) );
	}

	/**
	 * creates sharedStrings.xml if there are entries in the SST
	 * and writes it to the root of the OPC ZIP
	 *
	 * @param bk
	 */
	private void writeSSTOOXML( WorkBookHandle bk ) throws IOException
	{
		// SHAREDSTRINGS.XML
		nextZipEntry( "xl/sharedStrings.xml" );
		bk.getWorkBook().getSharedStringTable().writeOOXML( writer );

		wbContentList.add( new String[]{ "/xl/sharedStrings.xml", "sst" } );
	}

	/**
	 * Calls all the necessary methods to create style OOXML, workbook OOXML, sheet(s) OOXML ...
	 * Expects global  zip var to be set to correct zip file
	 *
	 * @param bk
	 * @param output
	 * @throws IOException
	 * @see getOOXML
	 */
	protected void writeOOXML( WorkBookHandle bk ) throws IOException
	{
		int origcalcmode = bk.getWorkBook().getCalcMode();
		bk.getWorkBook().setCalcMode( WorkBook.CALCULATE_EXPLICIT );  // don't recalculate 
		writeExternalOOXML( bk );     // all workbook-level objects we cannot handle at this point such as docProps, themes, or are pass-throughs, such as vbaprojects ...        
		writeSSTOOXML( bk );
		// writeStylesOOXML AFTER sheet OOXML in order to capture any dxf's  (differential xf's used in conditional formatting and others)
		// but ensure that workbook.xml.rels knows about the file styles.xml ***
		wbContentList.add( new String[]{ "/xl/styles.xml", "styles" } );
		writeWorkBookOOXML( bk );
		WorkSheetHandle[] wsh = bk.getWorkSheets();
		bk.getWorkBook().setDxfs( null ); // rebuild
		for( int i = 0; i < wsh.length; i++ )
		{
			writeSheetOOXML( bk, wsh[i], i );
		}
		writeStylesOOXML( bk );   // must do AFTER sheet OOXML to capture any dxf's (differential xf's used in conditional formatting and others)
		bk.getWorkBook().setCalcMode( origcalcmode ); // reset

	}

	/**
	 * Creates Styles.xml with font and xf information, and writes it to the root directory of the OPC ZIP
	 * <p/>
	 * A Style is a named collection of formatting elements.
	 * A cell style can specify number format, cell alignment, font information, cell border specifications, colors, and background / foreground fills.
	 * Table styles specify formatting elements for the regions of a table (e.g. make the header row & totals bold face, and apply light
	 * gray fill to alternating rows in the data portion of the table to achieve striped or banded rows).
	 * PivotTable styles specify formatting elements for the regions of a PivotTable (e.g. 1st & 2nd level subtotals, row axis,
	 * column axis, and page fields).
	 * <p/>
	 * A Style can specify color, fonts, and shape effects directly, or these elements can be referenced indirectly by
	 * referring to a Theme definition. Using styles allows for quicker application of formatting and more consistently
	 * stylized documents.
	 * Themes define a set of colors, font information, and effects on shapes (including Charts). If a style or
	 * formatting element defines its color, font, or effect by referencing a theme, then picking a new theme switches
	 * all the colors, fonts, and effects for that formatting element.
	 * Applying Direct Formatting means that particular elements of formatting (e.g. a bold font face or a number
	 * format) have been applied, but the elements of formatting have been chosen individually instead of
	 * collectively by choosing a named Style. Note that when applying direct formatting, themes can still be
	 * referenced, causing those elements to change as the theme is changed.
	 * <p/>
	 * <p/>
	 * Styles.xml may contain:
	 * <p/>
	 * <p/>
	 * <styleSheet>          ROOT
	 * <p/>
	 * borders (Borders) 3.8.5
	 * cellStyles (Cell Styles) for a named cell style ... built-in styles are referenced here by name e.g. Heading 1, Normal (= Default) ...
	 * [name= "" buildinId="xfid"]
	 * cellStyleXfs (Formatting Records) Master formatting xf's, can override others ...
	 * a cell can have both direct formatting e.g. bold, and a cell style,
	 * therefore, both cellStyleXf + cell xf records must be read
	 * cellXfs (Cell Formats)    Cell xf's; cells in the <sheet> section of Workbook.xml reference the 0-based xf index here
	 * colors (Colors)
	 * dxfs (Formats)            Differential formatting, for all non-cell elements
	 * extLst (Future Feature Data Storage Area)
	 * fills (Fills)
	 * fonts (Fonts)
	 * numFmts (Number Formats)
	 * tableStyles (Table Styles)
	 *
	 * @param bk
	 */
	private void writeStylesOOXML( WorkBookHandle bk ) throws IOException
	{

		StringBuffer stylesooxml = new StringBuffer();
		stylesooxml.append( xmlHeader );
		stylesooxml.append( "<styleSheet xmlns=\"" + xmlns + "\">" );
		stylesooxml.append( "\r\n" );

		// Now create nodes for various XF elements
		AbstractList xfs = bk.getWorkBook().getXfrecs();

		ArrayList cellxfs = new ArrayList();   // references various style source elements for ea xf 
		ArrayList fills = new ArrayList();
		ArrayList borders = new ArrayList();
		ArrayList numfmts = new ArrayList();
		ArrayList fonts = new ArrayList();

		// input default fills -- both appear to be required
		fills.add( Fill.getOOXML( 0, -1, -1 ) ); // none
		fills.add( Fill.getOOXML( 17, -1, -1 ) ); // gray125
		// input default borders element (= no borders)
		borders.add( Border.getOOXML( new int[]{ -1, -1, -1, -1, -1 }, new int[]{ 0, 0, 0, 0, 0 } ) );

		// Iterate the xf's and populate values
		for( int i = 0; i < xfs.size(); i++ )
		{
			Xf xf = (Xf) xfs.get( i );
			addXFToStyle( xf, cellxfs, fills, borders, numfmts, fonts );

		}

		//** stylesheet element contains an ordered SEQUENCE of elements **//

		// Number formats
		if( numfmts.size() > 0 )
		{
			stylesooxml.append( "<numFmts count=\"" + numfmts.size() + "\">" );
			stylesooxml.append( "\r\n" );
			for( Object numfmt : numfmts )
			{
				stylesooxml.append( (String) numfmt );
				stylesooxml.append( "\r\n" );
			}
			stylesooxml.append( "</numFmts>" );
			stylesooxml.append( "\r\n" );
		}

		// fonts element
		stylesooxml.append( "<fonts count=\"" + fonts.size() + "\">" );
		stylesooxml.append( "\r\n" );
		for( Object font : fonts )
		{
			stylesooxml.append( (String) font );
			stylesooxml.append( "\r\n" );
		}
		stylesooxml.append( "</fonts>" );
		stylesooxml.append( "\r\n" );

		// fill patterns element - always has two defaults
		stylesooxml.append( "<fills count=\"" + fills.size() + "\">" );
		stylesooxml.append( "\r\n" );
		for( Object fill : fills )
		{
			stylesooxml.append( (String) fill );
			stylesooxml.append( "\r\n" );
		}
		stylesooxml.append( "</fills>" );
		stylesooxml.append( "\r\n" );

		//borders element - has one default
		stylesooxml.append( "<borders count=\"" + borders.size() + "\">" );
		stylesooxml.append( "\r\n" );
		for( Object border : borders )
		{
			stylesooxml.append( (String) border );
			stylesooxml.append( "\r\n" );
		}
		stylesooxml.append( "</borders>" );
		stylesooxml.append( "\r\n" );

		// cellXfs
		stylesooxml.append( "<cellXfs count=\"" + cellxfs.size() + "\">" );
		stylesooxml.append( "\r\n" );
		for( Object cellxf : cellxfs )
		{
			// xfId= 0 based index of an xf record contained in cellStyleXfs corresponding to the
			// cell style applied to the cell (only for celLXfs, not cellStyleXfs)
			stylesooxml.append( "<xf " );
			int[] refs = (int[]) cellxf;
			// all id refs are 0-based
			int ftId = refs[0];    // font ref
			int fId = refs[1];     // fill ref
			int bId = refs[2];     ///border ref
			int nId = refs[3];     // number format ref
			int ha = refs[4];
			int va = refs[5];
			int wr = refs[6];
			int ind = refs[7];
			int rot = refs[8];
			int hidden = refs[9];
			int locked = refs[10];
			int shrink = refs[11];
			int rtoleft = refs[12];

			if( nId > -1 )
			{
				stylesooxml.append( " numFmtId=\"" + nId + "\"" );
			}
			if( ftId > -1 )
			{
				stylesooxml.append( " fontId=\"" + ftId + "\"" );
			}
			if( fId > -1 )
			{
				stylesooxml.append( " fillId=\"" + fId + "\"" );
			}
			if( fId > 0 )
			{
				stylesooxml.append( " applyFill=\"1\"" );
			}
			if( bId > -1 )
			{
				stylesooxml.append( " borderId=\"" + bId + "\"" );
			}

			// TODO: shrinkToFit ...
			boolean alignblock = ((ha != 0) || (va != 2) || (wr != 0) || (ind != 0) || (rot != 0));
			boolean protectblock = ((hidden == 1) || (locked == 0));
			if( alignblock || protectblock )
			{
				stylesooxml.append( ">" );
			}
			if( (ha != 0) || (va != 2) || (wr != 0) || (ind != 0) || (rot != 0) )
			{
				stylesooxml.append( "<alignment" );
				if( ha != 0 )  //default=general
				{
					stylesooxml.append( " horizontal=\"" + horizontalAlignment[ha] + "\"" );
				}
				if( va != 2 )  //default= bottom
				{
					stylesooxml.append( " vertical=\"" + verticalAlignment[va] + "\"" );
				}
				if( wr == 1 )
				{
					stylesooxml.append( " wrapText=\"1\"" );
				}
				if( ind > 0 )
				{
					stylesooxml.append( " indent=\"" + ind + "\"" );
				}
				if( rot != 0 )
				{
					stylesooxml.append( " textRotation=\"" + rot + "\"" );
				}
				if( shrink == 1 )
				{
					stylesooxml.append( " shrinkToFit=\"1\"" );
				}
				if( rtoleft != 0 )
				{
					stylesooxml.append( " readingOrder=\"" + rtoleft + "\"" );
				}
				stylesooxml.append( "/>\r\n" );
			}
			if( (hidden == 1) || (locked == 0) )
			{ // if not the default protection settings, add protection element
				stylesooxml.append( "<protection" );
				if( hidden == 1 )
				{
					stylesooxml.append( " hidden=\"1\"" );
				}
				if( locked == 0 )
				{
					stylesooxml.append( " locked=\"0\"" );
				}
				stylesooxml.append( "/>\r\n" );
			}
			if( alignblock || protectblock )
			{
				stylesooxml.append( "</xf>\r\n" );
			}
			else
			{
				stylesooxml.append( "/>\r\n" );
			}
		}
		stylesooxml.append( "</cellXfs>" );
		stylesooxml.append( "\r\n" );

		// cellStyles -- for named styles -- NEEDED??

		// dxf's -- incremental style info
		if( bk.getWorkBook().getDxfs() != null )
		{
			ArrayList dxfs = bk.getWorkBook().getDxfs();
			if( dxfs.size() > 0 )
			{
				stylesooxml.append( "<dxfs count=\"" + dxfs.size() + "\">" );
				for( Object dxf : dxfs )
				{
					stylesooxml.append( ((Dxf) dxf).getOOXML() );
				}
				stylesooxml.append( "</dxfs>" );
			}
		}
		// NOTE: indexed colors are depreciated and represent the hard-coded default palate but necessary for proper color translation
		// Indexed Colors:
		stylesooxml.append( "<colors>" );
		stylesooxml.append( "\r\n" );
		stylesooxml.append( "<indexedColors>" );
		stylesooxml.append( "\r\n" );
		for( int i = 0; i < bk.getWorkBook().getColorTable().length; i++ )
		{
			stylesooxml.append( "<rgbColor rgb=\"" + "00" + FormatHandle.colorToHexString( bk.getWorkBook().getColorTable()[i] ).substring(
					1 ) + "\"/>" );
			stylesooxml.append( "\r\n" );
		}
		stylesooxml.append( "</indexedColors>" );
		stylesooxml.append( "\r\n" );
		stylesooxml.append( "</colors>" );
		stylesooxml.append( "\r\n" );
		stylesooxml.append( "</styleSheet>" );
		addDeferredFile( stylesooxml, "xl/styles.xml" );
	}

	/**
	 * Using the xf record passed in, populate the required variables for writing out to
	 * styles.xml
	 *
	 * @param xfs
	 * @param cellxfs
	 * @param fills
	 * @param borders
	 * @param numfmts
	 * @param fonts
	 */
	private void addXFToStyle( Xf xf, ArrayList cellxfs, ArrayList fills, ArrayList borders, ArrayList numfmts, ArrayList fonts )
	{

		int[] refs = new int[13];

		// fonts
		String s = xf.getFont().getOOXML();
		int id = fonts.indexOf( s );
		if( id == -1 )
		{
			fonts.add( s );
			refs[0] = fonts.size() - 1;
		}
		else
		{
			refs[0] = id;
		}

		// fills
		id = 0;
		if( xf.getFill() != null )
		{
			s = xf.getFill().getOOXML();
			id = fills.indexOf( s );
		}
		else if( xf.getFillPattern() > 0 )
		{
			s = Fill.getOOXML( xf );
			id = fills.indexOf( s );
		}
		if( id == -1 )
		{
			fills.add( s );
			refs[1] = fills.size() - 1;
		}
		else
		{
			refs[1] = id;
		}

		// borders
		s = Border.getOOXML( xf/*fh[i])*/ );
		id = borders.indexOf( s );
		if( id == -1 )
		{
			borders.add( s );
			refs[2] = borders.size() - 1;
		}
		else
		{
			refs[2] = id;
		}

		if( xf.getIfmt() > FormatConstants.BUILTIN_FORMATS.length )
		{    // only input user defined formats ...
			s = NumFmt.getOOXML( xf );
			id = numfmts.indexOf( s );
			if( id == -1 )
			{
				numfmts.add( s );
			}
		}

		refs[3] = xf.getIfmt();
		refs[4] = xf.getHorizontalAlignment();
		refs[5] = xf.getVerticalAlignment();
		refs[6] = xf.getWrapText() ? 1 : 0;
		refs[7] = xf.getIndent();
		refs[8] = xf.getRotation();
		refs[9] = xf.isFormulaHidden() ? 1 : 0;
		refs[10] = xf.isLocked() ? 1 : 0;
		refs[11] = xf.isShrinkToFit() ? 1 : 0;
		refs[12] = xf.getRightToLeftReadingOrder();

		cellxfs.add( refs );        // link formatHandles to referenced formats

	}

	/**
	 * Creates workbook.xml containing worksheet information, and writes it out to
	 * the root of the OPC ZIP
	 * <p/>
	 * Format of the workbook.xml:
	 * (all elements are OPTIONAL except for <sheets>)
	 * <p/>
	 * <bookViews>       Window Position and height/width, Filter Options .. No limit to how many are defined
	 * <calcPr>          Stores calculation status and details
	 * <customWorkbookViews>
	 * <definedNames><definedName name="" [comment="" hidden="1/0" localSheetId="" for external refs]>RANGE or FORMULA</definedName>
	 * <extLst>          Future Feature- Data Storage Area
	 * <externalReferences>
	 * <fileRecoveryPr>  File Recovery Properties
	 * <fileSharing>     Specifies pwd + username
	 * <fileVersion>     tracks versions
	 * <functionGroups>
	 * <oleSize>         Embedded Object Size
	 * <pivotCaches>     Represents a cache of data for pivot tables and formulas
	 * <sheets><sheet r:id="relationship id" name="unique sheet name" sheetId="#" [state="visible"] /></sheets>          -- required
	 * <smartTagPr>
	 * <webPublishing>   Attibutes related to publishing on the web
	 * <workbookPr>      Workbook Properties: date1904,showObjects ...
	 * <workbookProtection>
	 *
	 * @param bk
	 */
	private void writeWorkBookOOXML( WorkBookHandle bk ) throws IOException
	{

		// create Zip Entry
		nextZipEntry( "xl/workbook.xml" );

		// WORKBOOK.XML
		writer.write( xmlHeader );
		writer.write( "\r\n" );

		// namespace
		writer.write( ("<workbook xmlns=\"" + xmlns + "\" xmlns:r=\"" + relns + "\">") );
		writer.write( "\r\n" );

		// IF MACRO-ENABLED, MUST HAVE CODENAME       // TODO: other attributes 
		if( bk.getWorkBook().getCodename() != null )
		{
			writer.write( ("<workbookPr codeName=\"" + bk.getWorkBook().getCodename() + "\"/>") );
			writer.write( "\r\n" );
		}

		// BOOKVIEW
		writer.write( "<bookViews>" );
		writer.write( "\r\n" );
		// TODO: Only 1?  TODO: Handle other workbookview options
		writer.write( "<workbookView" );
		writer.write( (" firstSheet=\"" + bk.getWorkBook().getFirstSheet() + "\"") );
		writer.write( (" activeTab=\"" + bk.getWorkBook().getSelectedSheetNum() + "\"") );
		if( !bk.showSheetTabs() )
		{
			writer.write( " showSheetTabs=\"0\"" );
		}
		writer.write( "/>\r\n" );
		writer.write( "</bookViews>" );
		writer.write( "\r\n" );

		// IDENTIFY SHEETS
		writer.write( "<sheets>" );
		writer.write( "\r\n" );
		WorkSheetHandle[] wsh = bk.getWorkSheets();
		for( int i = 0; i < wsh.length; i++ )
		{
			String s = "sheet" + (i + 1);        //Write SheetXML to SheetX.xml, 1-based
			writer.write( ("<sheet name=\"" + stripNonAscii( wsh[i].getSheetName() ) + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"") );
			if( wsh[i].getVeryHidden() )
			{
				writer.write( " state=\"veryHidden\"" );
			}
			else if( wsh[i].getHidden() )
			{
				writer.write( " state=\"hidden\"" );
			}
			writer.write( "/>" );
			writer.write( "\r\n" );
			wbContentList.add( i, new String[]{
					"/xl/worksheets/" + s + ".xml", "sheet"
			} ); // make sure rId in workbook.xml matches workbook.xml.rels
		}
		writer.write( "</sheets>" );
		writer.write( "\r\n" );

		int rId = wsh.length;  // start counting after sheets
		// EXTERNAL LINKS AFTER SHEETS
		if( getExternalRefType( "externalLink", bk.getWorkBook().getOOXMLObjects() ).size() > 0 )
		{ // has external refs
			ArrayList refs = getExternalRefType( "externalLink", bk.getWorkBook().getOOXMLObjects() );
			writer.write( "<externalReferences>" );
			writer.write( "\r\n" );
			for( int i = 0; i < refs.size(); i++ )
			{
				String[] r = (String[]) refs.get( i );
				r[3] = "rId" + (rId + 1);  // ensure rId is correct
				refs.remove( i );
				refs.add( i, r );
				writer.write( ("<externalReference r:id=\"rId" + (++rId) + "\"/>") );
				writer.write( "\r\n" );
			}
			writer.write( "</externalReferences>" );
			writer.write( "\r\n" );
		}

		// NAMES AFTER EXTERNAL REFS
		// TODO:  add handling for name parameters
		Name[] names = bk.getWorkBook().getNames();
		if( (names != null) && (names.length > 0) )
		{
			writer.write( "<definedNames>" );
			writer.write( "\r\n" );
			for( Name name : names )
			{
				String s = stripNonAsciiRetainQuote( name.getExpressionString().substring( 1 ) ).toString(); //avoid "="
				if( (s != null) && (s.length() != 0) && !s.startsWith( "#REF!" ) )
				{
					if( !name.isBuiltIn() )
					{
						writer.write( ("<definedName name=\"" + stripNonAscii( name.toString() ) + "\"") );
						if( name.getItab() > 0 )
						{
							writer.write( (" localSheetId=\"" + (name.getItab() - 1) + "\">") );
						}
						else
						{
							writer.write( (">") );
						}
					}
					else
					{
						//if (names[i].getBuiltInType()==Name.PRINT_TITLES) { // must set localsheetid
						writer.write( ("<definedName name=\"" + builtInNames[name.getBuiltInType()] + "\"") );
						if( name.getItab() > 0 )
						{
							writer.write( (" localSheetId=\"" + (name.getItab() - 1) + "\">") );
						}
						else
						{
							writer.write( (">") );
						}
					}
					writer.write( s );
					writer.write( "</definedName>" );
					writer.write( "\r\n" );
				}
			}
			writer.write( "</definedNames>" );
			writer.write( "\r\n" );
		}

		writer.write( "</workbook>" );
		writer.write( "\r\n" );

		// add workbook.xml to content list (for [Content_Types].xml)
		if( format == WorkBookHandle.FORMAT_XLTM )
		{ // macro-enabled template
			mainContentList.add( new String[]{ "/xl/workbook.xml", "documentTemplateMacroEnabled" } );
		}
		else if( (format == WorkBookHandle.FORMAT_XLSM) || hasMacros( bk ) )
		{// format can be XLSM even though it does not contain macros ((:
			mainContentList.add( new String[]{ "/xl/workbook.xml", "documentMacroEnabled" } );
			format = WorkBookHandle.FORMAT_XLSM;   // ensure flag is set properly - macro-enabled workbook
		}
		else if( format == WorkBookHandle.FORMAT_XLTX )
		{  // template
			mainContentList.add( new String[]{ "/xl/workbook.xml", "documentTemplate" } );
		}
		else                                             // regular xlsx workbook
		{
			mainContentList.add( new String[]{ "/xl/workbook.xml", "document" } );
		}

		writeRels( wbContentList, "xl/_rels/workbook.xml.rels" );   // write workbook.xml.rels
	}

	/**
	 * Handles preliminary writing of worksheets, essentially everything before <row> elements start
	 *
	 * @param sheet
	 * @param bk
	 * @param id
	 * @throws IOException
	 */
	protected void writeSheetPrefix( WorkSheetHandle sheet, WorkBookHandle bk, int id ) throws IOException
	{
		String slx = "xl/worksheets/sheet" + (id + 1) + ".xml";
		nextZipEntry( slx );

		writer.write( xmlHeader );
		writer.write( "\r\n" );
		writer.write( ("<worksheet xmlns=\"" + xmlns + "\" xmlns:r=\"" + relns + "\">") );

		if( sheet.getMysheet().getSheetPr() != null )       // sheet properties
		{
			writer.write( sheet.getMysheet().getSheetPr().getOOXML() );
		}

		// dimensions    // TODO: Deal with MAXROWS MAXCOLS in various Excel Versions
		int last = sheet.getLastCol() - 1;
		if( last == (WorkBook.MAXCOLS - 1) )
		{
			last = XLSConstants.MAXCOLS - 1;    // 20081204 KSC: eventually WorkBook.MAXCOLS will == Excel-7 MAXCOLS, but until, convert
		}
		if( last < 0 )
		{
			last = 0;
		}
		String d = ExcelTools.formatLocation( new int[]{ sheet.getFirstRow(), sheet.getFirstCol() } ) + ":" +
				ExcelTools.formatLocation( new int[]{ sheet.getLastRow(), last } );
		writer.write( ("<dimension ref=\"" + d + "\"/>") );
		writer.write( "\r\n" );

		// Sheet View Properties 
		writer.write( "<sheetViews>" );
		writer.write( "\r\n" );    // TODO: it's possible to have multiple sheetViews     
		if( sheet.getMysheet().getSheetView() == null )
		{
			sheet.getMysheet().setSheetView( new SheetView() );
		}

		// TODO: finish options:  colorId, defaultGridColor, rightToLeft
		// showFormulas, showRuler, showWhiteSpace, view
		// zoomScaleNormal, zoomScalePageLayoutView, zoomScaleSheetLayoutView
		// showFormulas, default= false
		if( !sheet.getShowGridlines() )
		{
			sheet.getMysheet().getSheetView().setAttr( "showGridLines", "0" ); // default= true
		}
		if( !sheet.getShowSheetHeaders() )
		{
			sheet.getMysheet().getSheetView().setAttr( "showRowColHeaders", "0" ); // default= true
		}
		if( !sheet.getShowZeroValues() )
		{
			sheet.getMysheet().getSheetView().setAttr( "showZeros", "0" ); // default= true
		}
		// rightToLeft
		if( sheet.getSelected() )
		{
			sheet.getMysheet().getSheetView().setAttr( "tabSelected", "1" );           // default= false
		}
		else
		{
			sheet.getMysheet().getSheetView().removeSelection();  // in case previously selected, remove any seletions
		}
//       if (sheet.getTopLeftCell()!=null) { sheet.getMysheet().getSheetView().setAttr("topLeftCell", sheet.getTopLeftCell()); }

		// showRuler
//       if (!sheet.getShowOutlineSymbols()) sheet.getMysheet().getSheetView().setAttr("showOutlineSymbols", "0");    // default= true
		// defaultGridColor, showWhiteSpace, view, topLeftCell, colorId
		if( sheet.getZoom() != 1.0 )
		{
			sheet.getMysheet().getSheetView().setAttr( "zoomScale", String.valueOf( new Double( sheet.getZoom() * 100 ).intValue() ) );
		}

		// zoomScalePageLayoutView, zoomScaleSheetLayoutView
		sheet.getMysheet().getSheetView().setAttr( "workbookViewId",
		                                           "0" );        // TODO: may be other workbookviews, can't always assume 0                     

		writer.write( sheet.getMysheet().getSheetView().getOOXML() );
		writer.write( "\r\n" );
		writer.write( "</sheetViews>" );
		writer.write( "\r\n" );

		// Sheet Format Properties
		writer.write( "<sheetFormatPr" );
		if( sheet.getMysheet().getDefaultColumnWidth() > -1 )
		{
			writer.write( (" defaultColWidth=\"" + sheet.getMysheet().getDefaultColumnWidth() + "\"") );
		}
		writer.write( (" defaultRowHeight=\"" + sheet.getMysheet().getDefaultRowHeight() + "\"") );    // required
		if( sheet.getMysheet().hasCustomHeight() )
		{
			writer.write( " customHeight=\"1\"" );
		}
		if( sheet.getMysheet().hasZeroHeight() )
		{
			writer.write( " zeroHeight=\"1\"" );
		}
		if( sheet.getMysheet().hasThickTop() )
		{
			writer.write( " thickTop=\"1\"" );
		}
		if( sheet.getMysheet().hasThickBottom() )
		{
			writer.write( " thickBottom=\"1\"" );
		}
		writer.write( "/>" );
		writer.write( "\r\n" );

		// Columns
		writer.write( getColOOXML( bk, sheet ).toString() );

		// Sheet Data - rows and cells 
		writer.write( "<sheetData>" );
		writer.write( "\r\n" );

	}

	/**
	 * writeSheetOOXML writes XML data in SheetML specificaiton to [worksheet].xml
	 * <p/>
	 * Main portion is the <sheetData> section, containing all row and cell information
	 * <p/>
	 * format of SheetML:
	 * <worksheet>     ROOT
	 * autoFilter (AutoFilter Settings)   Hides rows based upon criteria
	 * cellWatches (Cell Watch Items)
	 * colBreaks (Vertical Page Breaks)
	 * cols (Column Information)
	 * conditionalFormatting (Conditional Formatting)
	 * controls (Embedded Controls)
	 * customProperties (Custom Properties)
	 * customSheetViews (Custom Sheet Views)
	 * dataConsolidate (Data Consolidate)
	 * dataValidations (Data Validations)
	 * dimension (Worksheet Dimensions)
	 * drawing (Drawing)
	 * extLst (Future Feature Data Storage Area)
	 * headerFooter (Header Footer Settings)
	 * hyperlinks (Hyperlinks)
	 * ignoredErrors (Ignored Errors)
	 * legacyDrawing (Legacy Drawing Reference)
	 * legacyDrawingHF (Legacy Drawing Reference in Header Footer)
	 * mergeCells (Merge Cells)
	 * oleObjects (Embedded Objects)
	 * pageMargins (Page Margins)
	 * pageSetup (Page Setup Settings)
	 * phoneticPr (Phonetic Properties)
	 * picture (Background Image)
	 * printOptions (Print Options)
	 * protectedRanges (Protected Ranges)
	 * rowBreaks (Horizontal Page Breaks (Row))
	 * scenarios (Scenarios)
	 * sheetCalcPr (Sheet Calculation Properties)
	 * sheetData (Sheet Data)
	 * sheetFormatPr (Sheet Format Properties)
	 * sheetPr (Sheet Properties)
	 * sheetProtection (Sheet Protection Options)
	 * sheetViews (Sheet Views)
	 * smartTags (Smart Tags)
	 * sortState (Sort State)
	 * tableParts (Table Parts)
	 * webPublishItems (Web Publishing Items)  *
	 * <p/>
	 * <p/>
	 * ORDER OF THE ABOVE:
	 * (sheet properties -- all optional)
	 * <sheetPr filterMode="1"/>       indicates that an autofilter has been applied
	 * <dimension ref="RANGE"/>        indicates the used range on the sheet (there should be no data or formulas outside this range)
	 * <sheetViews>                    indicates which cell and sheet are active
	 * <sheetView tabSelected="1" workbookViewId="0">
	 * <selection activeCell="B3" sqref="B3"/>
	 * </sheetView>
	 * </sheetViews>
	 * <sheetFormatPr defaultRowHeight="15"/>      default row height
	 * <cols>
	 * <col min="1" max="1" width="12.85546875" bestFit="1" customWidth="1"/>
	 * <col min="3" max="3" width="3.28515625" customWidth="1"/>
	 * <col min="4" max="4" width="11.140625" bestFit="1" customWidth="1"/>
	 * <col min="8" max="8" width="17.140625" style="1" customWidth="1"/>
	 * </cols>
	 * <p/>
	 * **     <sheetData>     cell table - specifies rows and cells with types and values (see below) -- REQUIRED
	 * <p/>
	 * ("Supporting Features" -- all optional)
	 * <sheetProtection objects="0" scenarios="0"/>
	 * <autoFilter ref="D5:H11">
	 * <filterColumn colId="0">
	 * <customFilters and="1">
	 * <customFilter operator="greaterThan" val="0"/>
	 * <mergeCells>
	 * <phoneticPr>
	 * <conditionalFormatting>
	 * <printOptions/>
	 * <dataValidations>
	 * <hyperlinks>
	 * <printOptions>
	 * <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
	 * <pageSetup orientation="portrait" horizontalDpi="300" verticalDpi="300"/>
	 * <headerFooter>
	 * <rowBreaks>
	 * <colBreaks>
	 * <customProperties>
	 * <cellWatches>
	 * <ignoredErrors>
	 * <smartTags>
	 * <drawing>
	 * <legacyDrawing>
	 * <legacyDrawingHF>
	 * <picture>
	 * <oleObjects>
	 * <controls>
	 * <webPublishItems>
	 * <tableParts>
	 * <extLst>
	 *
	 * @param bk
	 * @param sheet
	 */
	protected void writeSheetOOXML( WorkBookHandle bk, WorkSheetHandle sheet, int id ) throws IOException
	{
		// Sst sst= bk.getWorkBook().getSharedStringTable();
		ArrayList hyperlinks = new ArrayList();
		// SHEETxx.XML
		writeSheetPrefix( sheet, bk, id );
		RowHandle[] rows = sheet.getRows();
		for( RowHandle row : rows )
		{
			try
			{ // note: row #, col #'s are 1-based, sst and style index are 0-based
				writeRow( row, hyperlinks );
				//} catch (RowNotFoundException re) {
				; // do nothing
			}
			catch( Exception e )
			{
				log.error( "OOXMLWriter.writeSheetOOXML writing rows: " + e.toString() );
				e.printStackTrace();
			}
		}
		writer.write( "</sheetData>" );
		writer.write( "\r\n" );

		// after sheetData include "supporting features"
		// *******************************************************************************************************
		// In Order:
		// sheetCalcPr
		// sheetProtection
		if( sheet.getProtected() )
		{
			String pwd = sheet.getHashedProtectionPassword();
			if( pwd != null )
			{
				writer.write( ("<sheetProtection password=\"" + sheet.getHashedProtectionPassword() + "\" sheet=\"1\" objects=\"1\" scenarios=\"1\"/>") );
			}
			else
			{
				writer.write( ("<sheetProtection sheet=\"1\" objects=\"1\" scenarios=\"1\"/>") );
			}
			writer.write( "\r\n" );
		}

		// protectedRanges
		// scenarios
		// autoFilter
		if( sheet.getMysheet().getOOAutoFilter() != null ) // TODO: Merge with 2003 AutoFilter
		{
			writer.write( sheet.getMysheet().getOOAutoFilter().getOOXML() );
		}
		// sortState
		// dataConsolidation
		// customSheetViews
		// mergeCells
		writeMergedCellRecords( sheet );

		// phoneticPr
		// conditionalFormatting
		if( sheet.getMysheet().getConditionalFormats() != null )
		{
			List condfmts = sheet.getMysheet().getConditionalFormats();
			int[] priority = new int[1];
			priority[0] = 1;
			for( Object condfmt : condfmts )
			{
				String cfmt = ((Condfmt) condfmt).getOOXML( bk, priority );
				writer.write( cfmt );
				writer.write( "\r\n" );
			}
		}
		// dataValidations
		if( sheet.hasDataValidations() )
		{
			writer.write( sheet.getMysheet().getDvalRec().getOOXML() );
		}
		// hyperlinks
		if( hyperlinks.size() > 0 )
		{
			writer.write( "<hyperlinks>" );
			writer.write( "\r\n" );
			for( Object hyperlink : hyperlinks )
			{
				String[] s = (String[]) hyperlink;
				if( !s[2].equals( "" ) ) // has a description
				{
					writer.write( "<hyperlink ref=\"" + s[0] + "\" r:id=\"rId" + (shContentList.size() + 1) + "\" display=\"" + s[2] + "\"/>" );
				}
				else
				{
					writer.write( "<hyperlink ref=\"" + s[0] + "\" r:id=\"rId" + (shContentList.size() + 1) + "\"/>" );
				}
				writer.write( "\r\n" );
				shContentList.add( new String[]{ s[1], "hyperlink" } );
			}
			hyperlinks.clear();
			writer.write( "</hyperlinks>" );
			writer.write( "\r\n" );
		}

		// now handle any external ooxml object references (in required order):
		// printerOptions
		// pageMargin
		// pageSetup

		// headerFooter
		// rowBreaks
		// colBreaks
		// customProperties
		// cellWatches
		// ignoredErrors
		// smartTags
		List externalOOXML = sheet.getMysheet().getOOXMLObjects();
		// printerOptions
		writeSheetLevelExternalReferenceOOXML( writer, "printerSettings", externalOOXML );
		// drawing objects linked to this sheet - charts, shapes ...
		writeDrawingObjects( writer, sheet, bk );
		// legacy drawing object (vml)
		if( writeLegacyDrawingObjects( writer, sheet, bk ) )
		{// legacyDrawing objects, linked to either oleObject embeds or control embeds
			writer.write( ("<legacyDrawing r:id=\"rId" + (shContentList.size()) + "\"/>\r\n") );
		}
		//  oleObjects
		writeSheetLevelExternalReferenceOOXML( writer, "oleObject", externalOOXML );
		// activeX objects
		writeSheetLevelExternalReferenceOOXML( writer, "activeX", externalOOXML );
		// comments (notes)
		writeComments( writer, sheet, bk );

		writer.write( "</worksheet>" ); // 20081028 KSC: Sheet xml should
		// be named SheetX.xml instead
		// of sheetname (see
		// writeWbOOXML as well)
		// finished, now write to <sheet>.xml + <sheet>.xml.rels, if has
		// associated content

		// write rels if necessary (printer settings, drawings ...)
		writeRels( shContentList, "xl/worksheets/_rels/sheet" + (id + 1) + ".xml.rels" );
		sheetsContentList.addAll( shContentList ); // since shContentList will be cleared out for each sheet, make sure sheet contents are stored
		// clear out content list for the sheet [associated documents for the
		// sheet, will be written to <SheetX>.xml.rels]
		shContentList.clear();
	}

	/**
	 * Write the merged cell records for a worksheet
	 */
	protected void writeMergedCellRecords( WorkSheetHandle sheet ) throws IOException
	{
		List mcs = sheet.getMysheet().getMergedCells(); // TODO: PROBLEM:
		// getMergedCellRecs
		// contains null values,
		// screwing up output
		// ... fix!!
		// Use getMergedCells method - which DOESN'T add new blank merged cell
		if( (mcs != null) && (mcs.size() > 0) )
		{
			StringBuffer mc = new StringBuffer();
			int cnt = 0;
			for( Object mc1 : mcs )
			{
				CellRange[] cr = ((Mergedcells) mc1).getMergedRanges();
				if( cr != null )
				{
					for( CellRange aCr : cr )
					{
						String rng = aCr.getRange();
						if( rng != null )
						{
							int z = rng.indexOf( "!" ); // strip sheetname
							mc.append( "<mergeCell ref=\"" + aCr.getRange().substring( z + 1 ) + "\"/>" );
							mc.append( "\r\n" );
							cnt++;
						}
					}
				}
			}
			if( mc.length() > 0 )
			{ // Only input element if actual valid merged
				// cell ranges exist,
				writer.write( ("<mergeCells count=\"" + cnt + "\">") );
				writer.write( "\r\n" );
				writer.write( mc.toString() );
				writer.write( "</mergeCells>" );
				writer.write( "\r\n" );
			}
		}

	}

	/**
	 * Writes a row and all contents to the zip output
	 *
	 * @param hyperlinks
	 * @throws IOException
	 * @throws FormulaNotFoundException
	 */
	public void writeRow( RowHandle row, ArrayList hyperlinks ) throws IOException
	{

		// row = sheet.getRow(x);
		// TODO: need spans?  
		// <row element> -- eventually will be refactored as Object
		String h = "";
		if( row.getHeight() != 255 )    // if it's not default
		{
			h = " ht=\"" + row.getHeight() / rowHtFactor + "\" customHeight=\"1\"";
		}
		writer.write( ("<row r=\"" + (row.getRowNumber() + 1) + "\"" + h) );
		if( (row.getFormatId() > 0) && (row.getFormatId() > row.getWorkBook().getWorkBook().getDefaultIxfe()) )      // row-level formatting specified
		{
			writer.write( (" s=\"" + row.getFormatId() + "\" customFormat=\"1\"") );
		}
		if( row.getHasAnyThickTopBorder() )
		{
			writer.write( " thickTop=\"1\"" );
		}
		if( row.getHasAnyBottomBorder() )
		{
			writer.write( " thickBot=\"1\"" );
		}

		if( row.isHidden() )
		{
			writer.write( " hidden=\"1\"" );
		}
		if( row.isCollapsed() )
		{
			writer.write( " collapsed=\"1\"" );                // 20090513 KSC: Added collapsed, outlineLevel [BUGTRACKER 2371]
		}
		if( row.getOutlineLevel() != 0 )
		{
			writer.write( (" outlineLevel=\"" + row.getOutlineLevel() + "\"") );
		}
		writer.write( ">" );
		writer.write( "\r\n" );
		// Cell element <c
		CellHandle[] ch = row.getCells();
		// iterate cells and output xml
		for( CellHandle aCh : ch )
		{
			int styleId = aCh.getCell().getIxfe();
			int dataType = aCh.getCellType();
			if( aCh.hasHyperlink() )
			{   // save; hyperlinks go after sheetData
				hyperlinks.add( new String[]{ aCh.getCellAddress(), aCh.getURL(), aCh.getURLDescription() } );
			}
			writer.write( ("<c r=\"" + aCh.getCellAddress() + "\"") );
			if( styleId > 0 )
			{
				writer.write( (" s=\"" + styleId + "\"") );
			}
			switch( dataType )
			{
				case CellHandle.TYPE_STRING:
					String s = aCh.getStringVal();
					boolean isErrVal = false;
					if( (s.indexOf( "#" ) == 0) )
					{   // 20090521 KSC: must test if it's an error string value
						isErrVal = (Collections.binarySearch( Arrays.asList( new String[]{
								"#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"
						} ), s.trim() ) > -1);
					}
					if( !isErrVal )
					{
						writer.write( " t=\"s\">" );
						// 20090520 KSC: can't use only string value as intra-cell formatting is possible zip.write("<v>" + ssts.indexOf(stripNonAscii(ch[j].getStringVal())) + "</v>");
						int v = ((Labelsst) aCh.getCell()).isst; // use isst instead of a lookup -- MUCHMUCHMUCH faster!
						writer.write( ("<v>" + v + "</v>") );
					}
					else
					{// it's an error value, must have type of "e"
						writer.write( " t=\"e\">" );
						writer.write( ("<v>" + s + "</v>") );
					}
					break;
				case CellHandle.TYPE_DOUBLE:
				case CellHandle.TYPE_FP:
				case CellHandle.TYPE_INT:
					writer.write( (" t=\"n\"><v>" + aCh.getVal() + "</v>") );
					break;
				case CellHandle.TYPE_FORMULA:
					FormulaHandle fh;
					try
					{
						fh = aCh.getFormulaHandle();
						writer.write( fh.getOOXML() );
					}
					catch( FormulaNotFoundException e )
					{
						log.error( "Error getting formula handle in OOXML Writer" );
					}
					break;
				case CellHandle.TYPE_BOOLEAN:
					writer.write( (" t=\"b\"><v>" + aCh.getIntVal() + "</v>") );
					break;
				case CellHandle.TYPE_BLANK:
					writer.write( ">" );
					break;
			}
			writer.write( "</c>" );
			writer.write( "\r\n" );
		}
		writer.write( "</row>" );
		writer.write( "\r\n" );
	}

	/**
	 * retrieves the column ooxml possible attributes: bestFit, collapsed,
	 * customWidth, hidden, max, min, style, width
	 *
	 * @param bk
	 * @param sheet
	 * @return
	 */
	private StringBuffer getColOOXML( WorkBookHandle bk, WorkSheetHandle sheet )
	{
		StringBuffer colooxml = new StringBuffer();
		// ColHandle cols[]= sheet.getColumns();
		Iterator<Colinfo> iter = sheet.getMysheet().getColinfos().iterator();
		if( iter.hasNext() )
		{
			colooxml.append( "<cols>" );
			colooxml.append( "\r\n" );
			while( iter.hasNext() )
			{
				try
				{
					// col width = width/256???
					Colinfo c = iter.next();
					int collast = c.getColLast() + 1;
					double w = (c.getColWidth() / colWFactor);
					colooxml.append( "<col min=\"" + (c.getColFirst() + 1) + "\" max=\"" + collast + "\" width=\"" + (c.getColWidth() / colWFactor) + "\" customWidth=\"1\"" );
					if( c.isHidden() )
					{
						colooxml.append( " hidden=\"1\"" );
					}
					if( c.getIxfe() > 0 ) // column-level formatting specified
					{
						colooxml.append( " style=\"" + c.getIxfe() + "\"" );
					}

					colooxml.append( "/>" );
					colooxml.append( "\r\n" );
				}
				catch( Exception e )
				{
					log.warn("getColOOXML failed", e);
				}
			}
			colooxml.append( "</cols>" );
			colooxml.append( "\r\n" );
		}
		return colooxml;
	}

	/**
	 * write all necessary information describing ImageHandle im
	 * including .rels and drawing xml
	 *
	 * @param im
	 */
	private String getImageOOXML( ImageHandle im )
	{
		String ret = "";
		try
		{
			ret = im.getOOXML( drContentList.size() + 1 );   // obtain image OOXML; if errors don't write out to zip or contents
			// write image bytes to file
			String ext = im.getType();                   // 20090117 KSC: may come back as undefined, usually because it's an EMF and we can't process right now ...
			if( ext.equals( "undefined" ) )
			{
				ext = "emf";
			}
			String imageName = "image" + (++imgId) + "." + ext;
			// TODO: HANDLE REUSING IMAGES!!!!!!!!!!!!!!!!!!!!!
			addDeferredFile( im.getImageBytes(), mediaDir + "/" + imageName );

			// trap package contents for drawing.xml
			drContentList.add( new String[]{ "/" + mediaDir + "/" + imageName, "image" } );
		}
		catch( Exception e )
		{
			log.error( "OOXMLWriter.getImageOOXML: " + e.toString() );
			ret = "";
		}
		return ret;
	}

	/**
	 * write out the drawing objects - images and charts + drawingml
	 *
	 * @param out
	 * @param sheet
	 * @param bk
	 * @throws IOException
	 */
	private void writeDrawingObjects( Writer out, WorkSheetHandle sheet, WorkBookHandle bk ) throws IOException
	{
		// Drawing Objects  - Images, Charts & Shapes (OOXML-specific) = drawing, legacyDrawing, legacyDrawingHF, picture, oleObjects, controls
		StringBuffer drawing = new StringBuffer();
		ImageHandle[] imz = sheet.getImages();
		if( imz.length > 0 )
		{
			// For each image, create a Drawing reference in sheet xml + write imageOOXML to drawingX.xml
			for( ImageHandle anImz : imz )
			{
				// obtain image OOXML + write image file to ZIP
				drawing.append( getImageOOXML( anImz ) );
				drawing.append( "\r\n" );
			}
		}

		List charts = sheet.getMysheet().getCharts();
		int nUserShapes = 0; // if charts link to any drawing ml user shapes, see below for handling
		if( charts.size() > 0 )
		{
			// for each chart, create a chart.xml + trap references for drawingX.xml.rels 
			for( Object chart : charts )
			{
				try
				{   // obtain image OOXML + write image file to ZIP
					Chart c = (Chart) chart;
					drawing.append( getChartDrawingOOXML( new ChartHandle( c, bk ) ) );
					drawing.append( "\r\n" );
					if( c instanceof OOXMLChart )
					{
						ArrayList chartEmbeds = ((OOXMLChart) chart).getChartEmbeds();
						if( chartEmbeds != null )
						{
							int origDrawingId = drawingId;            // id for THIS CURRENT DRAWING ML describing this chart(s), etc.
							ArrayList chContentList = new ArrayList();
							for( Object chartEmbed : chartEmbeds )
							{
								// obtain external drawingml file(s) which define shape and write to zip
								String[] embed = (String[]) chartEmbed;
								if( embed[0].equals( "userShape" ) )
								{
									drawingId += (nUserShapes + 1);    // id for USER SHAPES drawingml 
									String f = embed[1];
									nUserShapes++;    // keep track of increment 
									writeExOOXMLFile( new String[]{ embed[0], drawingDir + "/", f }, chContentList );
								}
								else if( embed[0].equals( "image" ) )
								{
									String f = embed[1];
									writeExOOXMLFile( new String[]{ embed[0], mediaDir + "/", f }, chContentList );
								}
								else if( embed[0].equals( "themeOverride" ) )
								{
									String f = embed[1];
									writeExOOXMLFile( new String[]{ embed[0], themeDir + "/", f }, chContentList );
								}
							}
							writeRels( chContentList, chartDir + "/_rels/chart" + chartId + ".xml.rels" );
							sheetsContentList.addAll( chContentList ); // ensure embeds are written to main [Content_Types].xml 
							drawingId = origDrawingId;            // reset to id for THIS CURRENT DRAWING ML        			   
						}
					}
				}
				catch( Exception e/*ChartNotFoundException c*/ )
				{
					log.error( "OOXMLWriter.writeDrawingObjects failed getting Chart: " + e.toString() );
				}
			}
		}

		// OOXML shapes OTHER THAN vmldrawings, which are handled elsewhere (in writelegacyDrawingObjects) 
		if( sheet.getMysheet().getOOXMLShapes() != null )
		{
			HashMap shapes = sheet.getMysheet().getOOXMLShapes();
			Iterator i = shapes.keySet().iterator();
			while( i.hasNext() )
			{
				String key = (String) i.next();
				if( key.equals( "vml" ) )
				{
					continue; // handled in writelegacyDrawingObjects
				}
				Object o = shapes.get( key );
				if( o instanceof TwoCellAnchor )
				{
					TwoCellAnchor t = (TwoCellAnchor) o;
					drawing.append( t.getOOXML() );
					if( t.getEmbedFilename() != null )
					{    // shape has embedded images we must also store
						String f = t.getEmbedFilename();
						int rId = writeExOOXMLFile( new String[]{ "image", mediaDir + "/", externalDir + f }, drContentList );
						t.setEmbed( "rId" + rId );
					}
				}
				else if( o instanceof OneCellAnchor )
				{
					OneCellAnchor oca = (OneCellAnchor) o;
					drawing.append( oca.getOOXML() );
					// TODO: trap embedded images as in twoCellAnchor
					//drContentList.add(new String[] { "/" + mediaDir + "/"+ imageName, "image"});
				}
				drawing.append( "\r\n" );
			}
		}

		//  
		if( drawing.length() > 0 )
		{   // then have drawing objects to write out
			// one drawingX.xml per sheet, write reference in sheet
			out.write( ("<drawing r:id=\"rId" + (shContentList.size() + 1) + "\"/>") );
			out.write( "\r\n" );    // link drawing.xml to specific image
			// write out drawingml
			StringBuffer drawingml = new StringBuffer();
			drawingml.append( xmlHeader );
			drawingml.append( "\r\n" );
			drawingml.append( "<xdr:wsDr xmlns:xdr=\"" + drawingns + "\" xmlns:a=\"" + drawingmlns + "\">" );
			drawingml.append( "\r\n" );
			drawingml.append( drawing );
			drawingml.append( "</xdr:wsDr>" );
			drawingml.append( "\r\n" );
			// write drawingX.xml to Zip

			// FIX
			addDeferredFile( drawingml, drawingDir + "/drawing" + (++drawingId) + ".xml" );

			shContentList.add( new String[]{ "/" + drawingDir + "/drawing" + (drawingId) + ".xml", "drawing" } );

			// write drawingX.xml.rels to Zip
			writeRels( drContentList, drawingDir + "/_rels/drawing" + drawingId + ".xml.rels" );

			sheetsContentList.addAll( drContentList );
			drContentList.clear();
			drawingId += nUserShapes;  // if wrote drawingml for usershapes, must increment next drawingId to AFTER these files
		}
	}

	/**
	 * write all necessary information describing ChartHandle ch
	 * including .rels and drawing xml
	 *
	 * @param ch
	 * @return
	 */
	private String getChartDrawingOOXML( ChartHandle ch )
	{
		String ret = "";
		try
		{
			// trap package contents for drawing.xml
			StringBuffer chartml = new StringBuffer();
			chartml.append( xmlHeader );
			chartml.append( "\r\n" );
			chartml.append( "<c:chartSpace xmlns:c=\"" + chartns + "\" xmlns:a=\"" + drawingmlns + "\" xmlns:r=\"" + relns + "\">" );
			chartml.append( "\r\n" );
			chartml.append( ch.getOOXML( drContentList.size() + 1 ) );
			chartml.append( "</c:chartSpace>" );
			chartml.append( "\r\n" );
			ret = ch.getChartDrawingOOXML( drContentList.size() + 1 );   // if errors out don't write/add to zip
			addDeferredFile( chartml, chartDir + "/chart" + (++chartId) + ".xml" );
			drContentList.add( new String[]{ "/" + chartDir + "/chart" + chartId + ".xml", "chart" } );
		}
		catch( Exception e )
		{
			log.error( "er: " + e.toString() );
			ret = "";
		}
		return ret;
	}

	/**
	 * write notes or comments for the specific sheet
	 *
	 * @param zip
	 * @param sheet
	 * @param bk
	 * @throws IOException
	 */
	private void writeComments( Writer out, WorkSheetHandle sheet, WorkBookHandle bk ) throws IOException
	{
		CommentHandle[] nh = sheet.getCommentHandles();
		if( (nh == null) || (nh.length == 0) )
		{
			return;
		}
		StringBuffer comments = new StringBuffer();
		comments.append( xmlHeader );
		comments.append( "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n<authors>" );
		// run thru 1x to get authors
		ArrayList authors = new ArrayList();
		for( CommentHandle aNh1 : nh )
		{
			if( !authors.contains( aNh1.getAuthor() ) )
			{
				comments.append( "\r\n<author>" + OOXMLAdapter.stripNonAscii( aNh1.getAuthor() ) + "</author>" );
				authors.add( aNh1.getAuthor() );
			}
		}
		comments.append( "\r\n</authors>\r\n<commentList>" );
		for( CommentHandle aNh : nh )
		{
			comments.append( "\r\n" + aNh.getOOXML( authors.indexOf( aNh.getAuthor() ) ) );
		}
		comments.append( "\r\n</commentList>\r\n</comments>" );
		addDeferredFile( comments, "xl/comments" + (++commentsId) + ".xml" );

		shContentList.add( new String[]{
				"/xl/comments" + (commentsId) + ".xml", "comments"
		} );

	}

	/**
	 * writes the legacy drawing xml in drawing directory/vmlDrawingX.vml plus the rels for any embedded objects
	 * <br>most of the vml is saved from the original (if any); the vml used to define notes is genearated here
	 *
	 * @param zip
	 * @param sheet
	 * @param bk
	 * @return true if there is vml (legacy drawing info and/or notes) for this sheet
	 * @throws IOException
	 */
	private boolean writeLegacyDrawingObjects( Writer out, WorkSheetHandle sheet, WorkBookHandle bk ) throws IOException
	{
		// TODO: handle embeds + rels -- correct???? ****************************
		// TODO: data=1 --> is that number of shapes??  <o:idmap v:ext="edit" data=""/>
		// number of shape id's??????
		String[] embeds = null;
		StringBuffer vml = null;
		if( sheet.getMysheet().getOOXMLShapes() != null )
		{
			if( sheet.getMysheet().getOOXMLShapes().get( "vml" ) instanceof Object[] )
			{ // original vml contains embeds
				Object[] o = (Object[]) sheet.getMysheet().getOOXMLShapes().get( "vml" );
				vml = (StringBuffer) o[0];
				embeds = (String[]) o[1];
			}
			else
			{
				vml = (StringBuffer) sheet.getMysheet().getOOXMLShapes().get( "vml" );
			}
		}
		// add Notes to the vml, if any
		CommentHandle[] nh = sheet.getCommentHandles();
		if( (nh != null) && (nh.length > 0) )
		{
			if( vml == null )
			{
				vml = new StringBuffer();
				// add apparently required shapelayout element 
				vml.append( "<o:shapelayout v:ext=\"edit\">" +
						            "<o:idmap v:ext=\"edit\" data=\"1\"/>" + /*        --> data="1" --> number of elements in the vml?????*/
						            "</o:shapelayout>" );
			}
			// add shapetype which defines textbox (==202)
			vml.append( "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">" +
					            "<v:stroke joinstyle=\"miter\"/>" +
					            "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>" +
					            "</v:shapetype>" );

			for( CommentHandle aNh : nh )
			{
				boolean hidden = aNh.getIsHidden();
				int row = aNh.getRowNum();
				int col = aNh.getColNum();
				short[] bounds = aNh.getTextBoxBounds();
				int spid = aNh.getInternalNoteRec().getSPID();
				vml.append( "<v:shape id=\"_x0000_s" + spid + "\"" +  /* id of text box = id of mso */
						            " type=\"#_x0000_t202\"" +                           /* type of text box */
						            " style=\"position:absolute;" +
						            "margin-left:203.25pt;" +
						            "margin-top:37.5pt;" +
						            "width:96pt;" +
						            "height:55.5pt;" +
						            "z-index:1;" +
						            "visibility:" + ((hidden) ? "hidden" : "visible") + "\"" +  /* shown or hidden */
						            " fillcolor=\"#ffffe1\"" +
						            " o:insetmode=\"auto\">" );
				vml.append( "<v:fill color2=\"#ffffe1\"/>" +                /* general textbox characteristics */
						            " <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>" +
						            " <v:path o:connecttype=\"none\"/>" +
						            "<v:textbox style=\"mso-direction-alt:auto\">" +
						            "<div style=\"text-align:left\"/>" +
						            " </v:textbox>" );
				vml.append( "<x:ClientData ObjectType=\"Note\">" +      /* note object */
						            "<x:MoveWithCells/>" +
						            "<x:SizeWithCells/>" );
				vml.append( "<x:Anchor>" );                              /* bounds of text box */
				for( int j = 0; j < bounds.length; j++ )
				{
					vml.append( bounds[j] + ((j < (bounds.length - 1)) ? "," : "") );
				}
				vml.append( "</x:Anchor>" );

				vml.append( "<x:AutoFill>False</x:AutoFill>" +          /* row/col where note is attached to */
						            "<x:Row>" + row + "</x:Row>" +
						            "<x:Column>" + col + "</x:Column>" );
				if( !hidden )
				{
					vml.append( "<x:Visible/>" );
				}
				vml.append( "</x:ClientData>" + "</v:shape>" );
			}
		}
		if( (vml != null) && (vml.length() > 0) )
		{
			// start with ns info
			vml.insert( 0, "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"" +
					" xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
					" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">" );
			vml.append( "</xml>" );
			addDeferredFile( vml, drawingDir + "/vmlDrawing" + (++vmlId) + ".vml" );

			shContentList.add( new String[]{
					"/" + drawingDir + "/vmlDrawing" + (vmlId) + ".vml", "vmldrawing"
			} );

			if( embeds != null )
			{
				ArrayList vmlContentList = new ArrayList();
				for( String embed : embeds )
				{
					String pp = embed.trim();   // file on disk or saved filename
					String typ = pp.substring( 0, pp.indexOf( "/" ) );
					pp = pp.substring( pp.indexOf( "/" ) + 1 );
					int z = pp.lastIndexOf( "/" ) + 1;
					String pth = pp.substring( 0, z );
					pp = pp.substring( z );
					String ff = getExOOXMLFileName( pp ); // desired or outut filename
					deferredFiles.put( pth + ff, externalDir + pp );
					vmlContentList.add( new String[]{ "/" + pth + ff, typ } );
				}
				writeRels( vmlContentList, drawingDir + "/_rels/vmlDrawing" + vmlId + ".vml.rels" );
				// NOTE: vmlfiles do not get listed in main [Content_Types].xml
			}
			return true;
		}
		return false;
	}

}
