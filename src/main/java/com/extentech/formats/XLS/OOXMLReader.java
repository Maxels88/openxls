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
import com.extentech.ExtenXLS.DateConverter;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.OOXML.Border;
import com.extentech.formats.OOXML.Dxf;
import com.extentech.formats.OOXML.Fill;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.OOXML.PivotCacheDefinition;
import com.extentech.formats.OOXML.PivotTableDefinition;
import com.extentech.formats.OOXML.Theme;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import org.xmlpull.v1.XmlPullParserFactory;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Stack;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * Breaking out functionality for reading out of OOXMLAdapter
 */
public class OOXMLReader extends OOXMLAdapter implements OOXMLConstants
{
	private static final Logger log = LoggerFactory.getLogger( OOXMLReader.class );
//	private int defaultXf= 0;	// usual for OOXML files; however, those which are converted from XLS may have default xf as 15 

	/*****************************************************************************************************************************************/
	/**
	 * Parsing/Reading OOXML Input Section /
	 *****************************************************************************************************************************************/

	/**
	 * OOXML parseNBind - reads in an OOXML (Excel 7) workbook
	 *
	 * @param bk    WorkBookHandle - workbook to input
	 * @param fName OOXML filename (must be a ZIP file in OPC format)
	 * @throws XmlPullParserException
	 * @throws IOException
	 * @throws CellNotFoundException
	 */
	public void parseNBind( WorkBookHandle bk, String fName ) throws XmlPullParserException, IOException, CellNotFoundException
	{
		ZipFile zip = new ZipFile( fName );

		inputEncoding = System.getProperty( "file.encoding" );
		if( inputEncoding == null )
		{
			inputEncoding = "UTF-8";
		}
		// KSC: replaced with above       isUTFEncoding= (System.getProperty("file.encoding").startsWith("UTF"));
		// clear out state vars
		bk.getWorkBook().setIsExcel2007( true );
		int origcalcmode = bk.getWorkBook().getCalcMode();
		bk.getWorkBook().setCalcMode( WorkBook.CALCULATE_EXPLICIT );   // don't calculate formulas on input
		// TODO: read in format type (doc, macro-enabled doc, template, macro-enabled tempate) from CONTENT_LIST ???       
		ZipEntry rels = getEntry( zip, "_rels/.rels" );
			log.debug( "parseNBind about to call parseRels on: " + rels.toString() );
		mainContentList = parseRels( wrapInputStream( zip.getInputStream( rels ) ) );
		bk.getWorkBook().getFactory().setFileName( fName );
		bk.setDupeStringMode( WorkBookHandle.SHAREDUPES );

       /* KSC: remove Xf recs first  -- NOTE has some issues for XLS->XLSX -- must fix !! */
		bk.getWorkBook().removeXfRecs();
		bk.getWorkBook().setDefaultIxfe( 0 );

		externalDir = getTempDir( bk.getWorkBook().getFactory().getFileName() );
		ArrayList formulas = new ArrayList();     // set in parseSheetXML, must process formulas after all sheets/cells have been added 
		ArrayList hyperlinks = new ArrayList();   // set in parseSheetXML, links with hyperlink target info in sheetX.xml.rels            ""
		HashMap inlineStrs = new HashMap();               // set in parseSheetXML, stores inlinestring text with addresses for entry after all sheets have been added
		parseBookLevelElements( bk, null, zip, mainContentList, "", formulas, hyperlinks, inlineStrs, null, null );
		zip.close();
		// if hasn't been streamed, delete temp dir
		if( !bk.getWorkBook().getFactory().getFileName().endsWith( ".tmp" ) )
		{
			deleteDir( new File( externalDir ) );    // don't save temp files (pass-through's) -- can reinstate when needed
		}
		bk.getWorkBook().setCalcMode( origcalcmode );  // reset
	}

	protected static boolean parsePivotTables = true;        // KSC: TESTING -- only make true in testing

	/**
	 * parses OOXML content files given a content list cl from zip file zip
	 * recurses if content file has it's own content
	 * *************************************
	 * NOTE: certain elements we do not as of yet process; we "pass-through" or store such elements along with any embedded objects associated with them
	 * for example, activeX objects, vbaProject.bin, etc.
	 * *************************************
	 *
	 * @param bk        WorkBookHandle
	 * @param sheet     WorkSheetHandle (set if recursing)
	 * @param zip       currently open ZipOutputStream
	 * @param cl        ArrayList of Contents (type, filename, rId) to parse
	 * @param parentDir Parent Directory for relative paths in content lists
	 * @param formulas, hyperlinks, inlineStrs -- ArrayLists/Hashmaps stores sheet-specific info for later entry
	 * @throws CellNotFoundException
	 * @throws XmlPullParserException
	 */
	protected void parseBookLevelElements( WorkBookHandle bk,
	                                       WorkSheetHandle sheet,
	                                       ZipFile zip,
	                                       ArrayList cl,
	                                       String parentDir,
	                                       ArrayList formulas,
	                                       ArrayList hyperlinks,
	                                       HashMap inlineStrs,
	                                       HashMap<String, String> pivotCaches,
	                                       HashMap<String, WorkSheetHandle> pivotTables ) throws
	                                                                                      XmlPullParserException,
	                                                                                      CellNotFoundException
	{
		String p;    // target path
		ZipEntry target;
		ArrayList sst = new ArrayList(); // set in parseSSTXML, used in parsing sheet XML

		try
		{
			// parse content list for <elementName, target's path, rId>
			for( Object aCl : cl )
			{
				String[] c = (String[]) aCl;

					log.debug( "OOXMLReader.parse: " + c[0] + ":" + c[1] + ":" + c[2] );

				p = StringTool.getPath( c[1] );
				p = parsePathForZip( p, parentDir );

				String ooxmlElement = c[0];
				if( !ooxmlElement.equals( "hyperlink" ) )  // if it's a hyperlink reference, don't strip path info :)
				{
					c[1] = StringTool.stripPath( c[1] );
				}
				String f = c[1];    // root filename
				String rId = c[2];

				if( ooxmlElement.equals( "styles" ) )
				{
					target = getEntry( zip, p + f );
					parseStylesXML( bk, wrapInputStream( zip.getInputStream( target ) ) );
				}
				else if( ooxmlElement.equals( "sst" ) )
				{
					target = getEntry( zip, p + f );
					sst = Sst.parseOOXML( bk, wrapInputStream( zip.getInputStream( target ) ) );
				}
				else if( ooxmlElement.equals( "sheet" ) )
				{
					// sheet.xml
					target = getEntry( zip, p + f );

					try
					{
						int sheetnum = 1;
						try
						{
							String s = rId.substring( 3 );        // in form of "rIdXX" where XX is the sheet number
							sheetnum = Integer.valueOf( s ) - 1;  // embed attribute, specifies rId, important in OOXML
						}
						catch( Exception e )
						{
							log.warn( "OOXMLAdapter couldn't get sheet number from rid:" + rId , e);
						}
						sheet = bk.getWorkSheet( sheetnum );
						// since we're adding a lot of cells, put sheet in fast add mode    // put statement here AFTER sheet is set :)
						sheet.setFastCellAdds( true );

						sheet.getMysheet().parseOOXML( bk,
						                               sheet,
						                               wrapInputStream( zip.getInputStream( target ) ),
						                               sst,
						                               formulas,
						                               hyperlinks,
						                               inlineStrs );

						// sheet.xml.rels
						target = getEntry( zip, p + "_rels/" + f.substring( f.lastIndexOf( "/" ) + 1 ) + ".rels" );
						if( target != null )
						{
							try
							{
								HashMap pts = new HashMap<String, WorkSheetHandle>();
								sheet.getMysheet().parseSheetElements( bk,
								                                       zip,
								                                       parseRels( wrapInputStream( wrapInputStream( zip.getInputStream(
										                                       target ) ) ) ),
								                                       p,
								                                       externalDir,
								                                       formulas,
								                                       hyperlinks,
								                                       inlineStrs,
								                                       pts );
								if( pts.size() > 0 )
								{
									pivotTables.putAll( pts );
								}
							}
							catch( Exception e )
							{
								log.warn( "OOXMLAdapter.parse problem parsing rels in: " + bk.toString() + " " + e.toString(), e );
							}
						}
						// reset fast add mode
						sheet.setFastCellAdds( false );   // 20090713 KSC: moved from below
					}
					catch( WorkSheetNotFoundException we )
					{
						log.error( "OOXMLAdapter.parse: " + we.toString() );
					}
				}
				else if( ooxmlElement.equals( "document" ) )
				{ // main workbook document 
					// workbook.xml
					target = getEntry( zip, p + f );

						log.debug( "About to parseWBOOXML:" + bk.toString() );
					pivotCaches = new HashMap<>();
					parsewbOOXML( zip,
					              bk,
					              wrapInputStream( zip.getInputStream( target ) ),
					              p,
					              pivotCaches );    // seets, defined names, pivotcachedefinition ...  

					// now parse wb content - sheets and their sub-contents (charts, images, oleobjects...)
					pivotTables = new HashMap<>();
					parseBookLevelElements( bk, sheet, zip, wbContentList, p, formulas, hyperlinks, inlineStrs, pivotCaches, pivotTables );

					// after all sheet data has been added, now can add inline strings, if any
					if( inlineStrs != null )
					{
						addInlineStrings( bk, inlineStrs );
					}
					// after all sheet data has been added, now can add formulas
					addFormulas( bk, formulas );
					// after all sheet data and formulas, NOW can add pivot Tables
					addPivotTables( bk, zip, pivotTables );
				}
				else if( parsePivotTables && ooxmlElement.equals( "pivotCacheDefinition" ) )
				{    // workbook-parent + pivotTable-parent
					//pivotCaches.add(new int[] {cid, id});                    	                	 
					target = getEntry( zip, p + f );
					PivotCacheDefinition.parseOOXML( bk, pivotCaches.get( rId ), wrapInputStream( zip.getInputStream( target ) ) );
					target = getEntry( zip, p + "_rels/" + f.substring( f.lastIndexOf( "/" ) + 1 ) + ".rels" );
					if( target != null )
					{    // pivotCacheRecords ... 
						try
						{
							parseBookLevelElements( bk,
							                        sheet,
							                        zip,
							                        parseRels( wrapInputStream( wrapInputStream( zip.getInputStream( target ) ) ) ),
							                        p,
							                        formulas,
							                        hyperlinks,
							                        inlineStrs,
							                        pivotCaches,
							                        pivotTables );
						}
						catch( Exception e )
						{
							log.warn( "OOXMLAdapter.parse problem parsing rels in: " + bk.toString() + " " + e.toString(), e );
						}
					}
				}
				else if( parsePivotTables && ooxmlElement.equals( "pivotCacheRecords" ) )
				{    // pivotcacheDefinition-parent                 
				}
				else if( ooxmlElement.equals( "theme" ) || (ooxmlElement.equals( "themeOverride" )) )
				{ // read in theme colors
					target = getEntry( zip, p + f );
					if( target != null )
					{
						if( bk.getWorkBook().getTheme() == null )
						{
							bk.getWorkBook().setTheme( Theme.parseThemeOOXML( bk, wrapInputStream( zip.getInputStream( target ) ) ) );
						}
						else
						{
							bk.getWorkBook().getTheme().parseOOXML( bk,
							                                        wrapInputStream( zip.getInputStream( target ) ) );    // theme overrides
						}
					}
					handlePassThroughs( zip, bk, p, externalDir, c );
					// Below are elements we do not as of yet handle
				}
				else if( ooxmlElement.equals( "props" ) || ooxmlElement.equals( "exprops" ) || ooxmlElement.equals( "custprops" ) || ooxmlElement
						.equals( "connections" ) || ooxmlElement.equals( "calc" ) || ooxmlElement.equals( "vba" ) || ooxmlElement.equals(
						"externalLink" ) )
				{

					handlePassThroughs( zip, bk, p, externalDir, c );   // pass-through this file and any embedded objects as well
				}
				else
				{    // unknown type
					log.warn( "OOXMLReader.parse:  XLSX Option Not yet Implemented " + ooxmlElement );
				}
			}
		}
		catch( IOException e )
		{
			log.error( "OOXMLReader.parse failed: " + e.toString() );
		}
	}

	/**
	 * pass-through current OOXML element/file - i.e. save file to external directory on disk
	 * because it cannot be processed into our normal BIFF8 machinery
	 * also, if current OOXML element has an associated .rels file containing links to other files known as "embeds",
	 * store embeds on disk and link information to for later retrieval
	 * <p/>
	 * <br>Possible pass-through types:
	 * <br>props
	 * <br>exprops
	 * <br>custprops
	 * <br>connections
	 * <br>calc
	 * <br>vba
	 * <br>externalLink
	 *
	 * @param c String[] {type, filename, rid}
	 */
	protected static void handlePassThroughs( ZipFile zip, WorkBookHandle bk, String parentDir, String externalDir, String[] c ) throws
	                                                                                                                             IOException
	{
		passThrough( zip, parentDir + c[1], externalDir + c[1] ); // save the original target file for later re-packaging
		ZipEntry target = getEntry( zip,
		                            parentDir + "_rels/" + c[1].substring( c[1].lastIndexOf( "/" ) + 1 ) + ".rels" ); // is there an associated .rels file??
		if( target == null )  // no .rels, just link to original OOXML element/file
		{
			bk.getWorkBook().addOOXMLObject( new String[]{ c[0], parentDir, externalDir + c[1] } );
		}
		else              // handle embedded objects in \book-level objects (theme embeds, externalLinks
		{
			bk.getWorkBook().addOOXMLObject( new String[]{
					c[0],
					parentDir,
					externalDir + c[1],
					c[2],
					null,
					Arrays.asList( storeEmbeds( zip, target, parentDir, externalDir ) )
					      .toString()/* 1.6 only Arrays.toString(storeEmbeds(zip, target, p))*/
			} );
		}
	}

	/**
	 * pass-through current sheet-level OOXML element/file - i.e. save file to external directory on disk
	 * because it cannot be processed into our normal BIFF8 machinery
	 * also, if current OOXML element has an associated .rels file containing links to other files known as "embeds",
	 * store embeds on disk and link information to for later retrieval
	 *
	 * @param c String[] {type, filename, rid}
	 */
	protected static void handleSheetPassThroughs( ZipFile zip,
	                                               WorkBookHandle bk,
	                                               Boundsheet sht,
	                                               String parentDir,
	                                               String externalDir,
	                                               String[] c,
	                                               String attrs ) throws IOException
	{
		passThrough( zip, parentDir + c[1], externalDir + c[1] ); // save the original target file for later re-packaging
		ZipEntry target = getEntry( zip,
		                            parentDir + "_rels/" + c[1].substring( c[1].lastIndexOf( "/" ) + 1 ) + ".rels" ); // is there an associated .rels file??
		if( target == null ) // no .rels, just link to original OOXML element/file
		{
			sht.addOOXMLObject( new String[]{ c[0], parentDir, externalDir + c[1], c[2], attrs } );
		}
		else                // handle embedded objects in sheet-level objects (activeX binaries ....)
		{
			sht.addOOXMLObject( new String[]{
					c[0],
					parentDir,
					externalDir + c[1],
					c[2],
					attrs,
					Arrays.asList( storeEmbeds( zip, target, parentDir, externalDir ) )
					      .toString() /* 1.6 only Arrays.toString(storeEmbeds(zip, target, p))*/
			} );
		}
	}

	/**
	 * handle OOXML files that we do not process at this time.
	 * <br>Writes the file in question from zip file fin to file directory file fout
	 *
	 * @param zip
	 * @param fin
	 * @param fout
	 * @throws IOException
	 */
	protected static void passThrough( ZipFile zip, String fin, String fout ) throws IOException
	{
		try
		{
			java.io.File outfile = new java.io.File( fout );
			// clean it up
			outfile.deleteOnExit();

			File dirs = outfile.getParentFile();
			if( (dirs != null) && !dirs.exists() )
			{
				dirs.mkdirs();
				dirs.deleteOnExit();
			}
			BufferedOutputStream fos = new BufferedOutputStream( new FileOutputStream( outfile ) );
			InputStream fis = OOXMLReader.wrapInputStream( zip.getInputStream( OOXMLReader.getEntry( zip, fin ) ) );
			int i = fis.read();
			while( i != -1 )
			{
				fos.write( i );
				i = fis.read();
			}
			fos.flush();
			fos.close();
			dirs.delete();
		}
		catch( Exception e )
		{
			; // OK for external links for FNFE
		}
	}

	/**
	 * given workbook.xml inputstream, parse OOXML into array list of content (only sheets and names at this point; eventually will handle docProps ...)
	 *
	 * @param bk         WorkBookHandle
	 * @param ii         inputStream
	 * @param namedRange ArrayList to hold named ranges (must be added after all sheet data)
	 * @return
	 */
	ArrayList parsewbOOXML( ZipFile zip, WorkBookHandle bk, InputStream ii, String p, HashMap<String, String> pivotCaches )
	{
		ArrayList namedRanges = new ArrayList();  //must save and parse after all sheets have been added
		ArrayList contentList = new ArrayList();
		ArrayList sheets = new ArrayList();

		// set the default date format
		bk.getWorkBook().dateFormat = DateConverter.DateFormat.OOXML_1900;

		try
		{
			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();

			xpp.setInput( ii, null ); // using XML 1.0 specification 
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "sheet" ) )
					{
						String name = ""; // name, sheetId, rId, state=hidden
						int id = 0;      // sheetId is used??
						int rId = 0;
						String hidden = "";
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( nm.equalsIgnoreCase( "name" ) )
							{
								name = v;
							}
							else if( nm.equalsIgnoreCase( "SheetId" ) )
							{
								id = Integer.valueOf( v ) - 1;
							}
							else if( nm.equalsIgnoreCase( "id" ) )    // rId
							{
								rId = Integer.valueOf( v.substring( 3 ) ) - 1;
							}
							else if( nm.equals( "state" ) )
							{
								hidden = v;
							}
						}
						// sheets may very well NOT be in order so must create after all sheets have been accounted for bk.createWorkSheet(name, id);
						for( int i = sheets.size(); i < rId; i++ )
						{
							sheets.add( "" );
						}
						sheets.add( rId, new String[]{ name, hidden } );
						contentList.add( new String[]{ "sheet", name } );
					}
					else if( tnm.equals( "workbookPr" ) )
					{ // TODO: get other attributes such as date1904
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String attrName = xpp.getAttributeName( i );
							String attrValue = xpp.getAttributeValue( i );
							if( attrName.equalsIgnoreCase( "codeName" ) )
							{
								bk.getWorkBook().setCodename( attrValue );
							}
							else if( attrName.equalsIgnoreCase( "dateCompatibility" ) && attrValue.equals( "1" ) )
							{
								bk.getWorkBook().dateFormat = DateConverter.DateFormat.LEGACY_1900;
							}
							else if( attrName.equalsIgnoreCase( "date1904" ) && attrValue.equals( "1" ) )
							{
								bk.getWorkBook().dateFormat = DateConverter.DateFormat.LEGACY_1904;
							}
						}
					}
					else if( tnm.equals( "workbookView" ) )
					{   // TODO: handle other workbookview attributes
						String n = "";
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							n = xpp.getAttributeName( i );
							if( n.equalsIgnoreCase( "firstSheet" ) )
							{
								bk.getWorkBook().setFirstSheet( Integer.valueOf( xpp.getAttributeValue( i ) ) );
							}
							//else if (n.equalsIgnoreCase("activeTab"))
							//bk.getWorkBook().setActiveTab(Integer.valueOf(xpp.getAttributeValue(i)).intValue());
							else if( n.equals( "showSheetTabs" ) )
							{
								boolean b = (!xpp.getAttributeValue( i ).equals( "0" ));
								bk.getWorkBook().setShowSheetTabs( b );
							}
						}               
/*                   } else if (tnm.equals("externalReference")) {  // TODO: HANDLE! nothing really to do here as it just denotes an rId, handled upon encountering externalLink */
					}
					else if( tnm.equals( "definedName" ) )
					{    // have to process after sheets have been added, so save info
						// attributes:  (not used by us) comment, customMenu, description, help, shortcutKey, statusBar, vbProcedure, workbookParameter, publishToServer
						// (to deal with later) function, functionGroupId -- Specifies a boolean value that indicates that the defined name refers to a user-defined function/add-in ...
						// xlm (External Function)
						// (should deal with) hidden, localSheetId 
						String nm = "";
						String id = "";
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "name" ) )        // can be built-in:  _xlnm.Print_Area, _xlnm.Print_Titles, _xlnm.Criteria, _xlnm ._FilterDatabase, _xlnm .Extract, _xlnm .Consolidate_Area, _xlnm .Database, _xlnm .Sheet_Title 
							{
								nm = xpp.getAttributeValue( i );
							}
							else if( n.equals( "localSheetId" ) )        // Specifies the sheet index in this workbook where data from an external reference is displayed.
							{
								id = xpp.getAttributeValue( i );
							}
						}
						String name = getNextText( xpp );    // value can be a function, a
						// has an external wb specification, remove as messes up parsing
						if( !id.equals( "" ) && name.startsWith( "[" ) )
						{ // remove external denotation [#]sheet!range
							int n = 0;
							while( (n = name.indexOf( "[" )) > -1 )
							{
								name = name.substring( 0, n ) + name.substring( n + name.substring( n ).indexOf( "]" ) + 1 );
							}
						}
						namedRanges.add( new String[]{ nm, id, name } );
					}
					else if( tnm.equals( "pivotCache" ) )
					{
						String cid = "";
						String rid = "";
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							if( n.equals( "cacheId" ) )
							{
								cid = xpp.getAttributeValue( i );
							}
							else if( n.equals( "id" ) )
							{
								rid = xpp.getAttributeValue( i );
							}
						}
						pivotCaches.put( rid, cid );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "OOXMLAdapter.parsewbOOXML failed: " + e.toString() );
		}
		for( int i = 0; i < sheets.size(); i++ )
		{
			Object obx = sheets.get( i );

			String[] sh = null;
			if( obx instanceof String )
			{
				sh = new String[1];
				sh[0] = obx.toString();
			}
			else
			{
				sh = (String[]) obx;
			}
			String name = sh[0];
			if( (name != null) && !name.equals( "" ) )
			{
				bk.createWorkSheet( name, i );
				try
				{
					bk.getWorkBook().setDefaultIxfe( 0 );    // OOXML default xf= 0, 2003-vers, default= 15 see below
					if( sh[1].equals( "hidden" ) )
					{
						bk.getWorkSheet( i ).setHidden( true );
					}
					else if( sh[1].equals( "veryHidden" ) )
					{
						bk.getWorkSheet( i ).setVeryHidden( true );
					}
				}
				catch( Exception e )
				{
					// shouldn't!
				}
			}
		}
		// workbook.xml.rels
		ZipEntry target = getEntry( zip, p + "_rels/workbook.xml.rels" );
		try
		{
			wbContentList = parseRels( wrapInputStream( wrapInputStream( zip.getInputStream( target ) ) ) );
		}
		catch( IOException e )
		{
			log.warn( "OOXMLReader.parseWbOOXML: " + e.toString(), e );
		}
		// for workbook contents, MUST PROCESS themes before styles, sst and styles, etc. before SHEETS
		reorderWbContentList( wbContentList );

		// add all named ranges
		addNames( bk, namedRanges );

		return contentList;
	}

	/**
	 * get correct path for zip access based on path p and parent directory parentDir
	 *
	 * @param p
	 * @param parentDir
	 */
	protected static String parsePathForZip( String p, String parentDir )
	{
		if( ((p.indexOf( "/" ) != 0) || (p.indexOf( "\\" ) == 0)) )
		{
			while( p.indexOf( ".." ) == 0 )
			{
				p = p.substring( 3 );
				if( !parentDir.equals( "" ) && ((parentDir.charAt( parentDir.length() - 1 ) == '/') || (parentDir.charAt( parentDir.length() - 1 ) == '\\')) )
				{
					parentDir = parentDir.substring( 0, parentDir.length() - 2 );
				}
				int z = parentDir.lastIndexOf( "/" );
				if( z == -1 )
				{
					z = parentDir.lastIndexOf( "\\" );
				}
				parentDir = parentDir.substring( 0, z + 1 );
			}

			p = parentDir + p;

			//if(DEBUG)
			//  Logger.logInfo("parsePathForZip:"+p);

		}
		else if( !p.equals( "" ) )
		{
			p = p.substring( 1 );
		}

		return p;
	}

	/**
	 * retrieves the entire element at the current position in the xpp pullparser,
	 * as a string, and advances the pullparser position to the next element
	 *
	 * @param xpp
	 * @return
	 */
	protected static String getCurrentElement( XmlPullParser xpp )
	{
		StringBuffer el = new StringBuffer();
		try
		{
			int eventType = xpp.getEventType();
			String elname = xpp.getName();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( (eventType == XmlPullParser.START_TAG) || (eventType == XmlPullParser.TEXT) )
				{
					el.append( xpp.getText() );
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String t = xpp.getText();
					if( t.indexOf( "</" ) == 0 )
					{
						el.append( t );
					}
					if( xpp.getName().equals( elname ) )
					{
						break;
					}

				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "OOXMLAdapter.getCurrentElement: " + e.toString() );
		}
		return el.toString();
	}

	/**
	 * XmlPullParser positioned on <is> child of the <c> (cell) element in sheetXXX.xml
	 *
	 * @param xpp XmlPullParser
	 * @return String inline text
	 * @throws XmlPullParserException
	 * @throws IOException
	 */
	protected static String getInlineString( XmlPullParser xpp ) throws XmlPullParserException, IOException
	{
		int eventType = xpp.next();
		String ret = "";
		while( (eventType != XmlPullParser.END_DOCUMENT) &&
				(eventType != XmlPullParser.END_TAG) &&
				(eventType != XmlPullParser.TEXT) )
		{
			eventType = xpp.next();
		}
		if( eventType == XmlPullParser.TEXT )
		{
			ret = xpp.getText();
		}

		try
		{
			return new String( ret.getBytes(), inputEncoding );
		}
		catch( Exception e )
		{
		}  // inputEncoding can be null
		return ret;
	}

	// records used in parsing styles.xml
	ArrayList borders;
	ArrayList<Integer> fontmap;
	ArrayList<Fill> fills;
	ArrayList<Dxf> dxfs;
	HashMap fmts;
	int nXfs;

	/**
	 * given Styles.xml OOXML input stream, parse and input into workbook
	 *
	 * @param bk WorkBookHandle
	 * @param ii InputStream
	 */
	void parseStylesXML( WorkBookHandle bk, InputStream ii )
	{
		try
		{
			borders = new ArrayList();
			fontmap = new ArrayList<>();
			fills = new ArrayList<>();
			dxfs = new ArrayList<>();
			fmts = new HashMap();
			nXfs = 0;                                  // position in xfrecs array is vital as cells will reference the styleId/xfId
			int indexedColor = 0;                          // index into COLOR_TABLE 

			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();

			xpp.setInput( ii, null ); // using XML 1.0 specification 
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "font" ) )
					{
						Font f = Font.parseOOXML( xpp, bk );
						int idx = FormatHandle.addFont( f, bk );
						fontmap.add( idx );
					}
					else if( tnm.equals( "dxfs" ) )
					{   // differential formatting (conditional formatting) style
					}
					else if( tnm.equals( "dxf" ) )
					{ // incremental style info -- for conditional save
						Dxf d = (Dxf) Dxf.parseOOXML( xpp, bk ).cloneElement();
						dxfs.add( d );
					}
					else if( tnm.equals( "fill" ) )
					{
						Fill f = (Fill) Fill.parseOOXML( xpp, false, bk );
						fills.add( f );    //new Object[] { Integer.valueOf(fp), fgColor, bgColor});                       
					}
					else if( tnm.equals( "numFmt" ) )
					{
						int fmtId = 0;
						int newFmtId = 0;
						String xmlFormatPattern = "";
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String nm = xpp.getAttributeName( i );
							if( nm.equals( "numFmtId" ) )
							{
								fmtId = Integer.valueOf( xpp.getAttributeValue( i ) );
							}
							else if( nm.equals( "formatCode" ) )
							{
								xmlFormatPattern = xpp.getAttributeValue( i );
								xmlFormatPattern = Xf.unescapeFormatPattern( xmlFormatPattern );
							}
						}
						newFmtId = Xf.addFormatPattern( bk.getWorkBook(), xmlFormatPattern );
						fmts.put( fmtId, newFmtId );  // map our format id with original
					}
					else if( tnm.equals( "border" ) )
					{ // TODO: use Border element to parse
						Border b = (Border) Border.parseOOXML( xpp, bk ).cloneElement();
						borders.add( b );
					}
					else if( tnm.equals( "cellXfs" ) )
					{
						while( eventType != XmlPullParser.END_DOCUMENT )
						{
							if( eventType == XmlPullParser.START_TAG )
							{
								parseCellXf( xpp, bk );
							}
							else if( (eventType == XmlPullParser.END_TAG) && xpp.getName().equals( "cellXfs" ) )
							{
								break;
							}
							eventType = xpp.next();
						}
					}
					else if( tnm.equals( "rgbColor" ) )
					{
						// save custom indexed colors
						String clr = "#" + xpp.getAttributeValue( 0 ).substring( 2 );
						//System.out.println(clr);
						// usually the same as COLORTABLE but sometimes different too :)
						try
						{
							bk.getWorkBook().getColorTable()[indexedColor++] = FormatHandle.HexStringToColor( clr );
						}
						catch( ArrayIndexOutOfBoundsException e )
						{
							// happens?
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "worksheet" ) ) // we're done!
					{
						break;
					}
				}
				else if( eventType == XmlPullParser.TEXT )
				{
				}
				if( eventType != XmlPullParser.END_DOCUMENT )
				{
					eventType = xpp.next();
				}
			}
			if( dxfs.size() > 0 )
			{
				bk.getWorkBook().setDxfs( dxfs );
			}

		}
		catch( Exception e )
		{
			log.error( "OOXMLReader.parseStylesXML: " + e.toString() );
		}
	}

	/**
	 * Parse the cellXF's section of the styles.xml file
	 *
	 * @param xpp
	 * @param bk
	 */
	private void parseCellXf( XmlPullParser xpp, WorkBookHandle bk )
	{
		String tnm = xpp.getName();
		if( tnm.equals( "xf" ) )
		{
			int f = 0;
			int fmtId = -1;
			int fillId = -1;
			int borderId = -1;
			for( int i = 0; i < xpp.getAttributeCount(); i++ )
			{
				String nm = xpp.getAttributeName( i );
				// all id's are 0-based
				if( nm.equals( "fontId" ) )
				{
					f = Integer.valueOf( xpp.getAttributeValue( i ) );
				}
				else if( nm.equals( "numFmtId" ) )
				{
					fmtId = Integer.valueOf( xpp.getAttributeValue( i ) );
				}
				else if( nm.equals( "fillId" ) )
				{
					fillId = Integer.valueOf( xpp.getAttributeValue( i ) );
				}
				else if( nm.equals( "borderId" ) )
				{
					borderId = Integer.valueOf( xpp.getAttributeValue( i ) );
				}
			}
			f = fontmap.get( f );  // FONT
			Xf xf = null;
			if( nXfs < bk.getWorkBook().getXfrecs().size() )    // either alter existing default xf or create new xf
			{
				xf = (Xf) bk.getWorkBook().getXfrecs().get( nXfs );
			}
			if( xf == null )// if it doesn't exist, create new otherwise overwrite orig
			{
				xf = Xf.updateXf( null, f, bk.getWorkBook() );
			}
			else
			{
				xf.setFont( f );
				xf.setFormat( (short) 0 );
			}
			if( fmtId > 0 )
			{ // NUMBER FORMAT 0 is default
				if( fmts.get( (Integer.valueOf( fmtId )) ) != null )  // map it
				{
					fmtId = (Integer) (fmts.get( (Integer.valueOf( fmtId )) ));
				}
				xf.setFormat( (short) fmtId );
			}
			if( borderId > -1 )
			{     // BORDER
				Border b = (Border) borders.get( borderId );
				xf.setAllBorderLineStyles( b.getBorderStyles() );    //bs);
				xf.setAllBorderColors( b.getBorderColorInts() );
			}
			if( fillId > 0 )
			{    // FILL 0 is default
				xf.setFill( fills.get( fillId ) );
			}
			// is xf 15 the default? (will happen if converted from xls) ******* very important to avoid unnecessary blank creation *******
			// see TestCorruption.TestStackOverflow
			if( (nXfs == 15) && xf.toString().equals( bk.getWorkBook().getXf( 0 ).toString() ) )
			{
				bk.getWorkBook().setDefaultIxfe( 15 );
			}
			nXfs++;

		}
		else if( tnm.equals( "protection" ) )
		{
			Xf xf = bk.getWorkBook().getXf( nXfs - 1 );
			for( int j = 0; j < xpp.getAttributeCount(); j++ )
			{
				String n = xpp.getAttributeName( j );
				String v = xpp.getAttributeValue( j );
				if( n.equals( "hidden" ) )
				{
					xf.setFormulaHidden( v.equals( "1" ) );
				}
				else if( n.equals( "locked" ) )
				{
					xf.setLocked( v.equals( "1" ) );
				}
			}
		}
		else if( tnm.equals( "alignment" ) )
		{
			Xf xf = bk.getWorkBook().getXf( nXfs - 1 );
			for( int j = 0; j < xpp.getAttributeCount(); j++ )
			{
				String n = xpp.getAttributeName( j );
				String v = xpp.getAttributeValue( j );
				if( n.equals( "horizontal" ) )
				{
					int ha = sLookup( v, horizontalAlignment );
					xf.setHorizontalAlignment( ha );
				}
				else if( n.equals( "vertical" ) )
				{
					int va = sLookup( v, verticalAlignment );
					xf.setVerticalAlignment( va );
				}
				else if( n.equals( "indent" ) )
				{
					xf.setIndent( Integer.valueOf( v ) );
				}
				else if( n.equals( "wrapText" ) )
				{
					xf.setWrapText( true );
				}
				else if( n.equals( "textRotation" ) )
				{
					xf.setRotation( Integer.valueOf( v ) );
				}
				else if( n.equals( "shrinkToFit" ) )
				{
					xf.setShrinkToFit( true );
				}
				else if( n.equals( "readingOrder" ) )
				{
					xf.setRightToLeftReadingOrder( Integer.valueOf( v ) );
				}
			}
		}

	}

	/**
	 * look up string index in string array
	 *
	 * @param s
	 * @param sarr String[]
	 * @return int index into sarr
	 */
	private static int sLookup( String s, String[] sarr )
	{
		if( (sarr != null) && (s != null) )
		{
			for( int i = 0; i < sarr.length; i++ )
			{
				if( s.equals( sarr[i] ) )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * intercept Sheet adds and hand off to parse event listener as needed
	 */
	static CellHandle sheetAdd( WorkSheetHandle sheet, Object val, int r, int c, int fmtid )
	{
		return sheet.add( val, r, c, fmtid );
	}

	/**
	 * take a passthrough element such as vmldrawing or theme which contains embedded objects (images), retrieve and store
	 * for later re-writing to zip
	 *
	 * @param zip    open ZipFile
	 * @param target ZipEntry pointing to .rels
	 * @param p      path
	 * @return String[] array of embeds
	 */
	protected static String[] storeEmbeds( ZipFile zip, ZipEntry target, String p, String externalDir ) throws IOException
	{
		//if(DEBUG) Logger.logInfo("storeEmbeds about to call parseRels on: " + target.toString());

		ArrayList embeds = parseRels( wrapInputStream( wrapInputStream( zip.getInputStream( target ) ) ) ); // obtain a list of image file references for use in later parsing
		Collections.sort( embeds, new Comparator()
		{
			@Override
			public int compare( Object o1, Object o2 )
			{
				Integer a = Integer.valueOf( ((String[]) o1)[2].substring( 3 ) );
				Integer b = Integer.valueOf( ((String[]) o2)[2].substring( 3 ) );
				return a.compareTo( b );
			}
		} );
		String[] strEmbeds = new String[embeds.size()];
		for( int j = 0; j < embeds.size(); j++ )
		{
			String[] v = (String[]) embeds.get( j );
			String path = StringTool.getPath( v[1] );
			path = parsePathForZip( path, p );
			v[1] = StringTool.stripPath( v[1] );
			if( !v[0].equalsIgnoreCase( "externalLinkPath" ) )  // it's OK for externally referenced book not to be present
			{
				try
				{
					passThrough( zip, path + v[1], externalDir + v[1] );    // save the original target file for later re-packaging
				}
				catch( NullPointerException e )
				{
					//   if (!v[0].equalsIgnoreCase("externalLinkPath"))  // it's OK for externally referenced book not to be present
					throw new NullPointerException();
					//}
				}
			}
			strEmbeds[j] = v[0] + "/" + path + v[1];
		}
		return strEmbeds;
	}

	/**
	 * given a list of all named ranges in the workbook, add all
	 *
	 * @param bk
	 * @param namedRanges
	 */
	static void addNames( WorkBookHandle bk, ArrayList namedRanges )
	{
		// now input named ranges before processing sheet data 
		for( Object namedRange : namedRanges )
		{
			String[] s = (String[]) namedRange;
			if( !(s[0].equals( "" ) && s[2].equals( "" )) )
			{
				try
				{
					if( s[0].indexOf( "_xlnm" ) == 0 )
					{    // it's a built-in
						String sh = s[2].substring( 0, s[2].indexOf( "!" ) );
//                      String[] addresses= StringTool.splitString(s[2], ",");
//                      for (int k= 0; k < addresses.length; k++) {
						if( s[0].equals( "_xlnm.Print_Area" ) )
						{
							try
							{
								bk.getWorkSheet( sh ).getMysheet().setPrintArea( s[2] );//addresses[k]);
							}
							catch( OutOfMemoryError e )
							{
								log.error( "OOXMLAdapter.parse OOME setting PrintArea", e );
								throw e;
							}
						}
						else if( s[0].equals( "_xlnm.Print_Titles" ) )
						{
							try
							{
								bk.getWorkSheet( sh ).getMysheet().setPrintTitles( s[2] ); //addresses[k]);
							}
							catch( OutOfMemoryError e )
							{
								log.error( "OOXMLAdapter.parse OOME setting PrintTitles", e );
								throw e;
							}
						}
						// TODO: handle other built-in named ranges
						// _xlnm._FilterDatabase, _xlnm.Criteria, _xlnm.Extract
//                      }
					}
					else
					{
						if( !s[2].startsWith( "[" ) )
						{ // skip names in external workbooks
							int scope = 0;
							if( !s[1].equals( "" ) )
							{
								scope = (Integer.parseInt( s[1] ) + 1);
							}
							new Name( bk.getWorkBook(), s[0], s[2], (scope) );
						}
					}
				}
				catch( NumberFormatException es )
				{
					; // this is usually a named range that is currently #REF!
				}
				catch( Exception e )
				{
					log.warn( "Failed to addNames", e );
					//log.error("OOXMLAdapter.parse: failed creating Named Range:" + e.toString() + s[0] + ":" + s[2]);
				}
			}
			else
			{
				log.error( "OOXMLAdapter.parse: failed retrieving Named Range" );
			}
		}
	}

	/**
	 * given a HashMap of inline Strings per cell address, set cell value to string
	 * <br>NOTE: cells must exist with proper format before calling this method
	 *
	 * @param bk
	 * @param inlineStrs HashMap
	 */
	static void addInlineStrings( WorkBookHandle bk, HashMap inlineStrs )
	{
		Iterator ii = inlineStrs.keySet().iterator();
		while( ii.hasNext() )
		{
			String cellAddr = (String) ii.next();
			String s = (String) inlineStrs.get( cellAddr );
			int[] rc = ExcelTools.getRowColFromString( cellAddr );
			try
			{
				CellHandle ch = bk.getCell( cellAddr );   // should have been added already
				ch.setVal( s );
			}
			catch( Exception ex )
			{
				;
			}
		}
	}

	/**
	 * intercept Sheet adds and hand off to parse event listener as needed
	 */
	protected static CellHandle sheetAdd( WorkSheetHandle sheet, Object val, Object cachedval, int r, int c, int fmtid )
	{
		CellHandle ch = sheetAdd( sheet, val, r, c, fmtid );
		((Formula) ch.getCell()).setCachedValue( cachedval );
		return ch;
	}

	/**
	 * given an array list of every formula in the workbook, iterate list, parse and add approrpriately
	 *
	 * @param bk
	 * @param formulas
	 */
	void addFormulas( WorkBookHandle bk, ArrayList formulas )
	{
		// after sheets, now can input formulas
		WorkSheetHandle sheet = null;
		HashMap sharedFormulas = new HashMap();
		for( Object formula : formulas )
		{
			String[] s = (String[]) formula;
			//formulas:  0=sheetname, 1= cell address, 2=formula including =, 3=shared formula index, 4=array refs, 5=formula type, 6=calculate always flag, 7=format id, 8=cached value
			if( (s[0].equals( "" ) || s[1].equals( "" )) || (s.length < 8) )
			{
				continue; // no address or formula - should ever happen?
			}
			try
			{
				// for clarity, assign values to most common ops
				String addr = s[1];
				int[] rc = ExcelTools.getRowColFromString( addr );
				String fStr = s[2];
				String type = s[5];
				String fType = "";
				if( s[5].indexOf( '/' ) > 0 )
				{
					type = s[5].split( "/" )[0];
					fType = s[5].split( "/" )[1];
				}
				int fmtid = 0;
				try
				{
					fmtid = Integer.valueOf( s[7] );
				}
				catch( Exception e )
				{
					;
				}
				Object cachedValue = s[8];
				if( type.equals( "n" ) )
				{
					try
					{
						cachedValue = Integer.valueOf( (String) cachedValue );
					}
					catch( NumberFormatException e )
					{
						cachedValue = new Double( (String) cachedValue );
					}
				}
				else if( type.equals( "b" ) )
				{
					cachedValue = Boolean.valueOf( (String) cachedValue );
				}
				// type e -- input calculation exception?
				CellHandle ch = null;  // normal case but may be created * as a blank * if part of a merged cell range or dv ... 
				try
				{
					sheet = bk.getWorkSheet( s[0] );
					ch = sheet.getCell( addr );   // if exists, grab it;                               
				}
				catch( Exception ex )
				{
					;
				}
				if( fStr.equals( "null" ) )
				{ // when would this ever occur?
					log.warn( "OOXMLAdapter.parse: invalid formula encountered at " + addr );
				}

				if( fType.equals( "array" ) )
				{
	                  /*
                       * For a multi-cell formula, the r attribute of the top-left cell 
                       * of the range 1 of cells to which that formula applies
                         shall designate the range of cells to which that formula applies
                       */
					int[] arrayref = null;
					if( s[4] != null )
					{   // if has the ref attribute means its the PARENT array formula
						sheet.getMysheet().addParentArrayRef( s[1], s[4] );
						arrayref = ExcelTools.getRangeRowCol( s[4] );
					}
					else
					{
						arrayref = rc;
					}
                      /* must enter array formulas for each cell in range denoted by array ref*/
					for( int r = arrayref[0]; r <= arrayref[2]; r++ )
					{
						for( int c = arrayref[1]; c <= arrayref[3]; c++ )
						{
							try
							{
								ch = sheet.getCell( r, c );   // if exists, grab it;                               
							}
							catch( Exception ex )
							{
								;
							}
							if( ch == null )
							{
								ch = sheetAdd( sheet, "{" + fStr + "}", cachedValue, r, c, fmtid );
							}
							else
							{
								ch.setFormatId( fmtid ); // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
								ch.setFormula( "{" + fStr + "}",
								               cachedValue ); // set cached value so don't have to recalculate; just sets cached value if formula is already set
							}
						}
					}
				}
				else if( fType.equals( "datatable" ) )
				{
					if( ch == null )
					{
						ch = sheetAdd( sheet, fStr, cachedValue, rc[0], rc[1], fmtid );
					}
					else
					{
						ch.setFormatId( fmtid ); // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
						ch.setFormula( fStr,
						               cachedValue );      // set cached value so don't have to recalculate; just sets cached value if formula is already set
					}
				}
				else if( fType.equals( "shared" ) && !s[3].equals( "" ) )
				{        // meaning that if it's set as shared but doesn't have a shared index, make regular function -- is what excel 2007 does :)  
					// Shared Formulas: there is the "master" shared formula which defines the formula + the range (=ref) of cells that the formula refers to
					// For references to the shared formula, the si index denotes the shared formula it refers to
					// one takes the master formula cell, compares with the current cell's address and increments the references in the master shared
					// formula accordingly -- algorithm of comparison and movement can be tricky
					Integer si = Integer.valueOf( s[3] );
					if( !sharedFormulas.containsKey( si ) )
					{
						// represents the "master" formula of a shared formula, movement is based upon relationship of subsequent cells to this cell
						if( ch == null )
						{
							ch = sheetAdd( sheet, fStr, cachedValue, rc[0], rc[1], fmtid );
						}
						else
						{
							ch.setFormatId( fmtid ); // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
							ch.setFormula( fStr,
							               cachedValue );  // set cached value so don't have to recalculate; just sets cached value if formula is already set
						}
						// see if it's a 3d range
						int[] range = ExcelTools.getRangeCoords( s[3] );
						range[0] -= 1;
						range[2] -= 1;
						Stack expressionStack = cloneStack( ch.getFormulaHandle().getFormulaRec().getExpression() );
						sharedFormulas.put( si, new Object[]{ expressionStack, rc, range } );
					}
					else
					{ // found shared formula- means already created; must get original and "move" based on position of this - the child - shared formula
						Object[] o = (Object[]) sharedFormulas.get( si );
						Stack ss = cloneStack( (Stack) o[0] );
						int[] rcOrig = ((int[]) o[1]);
						Formula.incrementSharedFormula( ss, rc[0] - rcOrig[0], rc[1] - rcOrig[1], (int[]) o[2] );

						if( ch == null )
						{
							ch = sheetAdd( sheet,
							               "=0",
							               null,
							               rc[0],
							               rc[1],
							               fmtid ); // add a basic formula; will be "overwritten" by expression, set below 
							ch.setFormula( ss,
							               cachedValue ); // must set child shared formulas via expression rather than via formula string as original formula string must be incremented 
						}
						else
						{
							ch.setFormula( ss,
							               cachedValue );   // must set child shared formulas via expression rather than via formula string as original formula string must be incremented 
							ch.setFormatId( fmtid ); // if exists may be part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
						}
					}
				}
				else
				{// it's a regular function                                 
					if( ch == null )
					// use parser-aware method
					{
						ch = sheetAdd( sheet, fStr, cachedValue, rc[0], rc[1], fmtid );
					}
					else
					{
						ch.setFormatId( fmtid ); // if exists most likely is part of a merged cell range and therefore it's correct format id may NOT have been set; see mergedranges in parseSheetXML below
						ch.setFormula( fStr,
						               cachedValue );      // set cached value so don't have to recalculate; just sets cached value if formula is already set
					}

				}
				if( (s[6] != null) && (ch != null) )
				{  // for formulas such as =TODAY
					BiffRec br = ch.getCell();
					if( br instanceof Formula )
					{
						ch.getFormulaHandle().setCalcAlways( true );
					}
				}
			}
			catch( FunctionNotSupportedException e )
			{
				log.error( "OOXMLAdapter.parse: failed setting formula " + s[1] + " to cell " + s[0] + ": " + e.toString() );
			}
			catch( Exception e )
			{
				log.error( "OOXMLAdapter.parse: failed setting formula " + s[1] + " to cell " + s[0] + ": " + e.toString() );
			}
		}
	}

	/**
	 * after all sheet data, etc is added, now add pivot tables
	 *
	 * @param bk          WorkBookHandle
	 * @param zip         open ZipFile
	 * @param pivotTables Strings name pivot table files within zip
	 */
	static void addPivotTables( WorkBookHandle bk, ZipFile zip, HashMap<String, WorkSheetHandle> pivotTables ) throws IOException
	{
		Iterator ii = pivotTables.keySet().iterator();
		while( ii.hasNext() )
		{
			String key = (String) ii.next();
			ZipEntry target = zip.getEntry( key );
/*            target= getEntry(zip,p + "_rels/" + c[1].substring(c[1].lastIndexOf("/")+1)+".rels");
        	ArrayList ptrels= parseRels(wrapInputStream(wrapInputStream(zip.getInputStream(target))));
        	if (ptrels.size() > 1) {	// what could this be?
        		Logger.logWarn("OOXMLReader.parse: Unknown Pivot Table Association: " + ptrels.get(1));
        	} 
        	String pcd= ((String[])ptrels.get(0))[1];
        	pcd= pcd.substring(pcd.lastIndexOf("/")+1);
        	Object cacheid= null;
            for (int z= 0; z < pivotCaches.size(); z++) {
        		Object[] o= (Object[]) pivotCaches.get(z);
        		if (pcd.equals(o[0])) {
        			cacheid= o[1];
        			break;
        		}                				
            }
        	
        	target = getEntry(zip,p + c[1]);*/
			WorkSheetHandle sheet = pivotTables.get( key );
			PivotTableDefinition.parseOOXML( bk, /*cacheid, */sheet.getMysheet(), wrapInputStream( zip.getInputStream( target ) ) );

		}
	}

	/**
	 * retrieve pass-through files (Files not processed by normal WBH channels) for later writing
	 *
	 * @param zipIn
	 * @param externalDir
	 */
	public static void refreshExternalFiles( ZipFile zipIn, String externalDir )
	{
		Enumeration<? extends java.util.zip.ZipEntry> ee = zipIn.entries();
		while( ee.hasMoreElements() )
		{
			ZipEntry ze = ee.nextElement();
			String zename = ze.getName();
			// these elements are handled, all else is not
			if( !(zename.equals( "xl/workbook.xml" ) ||
					zename.equals( "xl/styles.xml" ) ||
					zename.equals( "xl/sharedStrings.xml" ) ||
					zename.equals( "[Content_Types].xml" ) ||
					zename.equals( "_rels/.rels" ) ||
					zename.equals( "xl/workbook.xml.rels" ) ||
					zename.startsWith( "xl/charts" ) ||
					//zename.startsWith("xl/drawings") || may be am embed for a chart ...
					zename.startsWith( "xl/worksheets" )) )
			{
				try
				{
					int z = zename.lastIndexOf( "/" );
					OOXMLReader.passThrough( zipIn,
					                         zename,
					                         externalDir + zename.substring( z ) ); // save the original target file for later re-packaging
				}
				catch( Exception e )
				{
					log.error( "OOXMLReader.refreshExternalFiles: error retrieving zip entries: " + e.toString() );
				}

			}
		}

		// docProps
		// xl/media
		// xl/printerSettings
		// xl/theme
		// xl/activeX
		// NOT: xl/charts, xl/drawings, xl/worksheets, xl/_rels, xl/workbook.xml, xl/styles.xml, comments, sharedStrings
		// ?? drawngs/vmlDrawingX.xml
		// xl/

	}

	/**
	 * utility method which looks up a string rid and returns the associated object
	 * in a list of Object[] s
	 *
	 * @param lst source ArrayList
	 * @param rid String rid
	 * @return
	 */
	private static Object lookupRid( ArrayList lst, String rid )
	{
		for( Object aLst : lst )
		{
			Object[] o = (Object[]) aLst;
			if( rid.equals( o[0] ) )
			{
				return (o[1]);
			}
		}
		return null;
	}
}
