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
package org.openxls.ExtenXLS;

import org.openxls.formats.XLS.BiffRec;
import org.openxls.formats.XLS.BottomMargin;
import org.openxls.formats.XLS.Boundsheet;
import org.openxls.formats.XLS.HCenter;
import org.openxls.formats.XLS.LeftMargin;
import org.openxls.formats.XLS.Name;
import org.openxls.formats.XLS.PrintGrid;
import org.openxls.formats.XLS.PrintRowCol;
import org.openxls.formats.XLS.RightMargin;
import org.openxls.formats.XLS.Setup;
import org.openxls.formats.XLS.TopMargin;
import org.openxls.formats.XLS.VCenter;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.openxls.formats.XLS.WsBool;

import java.util.Iterator;

/**
 * The PrinterSettingsHandle gives you control over the printer settings for a Sheet such as whether to print in landscape or portrait mode.
 * <br/><br/>
 * The PrinterSettingsHandle provides fine-grained control over printing settings
 * in Excel.
 * <br/><br/>
 * NOTE: you can only view the effects of these methods in an open Excel file when
 * you use the "Print Setup" command.
 * <br/><br/>
 * ExtenXLS does not currently support directly sending
 * spreadsheet data to a printer.
 * <p/>
 * <br/><br/><b>
 * Example Usage:</b>
 * <pre>
 * ...
 * PrinterSettingsHandle printersetup = sheet.getPrinterSettings();
 * // Paper Size
 * printersetup.setPaperSize(PrinterSettingsHandle.PAPER_SIZE_LEDGER_17x11);
 * // Scaling
 * printersetup.setScale(125);
 * //	resolution
 * printersetup.setResolution(300);
 * ...
 * </pre>
 */
/* order of records in Page Settings Block:
 * HORIZONTALPAGEBREAKS 
   VERTICALPAGEBREAKS 
   HEADER
   FOOTER
   HCENTER 
   VCENTER 
   LEFTMARGIN 
   RIGHTMARGIN 
   TOPMARGIN 
   BOTTOMMARGIN 
   PLS
   SETUP 
   BITMAP 
 * 
 */
public class PrinterSettingsHandle implements Handle
{

	//	 paper size ints
	public static final int PAPER_SIZE_UNDEFINED = 0;
	public static final int PAPER_SIZE_LETTER_8_5x11 = 1; // Letter 81/2" x 11"
	public static final int PAPER_SIZE_LETTER_SMALL = 2; // Letter small 81/2" x 11"
	public static final int PAPER_SIZE_TABLOID_11x17 = 3; // Tabloid 11" x 17"
	public static final int PAPER_SIZE_LEDGER_17x11 = 4; // Ledger 17" x 11"
	public static final int PAPER_SIZE_LEGAL_8_5x14 = 5; // Legal 81/2" x 14"
	public static final int PAPER_SIZE_STATEMENT_5_5x8_5 = 6; // Statement 51/2" x 81/2"
	public static final int PAPER_SIZE_LETTER_EXTRA_9_5Ax12 = 50; // Letter Extra 91/2" x 12
	public static final int PAPER_SIZE_LEGAL_EXTRA_9_5Ax15 = 51; // Legal Extra 91/2" x 15"
	public static final int PAPER_SIZE_TABLOID_EXTRA_1111_16Ax18 = 52; // Tabloid Extra 1111/16" x 18"
	public static final int PAPER_SIZE_A4_EXTRA_235MM_X_322MM = 53; // A4 Extra 235mm x 322mm
	public static final int PAPER_SIZE_LETTER_TRANSVERSE_8_5Ax11 = 54; // Letter Transverse 81/2" x 11"
	public static final int PAPER_SIZE_EXECUTIVE_7_QUARTER_X_10_5 = 7; // Executive 71/4" x 101/2"
	public static final int PAPER_SIZE_TRANSVERSE_210MM_X_297MM = 55; // A4 Transverse 210mm x 297mm
	public static final int PAPER_SIZE_A3_297MM_X_420MM = 8; // 8 A3 297mm x 420mm
	public static final int PAPER_SIZE_LETTER_EXTRA_TRANSV_9_5_X_12 = 56; // 56 Letter Extra Transv. 91/2" x 12"
	public static final int PAPER_SIZE_A4_210MM_X_297MM = 9; // A4 210mm x 297mm
	public static final int PAPER_SIZE_SUPER_A_A4_227MM_X_356MM = 57; // Super A/A4 227mm x 356mm
	public static final int PAPER_SIZE_A4_SMALL_210MM_X_297MM = 10; // A4 small 210mm x 297mm
	public static final int PAPER_SIZE_SUPER_B_A3_305MM_X_487MM = 58; // Super B/A3 305mm x 487mm
	public static final int PAPER_SIZE_A5_148MM_X_210MM = 11; // A5 148mm x 210mm
	public static final int PAPER_SIZE_LETTER_PLUS = 59; // Letter Plus
	public static final int PAPER_SIZE_2_X_1211_16 = 81; // 2" x 1211/16"
	public static final int PAPER_SIZE_B4_JIS_257MM_X_364MM = 12; // B4 (JIS) 257mm x 364mm
	public static final int PAPER_SIZE_A4_PLUS_210MM_X_330MM = 60; // A4 Plus 210mm x 330mm
	public static final int PAPER_SIZE_B5_JIS_182MM_X_257MM = 13; // B5 (JIS) 182mm x 257mm
	public static final int PAPER_SIZE_A5_TRANSVERSE_148MM_X_210MM = 61; // A5 Transverse 148mm x 210mm
	public static final int PAPER_SIZE_FOLIO_8_5_X_13 = 14; // Folio 81/2" x 13"
	public static final int PAPER_SIZE_B5_JIS_TRANSVERSE_182MM_X_257MM = 62; // B5 (JIS) Transverse 182mm x 257mm
	public static final int PAPER_SIZE_QUATRO_215MM_X_275MM = 15; // Quarto 215mm x 275mm
	public static final int PAPER_SIZE_A3_EXTRA_322MM_X_445MM = 63; // A3 Extra 322mm x 445mm
	public static final int PAPER_SIZE_10Ax14_10_X_14 = 16; // 10x14 10" x 14"
	public static final int PAPER_SIZE_A5_EXTRA_174MM_X_235 = 64; // A5 Extra 174mm x 235mm
	public static final int PAPER_SIZE_11Ax17_11_X_17 = 17; // 11x17 11" x 17"
	public static final int PAPER_SIZE_B5_ISO_EXTRA_201MM_X_276MM = 65; // B5 (ISO) Extra 201mm x 276mm
	public static final int PAPER_SIZE_NOTE_8_5_X_11 = 18; // Note 81/2" x 11"
	public static final int PAPER_SIZE_A2_420MM_X_594MM = 66; // A2 420mm x 594mm
	public static final int PAPER_SIZE_ENVELOPE_9_3_78_X_8_78 = 19; // Envelope #9 37/8" x 87/8"
	public static final int PAPER_SIZE_A3_TRANSVERSE_297MM_X_420MM = 67; // A3 Transverse 297mm x 420mm
	public static final int PAPER_SIZE_ENVELOPE_10_4_18_X_9_5 = 20; // Envelope #10 41/8" x 91/2"
	public static final int PAPER_SIZE_EXTRA_TRANSVERSE_322MM_X_445MM = 68; // A3 Extra Transverse 322mm x 445mm
	public static final int PAPER_SIZE_ENVELOPE_11_4_5_X_10_38 = 21; // Envelope #11 41/2" x 103/8"
	public static final int PAPER_SIZE_DBL_JAP_POSTCARD_200MM_X_148MM = 69; // Dbl. Japanese Postcard 200mm x 148mm
	public static final int PAPER_SIZE_ENVELOPE_12_4_34_X_11 = 22; // Envelope #12 43/4" x 11"
	public static final int PAPER_SIZE_A6_105MM_X_148MM = 70; // A6 105mm x 148mm
	public static final int PAPER_SIZE_ENVELOPE_14_5_X_11_5 = 23; // Envelope #14 5" x 111/2"
	public static final int PAPER_SIZE_C_17_X_22_72 = 24; // C 17" x 22" 72
	public static final int PAPER_SIZE_D_22_X_34_73 = 25; // D 22" x 34" 73
	public static final int PAPER_SIZE_E_34_X_44_74 = 26; // E 34" x 44" 74
	public static final int PAPER_SIZE_DL_ENVELOPE_110MM_X_110MM_X_220MM = 27; // Envelope DL 110mm x 220mm
	public static final int PAPER_SIZE_LETTER_ROTATED_11_X_8_5 = 75; // Letter Rotated 11" x 81/2"
	public static final int PAPER_SIZE_ENVELOPE_C5_162MM_X_229M = 28; // Envelope C5 162mm x 229mm
	public static final int PAPER_SIZE_A3_ROTATED_420MM_X_297MM = 76; // A3 Rotated 420mm x 297mm
	public static final int PAPER_SIZE_ENVELOPE_C3_324MM_X_458MM = 29; // Envelope C3 324mm x 458mm
	public static final int PAPER_SIZE_A4_ROTATED_297MM_X_210MM = 77; // A4 Rotated 297mm x 210mm
	public static final int PAPER_SIZE_ENVELOPE_C4_229MM_X_324MM = 30; // Envelope C4 229mm x 324mm
	public static final int PAPER_SIZE_A5_ROTATED_210MM_X_148MM = 78; // A5 Rotated 210mm x 148mm
	public static final int PAPER_SIZE_ENVELOPE_C6_115MM_X_162MM = 31; // Envelope C6 114mm x 162mm
	public static final int PAPER_SIZE_ENVELOPE_C6_C5_114MM_X_229MM = 32; // Envelope C6/C5 114mm x 229mm
	public static final int PAPER_SIZE_B4_ISO_250MM_X_353MM = 33; // B4 (ISO) 250mm x 353mm
	public static final int PAPER_SIZE_B5_ISO_176MM_X_250MM = 34; // B5 (ISO) 176mm x 250mm
	public static final int PAPER_SIZE_DBL_JAP_POSTCARD_ROT_148MM_X_200MM = 82; // Dbl. Jap. Postcard Rot. 148mm x 200mm
	public static final int PAPER_SIZE_B6_ISO_125MM_X_176MM = 35; // B6 (ISO) 125mm x 176mm
	public static final int PAPER_SIZE_ENVELOPE_ITALY_10MM_X_230MM = 36; // Envelope Italy 110mm x 230mm 84
	public static final int PAPER_SIZE_ENVELOPE_MONARCH_3_7_8_X_7_5 = 37; // Envelope Monarch 37/8" x 71/2" 85
	public static final int PAPER_SIZE_6_3_4_ENVELOPE_3_5_8_X_6_5 = 38; // 63/4 Envelope 35/8" x 61/2" 86
	public static final int PAPER_SIZE_US_STANDARD_FANFOLD_147_8_X_11 = 39; // US Standard Fanfold 147/8" x 11" 87
	public static final int PAPER_SIZE_GERMAN_STD_FANFOLD_8_5_X_12 = 40; // German Std. Fanfold 81/2" x 12"
	public static final int PAPER_SIZE_GERMAN_LEGAL_FANFOLD_8_5_X_13 = 41; // German Legal Fanfold 81/2" x 13"
	// public static final int PAPER_SIZE_B4_ISO_250MM_X_353MM = 42 ; // B4 (ISO) 250mm x 353mm 
	public static final int PAPER_SIZE_JAP_POSTCARD_100M_X_148MM = 43; // Japanese Postcard 100mm x 148mm
	public static final int PAPER_SIZE_9_X_11 = 44; // 9x11 9" x 11"
	public static final int PAPER_SIZE_10_X_11 = 45; // 10x11 10" x 11"
	public static final int PAPER_SIZE_15_X_11 = 46; // 15x11 15" x 11"
	public static final int PAPER_SIZE_ENVELOPE_INVITE_220MM_X_220MM = 47; // Envelope Invite 220mm x 220mm
	public static final int PAPER_SIZE_B4_JIS_ROTATED_364MM_X_257MM = 79; // B4 (JIS) Rotated 364mm x 257mm
	public static final int PAPER_SIZE_B5_JIS_ROTATED_257MMX_X_182MM = 80; // B5 (JIS) Rotated 257mm x 182mm
	public static final int PAPER_SIZE_JAP_POSTCARD_ROT_148MM_X_100MM = 81; // Japanese Postcard Rot. 148mm x 100mm
	public static final int PAPER_SIZE_A6_ROTATED_148MM_X_105MM = 83; // A6 Rotated 148mm x 105mm
	public static final int PAPER_SIZE_B6_JIS_128MM_X_182MM = 88; // B6 (JIS) 128mm x 182mm
	public static final int PAPER_SIZE_B6_JIS_ROT_182MM_X_128MM = 89; // B6 (JIS) Rotated 182mm x 128mm
	public static final int PAPER_SIZE_12_X_11 = 90; // 12x11 12" x 11"

	/**
	 * default constructor
	 */
	public PrinterSettingsHandle( Boundsheet sheet )
	{
		this.sheet = sheet;

		Iterator iter = sheet.getPrintRecs().iterator();
		while( iter.hasNext() )
		{
			BiffRec record = (BiffRec) iter.next();

			if( record instanceof Setup )
			{
				printerSettings = (Setup) record;
			}
			else if( record instanceof HCenter )
			{
				hCenter = (HCenter) record;
			}
			else if( record instanceof VCenter )
			{
				vCenter = (VCenter) record;
			}
			else if( record instanceof LeftMargin )
			{
				leftMargin = (LeftMargin) record;        // missing in default set of records
			}
			else if( record instanceof RightMargin )
			{
				rightMargin = (RightMargin) record;        // missing in default set of records
			}
			else if( record instanceof TopMargin )
			{
				topMargin = (TopMargin) record;            // missing in default set of records
			}
			else if( record instanceof BottomMargin )
			{
				bottomMargin = (BottomMargin) record;    // missing in default set of records
			}
			else if( record instanceof PrintGrid )
			{
				grid = (PrintGrid) record;
			}
			else if( record instanceof PrintRowCol )
			{
				headers = (PrintRowCol) record;
			}
			else if( record instanceof WsBool )
			{
				wsBool = (WsBool) record;
			}
		}
		// Actually, the below comment is incorrect: do NOT do this unconditionally
		//printerSettings.setNoPrintData(false); // there IS printer setup data
	}

	Boundsheet sheet;

	private Setup printerSettings;
	private HCenter hCenter;
	private VCenter vCenter;
	private LeftMargin leftMargin;
	private RightMargin rightMargin;
	private TopMargin topMargin;
	private BottomMargin bottomMargin;
	private PrintGrid grid;
	private PrintRowCol headers;
	private WsBool wsBool;

	// the following are unimplemented printer setting recs:
	// HORIZONTALPAGEBREAKS;
	// VERTICALPAGEBREAKS;

	/**
	 * get the number of copies to print
	 *
	 * @return the number of copies
	 */
	public short getCopies()
	{
		return printerSettings.getCopies();
	}

	/**
	 * get the footer margin size in inches
	 *
	 * @return the footer margin in inches
	 */
	public double getFooterMargin()
	{
		return printerSettings.getFooterMargin();
	}

	/**
	 * get the header margin size in inches
	 *
	 * @return the header margin in inches
	 */
	public double getHeaderMargin()
	{
		return printerSettings.getHeaderMargin();
	}

	/**
	 * Gets the left print margin.
	 */
	public double getLeftMargin()
	{
		if( leftMargin == null )
		{
			leftMargin = new LeftMargin();
			sheet.addMarginRecord( leftMargin );
		}
		return leftMargin.getMargin();
	}

	/**
	 * Gets the right print margin.
	 */
	public double getRightMargin()
	{
		if( rightMargin == null )
		{
			rightMargin = new RightMargin();
			sheet.addMarginRecord( rightMargin );
		}
		return rightMargin.getMargin();
	}

	/**
	 * Gets the top print margin.
	 */
	public double getTopMargin()
	{
		if( topMargin == null )
		{
			topMargin = new TopMargin();
			sheet.addMarginRecord( topMargin );
		}
		return topMargin.getMargin();
	}

	/**
	 * Gets the bottom print margin.
	 */
	public double getBottomMargin()
	{
		if( bottomMargin == null )
		{
			bottomMargin = new BottomMargin();
			sheet.addMarginRecord( bottomMargin );
		}
		return bottomMargin.getMargin();
	}

	/**
	 * get the landscape orientation
	 *
	 * @return whether the print orientation is set to landscape
	 */
	public boolean getLandscape()
	{
		return printerSettings.getLandscape();
	}

	/**
	 * get the left-to-right print orientation
	 *
	 * @return whether the print orientation is set to left-to-right
	 */
	public boolean getLeftToRight()
	{
		return printerSettings.getLeftToRight();
	}

	/**
	 * get whether printing is in black and white
	 *
	 * @return black and white
	 */
	public boolean getNoColor()
	{
		return printerSettings.getNoColor();
	}

	/**
	 * get whether to ignore orientation
	 *
	 * @return ignore orientation
	 */
	public boolean getNoOrient()
	{
		return printerSettings.getNoOrient();
	}

	/**
	 * get whether printer data is missing
	 *
	 * @return whether the printer data is missing
	 */
	public boolean getNoPrintData()
	{
		return printerSettings.getNoPrintData();
	}

	/**
	 * get the page to start printing from
	 *
	 * @return the page to start printing from
	 */
	public short getPageStart()
	{
		return printerSettings.getPageStart();
	}

	/**
	 * Returns the paper size setting for the printer setup based on the
	 * following table:
	 *
	 * @return paper size
	 */
	public short getPaperSize()
	{
		return printerSettings.getPaperSize();
	}

	/**
	 * @return whether to print Notes.
	 */
	public boolean getPrintNotes()
	{
		return printerSettings.getPrintNotes();
	}

	/**
	 * get the print resolution
	 *
	 * @return the printer resolution in DPI
	 */
	public short getResolution()
	{
		return printerSettings.getResolution();
	}

	/**
	 * get the scale of the printer output in whole percentages
	 * <p/>
	 * ie: 25 = 25%
	 *
	 * @return the scale of printer output
	 */
	public short getScale()
	{
		return printerSettings.getScale();
	}

	/**
	 * use custom start page for auto numbering
	 *
	 * @return Returns whether to use a custom start page
	 */
	public boolean getUsePage()
	{
		return printerSettings.getUsePage();
	}

	/**
	 * get the vertical print resolution
	 *
	 * @return the vertical printer resolution in DPI
	 */
	public short getVerticalResolution()
	{
		return printerSettings.getVerticalResolution();
	}

	/**
	 * Whether the sheet should be centered horizontally.
	 */
	public boolean isHCenter()
	{
		return hCenter.isHCenter();
	}

	/**
	 * Whether the sheet should be centered vertically.
	 */
	public boolean isVCenter()
	{
		return vCenter.isVCenter();
	}

	/**
	 * Whether the grid lines will be printed.
	 */
	public boolean isPrintGridLines()
	{
		return grid.isPrintGrid();
	}

	/**
	 * Whether the row and column headers will be printed.
	 */
	public boolean isPrintRowColHeaders()
	{
		return headers.isPrintHeaders();
	}

	/**
	 * Gets whether the sheet will be printed fit to some number of pages.
	 */
	public boolean isFitToPage()
	{
		return wsBool.isFitToPage();
	}

	/**
	 * set for draft quality output
	 *
	 * @param whether to use draft quality
	 */
	public void setDraft( boolean b )
	{
		printerSettings.setDraft( b );
	}

	/**
	 * Set the output to print onto this number of pages high
	 * <p/>
	 * ie: setFitHeight(10) will stretch the print out to fit 10
	 * pages high
	 *
	 * @param number of pages to fit to height
	 */
	public void setFitHeight( int numpages )
	{
		printerSettings.setFitHeight( (short) numpages );
	}

	/**
	 * Set the output to print onto this number of pages wide
	 * <p/>
	 * ie: setFitWidth(10) will stretch the print out to fit 10
	 * pages wide
	 *
	 * @param number of pages to fit to width
	 */
	public void setFitWidth( int numpages )
	{
		printerSettings.setFitWidth( (short) numpages );
	}

	/**
	 * Sets whether the sheet will be printed fit to some number of pages.
	 */
	public void setFitToPage( boolean value )
	{
		wsBool.setFitToPage( value );
	}

	/**
	 * sets the footer margin in inches
	 *
	 * @param footer margin
	 */
	public void setFooterMargin( double f )
	{
		printerSettings.setFooterMargin( f );
	}

	/**
	 * Sets whether the page should be centered horizontally.
	 */
	public void setHCenter( boolean center )
	{
		hCenter.setHCenter( center );
	}

	/**
	 * sets the Header margin in inches
	 *
	 * @param header margin
	 */
	public void setHeaderMargin( double h )
	{
		printerSettings.setHeaderMargin( h );
	}

	/**
	 * Sets the sheet's left print margin.
	 */
	public void setLeftMargin( double value )
	{
		if( leftMargin == null )
		{
			leftMargin = new LeftMargin();
			sheet.addMarginRecord( leftMargin );
		}
		leftMargin.setMargin( value );
	}

	/**
	 * Sets the sheet's right print margin.
	 */
	public void setRightMargin( double value )
	{
		if( rightMargin == null )
		{
			rightMargin = new RightMargin();
			sheet.addMarginRecord( rightMargin );
		}
		rightMargin.setMargin( value );
	}

	/**
	 * Sets the sheet's top print margin.
	 */
	public void setTopMargin( double value )
	{
		if( topMargin == null )
		{
			topMargin = new TopMargin();
			sheet.addMarginRecord( topMargin );
		}
		topMargin.setMargin( value );
	}

	/**
	 * Sets the sheet's bottom print margin.
	 */
	public void setBottomMargin( double value )
	{
		if( bottomMargin == null )
		{
			bottomMargin = new BottomMargin();
			sheet.addMarginRecord( bottomMargin );
		}
		bottomMargin.setMargin( value );
	}

	/**
	 * set the print orientation to landscape or portrait
	 *
	 * @param landscape
	 */
	public void setLandscape( boolean b )
	{
		printerSettings.setNoOrient( false ); // use the orientation setting
		printerSettings.setLandscape( b );
	}

	/**
	 * set the print orientation to left-to-right printing
	 *
	 * @param leftToRight
	 */
	public void setLeftToRight( boolean b )
	{
		printerSettings.setLeftToRight( b );
	}

	/**
	 * sets the output to black and white
	 *
	 * @param noColor
	 */
	public void setNoColor( boolean b )
	{
		printerSettings.setNoColor( b );
	}

	/**
	 * set the default page to start printing from
	 *
	 * @param p
	 */
	public void setPageStart( short p )
	{
		printerSettings.setPageStart( p );
	}

	/**
	 * sets the whether to print cell notes
	 *
	 * @param printNotes whether to print Notes.
	 */
	public void setPrintNotes( boolean b )
	{
		printerSettings.setPrintNotes( b );
	}

	/**
	 * Set the output printer resolution
	 *
	 * @param resolution The resolution to set in DPI.
	 */
	public void setResolution( int r )
	{
		printerSettings.setResolution( (short) r );
	}

	/**
	 * scale the printer output in whole percentages
	 * <p/>
	 * ie: 25 = 25%
	 *
	 * @param scale The scale to set.
	 */
	public void setScale( int scale )
	{
		printerSettings.setScale( (short) scale );
	}

	/**
	 * @param usePage whether to use custom Page to start printing from.
	 */
	public void setUsePage( boolean usePage )
	{
		printerSettings.setUsePage( usePage );
	}

	/**
	 * @param verticalResolution The vertical Resolution in DPI
	 */
	public void setVerticalResolution( short verticalResolution )
	{
		printerSettings.setVerticalResolution( verticalResolution );
	}

	/**
	 * @param copies The number of copies to print
	 */
	public void setCopies( int copies )
	{
		printerSettings.setCopies( (short) copies );
	}

	/**
	 * sets the paper size based on the paper size table
	 *
	 * @param the paper size index
	 */
	public void setPaperSize( int p )
	{
		printerSettings.setPaperSize( (short) p );
	}

	/**
	 * get the draft quality setting
	 *
	 * @return draft quality
	 */
	public boolean getDraft()
	{
		return printerSettings.getDraft();
	}

	/**
	 * get the number of pages to fit the printout to height
	 *
	 * @return fit to height
	 */
	public short getFitHeight()
	{
		return printerSettings.getFitHeight();
	}

	/**
	 * get the number of pages to fit the printout to width
	 *
	 * @return fit to width
	 */
	public short getFitWidth()
	{
		return printerSettings.getFitWidth();
	}

	/**
	 * Sets whether the sheet should be centered vertically.
	 */
	public void setVCenter( boolean center )
	{
		vCenter.setVCenter( center );
	}

	/**
	 * Sets whether to print the grid lines.
	 */
	public void setPrintGrid( boolean print )
	{
		grid.setPrintGrid( print );
	}

	/**
	 * Sets whether to print the row and column headers.
	 */
	public void setPrintRowColHeaders( boolean print )
	{
		headers.setPrintHeaders( print );
	}

	/**
	 * Gets the range specifying the titles printed on each page.
	 */
	public String getTitles()
	{
		Name range = sheet.getName( "Built-in: PRINT_TITLES" );
		if( range == null )
		{
			return null;
		}
		return range.getExpressionString();
	}

	/**
	 * Sets the range specifying the titles printed on each page.
	 * The reference for the row(s) to repeat e.g. $1:$1 for row 1
	 * For Columns, type the reference to the column or columns that
	 * you want to set as a title e.g. $A:$B for columns A and B
	 */
	// note:  MUST be in $ROW:$ROW or $COL:$COL format, for both
	// can be $R:$R, $C:$C for both
	public void setTitles( String range )
	{
		Name name = sheet.getName( "Built-in: PRINT_TITLES" );
		if( name == null )
		{
			try
			{
				name = new Name( sheet.getWorkBook(), "Print_Titles" );
				name.setBuiltIn( (byte) 0x07 );    //do before setNewScope as it blows out itab
				name.setNewScope( sheet.getSheetNum() + 1 );
			}
			catch( WorkSheetNotFoundException e )
			{
				// This shouldn't be possible.
				throw new Error( "sheet not found re-scoping name" );
			}
		}
		// pre-process range to ensure in proper format, ensure all absolute ($) refs + 
		// handle wholerow-wholecol refs + complex ranges
		if( range == null )
		{
			return;    // TODO: Do what?? remove??
		}
		String[] ranges = range.split( "," );
		range = "";
		for( int i = 0; i < ranges.length; i++ )
		{
			if( i > 0 )    // concatenate terms into one ptgmemfunc-style expression
			{
				range += ",";
			}
			String r = "";
			int[] rc = ExcelTools.getRangeCoords( ranges[i] );
			if( rc[0] == rc[2] ) // varies by column
			{
				r = "$" + ExcelTools.getAlphaVal( rc[1] ) + ":$" + ExcelTools.getAlphaVal( rc[3] );
			}
			if( rc[1] == rc[3] )
			{// varies by row
				r = "$" + rc[0] + ":$" + rc[2];
			}
			range += sheet.getSheetName() + "!" + r;
		}
		name.setLocation( range );
	}

}