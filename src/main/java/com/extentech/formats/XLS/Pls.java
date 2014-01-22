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

/**
 * <b>Pls: Environment-Specific Print Record (4Dh)</b><br>
 * <p/>
 * PLS saves printer settings and driver info
 * <p/>
 * <p><pre>
 * <p/>
 * offset  name            size    contents
 * ---
 * 4       wEnv            2       Operating Environment
 * 0 = Windows
 * 1 = Mac
 * 6       rgb             var     DEVMODE Structure (see MS Win SDK)
 * <p/>
 * reserved (2 bytes): MUST be zero, and MUST be ignored.
 * rgb (variable): A DEVMODE structure, as defined in [DEVMODE],
 * which specifies the printer settings.
 * The size of this field is equal to the size of the current record
 * and all of the following Continue records, excluding the record's heading and reserved field.
 * <p/>
 * </p></pre>
 */
public final class Pls extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1949819999811121013L;

	// TODO: implement this class
	public void init()
	{
		super.init();
		getData();
	}

}

/* more documenation:
DEVMODE Structure

The DEVMODE data structure contains information about the initialization and environment of a printer or a display device.
Syntax

typedef struct _devicemode {
  TCHAR dmDeviceName[CCHDEVICENAME];
  WORD  dmSpecVersion;
  WORD  dmDriverVersion;
  WORD  dmSize;
  WORD  dmDriverExtra;
  DWORD dmFields;
  union {
    struct {
      short dmOrientation;
      short dmPaperSize;
      short dmPaperLength;
      short dmPaperWidth;
      short dmScale;
      short dmCopies;
      short dmDefaultSource;
      short dmPrintQuality;
    };
    struct {
      POINTL dmPosition;
      DWORD  dmDisplayOrientation;
      DWORD  dmDisplayFixedOutput;
    };
  };
  short dmColor;
  short dmDuplex;
  short dmYResolution;
  short dmTTOption;
  short dmCollate;
  TCHAR dmFormName[CCHFORMNAME];
  WORD  dmLogPixels;
  DWORD dmBitsPerPel;
  DWORD dmPelsWidth;
  DWORD dmPelsHeight;
  union {
    DWORD dmDisplayFlags;
    DWORD dmNup;
  };
  DWORD dmDisplayFrequency;
#if (WINVER >= 0x0400)
  DWORD dmICMMethod;
  DWORD dmICMIntent;
  DWORD dmMediaType;
  DWORD dmDitherType;
  DWORD dmReserved1;
  DWORD dmReserved2;
#if (WINVER >= 0x0500) || (_WIN32_WINNT >= 0x0400)
  DWORD dmPanningWidth;
  DWORD dmPanningHeight;
#endif 
#endif 
} DEVMODE, *PDEVMODE, *LPDEVMODE;

Members

dmDeviceName

    A zero-terminated character array that specifies the "friendly" name of the printer or display; for example, "PCL/HP LaserJet" in the case of PCL/HP LaserJet. This string is unique among device drivers. Note that this name may be truncated to fit in the dmDeviceName array.
dmSpecVersion

    The version number of the initialization data specification on which the structure is based. To ensure the correct version is used for any operating system, use DM_SPECVERSION.
dmDriverVersion

    The driver version number assigned by the driver developer.
dmSize

    Specifies the size, in bytes, of the DEVMODE structure, not including any private driver-specific data that might follow the structure's public members. Set this member to sizeof (DEVMODE) to indicate the version of the DEVMODE structure being used.
dmDriverExtra

    Contains the number of bytes of private driver-data that follow this structure. If a device driver does not use device-specific information, set this member to zero.
dmFields

    Specifies whether certain members of the DEVMODE structure have been initialized. If a member is initialized, its corresponding bit is set, otherwise the bit is clear. A driver supports only those DEVMODE members that are appropriate for the printer or display technology.

    The following values are defined, and are listed here with the corresponding structure members.
    Value	Structure member
    DM_ORIENTATION	dmOrientation
    DM_PAPERSIZE	dmPaperSize
    DM_PAPERLENGTH	dmPaperLength
    DM_PAPERWIDTH	dmPaperWidth
    DM_SCALE	dmScale
    DM_COPIES	dmCopies
    DM_DEFAULTSOURCE	dmDefaultSource
    DM_PRINTQUALITY	dmPrintQuality
    DM_POSITION	dmPosition
    DM_DISPLAYORIENTATION	dmDisplayOrientation
    DM_DISPLAYFIXEDOUTPUT	dmDisplayFixedOutput
    DM_COLOR	dmColor
    DM_DUPLEX	dmDuplex
    DM_YRESOLUTION	dmYResolution
    DM_TTOPTION	dmTTOption
    DM_COLLATE	dmCollate
    DM_FORMNAME	dmFormName
    DM_LOGPIXELS	dmLogPixels
    DM_BITSPERPEL	dmBitsPerPel
    DM_PELSWIDTH	dmPelsWidth
    DM_PELSHEIGHT	dmPelsHeight
    DM_DISPLAYFLAGS	dmDisplayFlags
    DM_NUP	dmNup
    DM_DISPLAYFREQUENCY	dmDisplayFrequency
    DM_ICMMETHOD	dmICMMethod
    DM_ICMINTENT	dmICMIntent
    DM_MEDIATYPE	dmMediaType
    DM_DITHERTYPE	dmDitherType
    DM_PANNINGWIDTH	dmPanningWidth
    DM_PANNINGHEIGHT	dmPanningHeight

     
dmOrientation

    For printer devices only, selects the orientation of the paper. This member can be either DMORIENT_PORTRAIT (1) or DMORIENT_LANDSCAPE (2).
dmPaperSize

    For printer devices only, selects the size of the paper to print on. This member can be set to zero if the length and width of the paper are both set by the dmPaperLength and dmPaperWidth members. Otherwise, the dmPaperSize member can be set to a device specific value greater than or equal to DMPAPER_USER or to one of the following predefined values.
    Value	Meaning
    DMPAPER_LETTER	Letter, 8 1/2- by 11-inches
    DMPAPER_LEGAL	Legal, 8 1/2- by 14-inches
    DMPAPER_9X11	9- by 11-inch sheet
    DMPAPER_10X11	10- by 11-inch sheet
    DMPAPER_10X14	10- by 14-inch sheet
    DMPAPER_15X11	15- by 11-inch sheet
    DMPAPER_11X17	11- by 17-inch sheet
    DMPAPER_12X11	12- by 11-inch sheet
    DMPAPER_A2	A2 sheet, 420 x 594-millimeters
    DMPAPER_A3	A3 sheet, 297- by 420-millimeters
    DMPAPER_A3_EXTRA	A3 Extra 322 x 445-millimeters
    DMPAPER_A3_EXTRA_TRAVERSE	A3 Extra Transverse 322 x 445-millimeters
    DMPAPER_A3_ROTATED	A3 rotated sheet, 420- by 297-millimeters
    DMPAPER_A3_TRAVERSE	A3 Transverse 297 x 420-millimeters
    DMPAPER_A4	A4 sheet, 210- by 297-millimeters
    DMPAPER_A4_EXTRA	A4 sheet, 9.27 x 12.69 inches
    DMPAPER_A4_PLUS	A4 Plus 210 x 330-millimeters
    DMPAPER_A4_ROTATED	A4 rotated sheet, 297- by 210-millimeters
    DMPAPER_A4SMALL	A4 small sheet, 210- by 297-millimeters
    DMPAPER_A4_TRANSVERSE	A4 Transverse 210 x 297 millimeters
    DMPAPER_A5	A5 sheet, 148- by 210-millimeters
    DMPAPER_A5_EXTRA	A5 Extra 174 x 235-millimeters
    DMPAPER_A5_ROTATED	A5 rotated sheet, 210- by 148-millimeters
    DMPAPER_A5_TRANSVERSE	A5 Transverse 148 x 210-millimeters
    DMPAPER_A6	A6 sheet, 105- by 148-millimeters
    DMPAPER_A6_ROTATED	A6 rotated sheet, 148- by 105-millimeters
    DMPAPER_A_PLUS	SuperA/A4 227 x 356 -millimeters
    DMPAPER_B4	B4 sheet, 250- by 354-millimeters
    DMPAPER_B4_JIS_ROTATED	B4 (JIS) rotated sheet, 364- by 257-millimeters
    DMPAPER_B5	B5 sheet, 182- by 257-millimeter paper
    DMPAPER_B5_EXTRA	B5 (ISO) Extra 201 x 276-millimeters
    DMPAPER_B5_JIS_ROTATED	B5 (JIS) rotated sheet, 257- by 182-millimeters
    DMPAPER_B6_JIS	B6 (JIS) sheet, 128- by 182-millimeters
    DMPAPER_B6_JIS_ROTATED	B6 (JIS) rotated sheet, 182- by 128-millimeters
    DMPAPER_B_PLUS	SuperB/A3 305 x 487-millimeters
    DMPAPER_CSHEET	C Sheet, 17- by 22-inches
    DMPAPER_DBL_JAPANESE_POSTCARD	Double Japanese Postcard, 200- by 148-millimeters
    DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED	Double Japanese Postcard Rotated, 148- by 200-millimeters
    DMPAPER_DSHEET	D Sheet, 22- by 34-inches
    DMPAPER_ENV_9	#9 Envelope, 3 7/8- by 8 7/8-inches
    DMPAPER_ENV_10	#10 Envelope, 4 1/8- by 9 1/2-inches
    DMPAPER_ENV_11	#11 Envelope, 4 1/2- by 10 3/8-inches
    DMPAPER_ENV_12	#12 Envelope, 4 3/4- by 11-inches
    DMPAPER_ENV_14	#14 Envelope, 5- by 11 1/2-inches
    DMPAPER_ENV_C5	C5 Envelope, 162- by 229-millimeters
    DMPAPER_ENV_C3	C3 Envelope, 324- by 458-millimeters
    DMPAPER_ENV_C4	C4 Envelope, 229- by 324-millimeters
    DMPAPER_ENV_C6	C6 Envelope, 114- by 162-millimeters
    DMPAPER_ENV_C65	C65 Envelope, 114- by 229-millimeters
    DMPAPER_ENV_B4	B4 Envelope, 250- by 353-millimeters
    DMPAPER_ENV_B5	B5 Envelope, 176- by 250-millimeters
    DMPAPER_ENV_B6	B6 Envelope, 176- by 125-millimeters
    DMPAPER_ENV_DL	DL Envelope, 110- by 220-millimeters
    DMPAPER_ENV_INVITE	Envelope Invite 220 x 220 mm
    DMPAPER_ENV_ITALY	Italy Envelope, 110- by 230-millimeters
    DMPAPER_ENV_MONARCH	Monarch Envelope, 3 7/8- by 7 1/2-inches
    DMPAPER_ENV_PERSONAL	6 3/4 Envelope, 3 5/8- by 6 1/2-inches
    DMPAPER_ESHEET	E Sheet, 34- by 44-inches
    DMPAPER_EXECUTIVE	Executive, 7 1/4- by 10 1/2-inches
    DMPAPER_FANFOLD_US	US Std Fanfold, 14 7/8- by 11-inches
    DMPAPER_FANFOLD_STD_GERMAN	German Std Fanfold, 8 1/2- by 12-inches
    DMPAPER_FANFOLD_LGL_GERMAN	German Legal Fanfold, 8 - by 13-inches
    DMPAPER_FOLIO	Folio, 8 1/2- by 13-inch paper
    DMPAPER_ISO_B4	B4 (ISO) 250- by 353-millimeters paper
    DMPAPER_JAPANESE_POSTCARD	Japanese Postcard, 100- by 148-millimeters
    DMPAPER_JAPANESE_POSTCARD_ROTATED	Japanese Postcard Rotated, 148- by 100-millimeters
    DMPAPER_JENV_CHOU3	Japanese Envelope Chou #3
    DMPAPER_JENV_CHOU3_ROTATED	Japanese Envelope Chou #3 Rotated
    DMPAPER_JENV_CHOU4	Japanese Envelope Chou #4
    DMPAPER_JENV_CHOU4_ROTATED	Japanese Envelope Chou #4 Rotated
    DMPAPER_JENV_KAKU2	Japanese Envelope Kaku #2
    DMPAPER_JENV_KAKU2_ROTATED	Japanese Envelope Kaku #2 Rotated
    DMPAPER_JENV_KAKU3	Japanese Envelope Kaku #3
    DMPAPER_JENV_KAKU3_ROTATED	Japanese Envelope Kaku #3 Rotated
    DMPAPER_JENV_YOU4	Japanese Envelope You #4
    DMPAPER_JENV_YOU4_ROTATED	Japanese Envelope You #4 Rotated
    DMPAPER_LAST	DMPAPER_PENV_10_ROTATED
    DMPAPER_LEDGER	Ledger, 17- by 11-inches
    DMPAPER_LEGAL_EXTRA	Legal Extra 9 1/2 x 15 inches.
    DMPAPER_LETTER_EXTRA	Letter Extra 9 1/2 x 12 inches.
    DMPAPER_LETTER_EXTRA_TRANSVERSE	Letter Extra Transverse 9 1/2 x 12 inches.
    DMPAPER_LETTER_ROTATED	Letter Rotated 11 by 8 1/2 inches
    DMPAPER_LETTERSMALL	Letter Small, 8 1/2- by 11-inches
    DMPAPER_LETTER_TRANSVERSE	Letter Transverse 8 1/2 x 11-inches
    DMPAPER_NOTE	Note, 8 1/2- by 11-inches
    DMPAPER_P16K	PRC 16K, 146- by 215-millimeters
    DMPAPER_P16K_ROTATED	PRC 16K Rotated, 215- by 146-millimeters
    DMPAPER_P32K	PRC 32K, 97- by 151-millimeters
    DMPAPER_P32K_ROTATED	PRC 32K Rotated, 151- by 97-millimeters
    DMPAPER_P32KBIG	PRC 32K(Big) 97- by 151-millimeters
    DMPAPER_P32KBIG_ROTATED	PRC 32K(Big) Rotated, 151- by 97-millimeters
    DMPAPER_PENV_1	PRC Envelope #1, 102- by 165-millimeters
    DMPAPER_PENV_1_ROTATED	PRC Envelope #1 Rotated, 165- by 102-millimeters
    DMPAPER_PENV_2	PRC Envelope #2, 102- by 176-millimeters
    DMPAPER_PENV_2_ROTATED	PRC Envelope #2 Rotated, 176- by 102-millimeters
    DMPAPER_PENV_3	PRC Envelope #3, 125- by 176-millimeters
    DMPAPER_PENV_3_ROTATED	PRC Envelope #3 Rotated, 176- by 125-millimeters
    DMPAPER_PENV_4	PRC Envelope #4, 110- by 208-millimeters
    DMPAPER_PENV_4_ROTATED	PRC Envelope #4 Rotated, 208- by 110-millimeters
    DMPAPER_PENV_5	PRC Envelope #5, 110- by 220-millimeters
    DMPAPER_PENV_5_ROTATED	PRC Envelope #5 Rotated, 220- by 110-millimeters
    DMPAPER_PENV_6	PRC Envelope #6, 120- by 230-millimeters
    DMPAPER_PENV_6_ROTATED	PRC Envelope #6 Rotated, 230- by 120-millimeters
    DMPAPER_PENV_7	PRC Envelope #7, 160- by 230-millimeters
    DMPAPER_PENV_7_ROTATED	PRC Envelope #7 Rotated, 230- by 160-millimeters
    DMPAPER_PENV_8	PRC Envelope #8, 120- by 309-millimeters
    DMPAPER_PENV_8_ROTATED	PRC Envelope #8 Rotated, 309- by 120-millimeters
    DMPAPER_PENV_9	PRC Envelope #9, 229- by 324-millimeters
    DMPAPER_PENV_9_ROTATED	PRC Envelope #9 Rotated, 324- by 229-millimeters
    DMPAPER_PENV_10	PRC Envelope #10, 324- by 458-millimeters
    DMPAPER_PENV_10_ROTATED	PRC Envelope #10 Rotated, 458- by 324-millimeters
    DMPAPER_QUARTO	Quarto, 215- by 275-millimeter paper
    DMPAPER_STATEMENT	Statement, 5 1/2- by 8 1/2-inches
    DMPAPER_TABLOID	Tabloid, 11- by 17-inches
    DMPAPER_TABLOID_EXTRA	Tabloid, 11.69 x 18-inches

     
dmPaperLength

    For printer devices only, overrides the length of the paper specified by the dmPaperSize member, either for custom paper sizes or for devices such as dot-matrix printers that can print on a page of arbitrary length. These values, along with all other values in this structure that specify a physical length, are in tenths of a millimeter.
dmPaperWidth

    For printer devices only, overrides the width of the paper specified by the dmPaperSize member.
dmScale

    Specifies the factor by which the printed output is to be scaled. The apparent page size is scaled from the physical page size by a factor of dmScale /100. For example, a letter-sized page with a dmScale value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width.
dmCopies

    Selects the number of copies printed if the device supports multiple-page copies.
dmDefaultSource

    Specifies the paper source. To retrieve a list of the available paper sources for a printer, use the DeviceCapabilities function with the DC_BINS flag.

    This member can be one of the following values, or it can be a device-specific value greater than or equal to DMBIN_USER.

    DMBIN_AUTO
    DMBIN_CASSETTE
    DMBIN_ENVELOPE
    DMBIN_ENVMANUAL
    DMBIN_FIRST
    DMBIN_FORMSOURCE
    DMBIN_LARGECAPACITY
    DMBIN_LARGEFMT
    DMBIN_LAST
    DMBIN_LOWER
    DMBIN_MANUAL
    DMBIN_MIDDLE
    DMBIN_ONLYONE
    DMBIN_TRACTOR
    DMBIN_SMALLFMT
    DMBIN_UPPER

dmPrintQuality

    Specifies the printer resolution. There are four predefined device-independent values:

    DMRES_HIGH
    DMRES_MEDIUM
    DMRES_LOW
    DMRES_DRAFT

    If a positive value is specified, it specifies the number of dots per inch (DPI) and is therefore device dependent.
dmPosition

    For display devices only, a POINTL structure that indicates the positional coordinates of the display device in reference to the desktop area. The primary display device is always located at coordinates (0,0).
dmDisplayOrientation

    For display devices only, the orientation at which images should be presented. If DM_DISPLAYORIENTATION is not set, this member must be zero. If DM_DISPLAYORIENTATION is set, this member must be one of the following values
*/