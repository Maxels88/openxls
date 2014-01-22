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
/**
 *
 */
package com.extentech.formats.OOXML;

/**
 *
 */
public interface OOXMLConstants
{
	public static final String xmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
	public static final String xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	public static final String pkgrelns = "http://schemas.openxmlformats.org/package/2006/relationships";    // .rels ns
	public static final String relns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	public static final String typens = "http://schemas.openxmlformats.org/package/2006/content-types";
	public static final String drawingns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
	public static final String drawingmlns = "http://schemas.openxmlformats.org/drawingml/2006/main";
	public static final String chartns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
	public static final String drawingDir = "xl/drawings";
	public static final String mediaDir = "xl/media";
	public static final String chartDir = "xl/charts";
	public static final String themeDir = "xl/theme";

	public static final String[][] contentTypes = {
			{ "document", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" },
			{ "documentMacroEnabled", "application/vnd.ms-excel.sheet.macroEnabled.main+xml" },
			{ "documentTemplateMacroEnabled", "application/vnd.ms-excel.template.macroEnabled.main+xml" },
			{ "documentTemplate", "application/vnd.openxmlformats-officedocument.spreadsheetml.template+xml" },	/* TODO: is this correct?? */
			{ "sheet", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" },
			{ "styles", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" },
			{ "sst", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" },
			{ "drawing", "application/vnd.openxmlformats-officedocument.drawing+xml" },
			{ "chart", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" },
			{ "theme", "application/vnd.openxmlformats-officedocument.theme+xml" },
			{ "themeOverride", "application/vnd.openxmlformats-officedocument.themeOverride+xml" },
			{ "props", "application/vnd.openxmlformats-package.core-properties+xml" },
			{ "exprops", "application/vnd.openxmlformats-officedocument.extended-properties+xml" },
			{ "custprops", "application/vnd.openxmlformats-officedocument.custom-properties+xml" },
			{ "connections", "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml" },
			{ "calc", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml" },
			{ "vba", "application/vnd.ms-office.vbaProject" },
			{ "vdependencies", "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml" },
			{ "table", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" },
			{ "vmldrawing", "application/vnd.openxmlformats-officedocument.vmlDrawing" },
			{ "externalLink", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml" },
			{ "oleObject", "application/vnd.openxmlformats-officedocument.oleObject" },
			{ "activeX", "application/vnd.ms-office.activeX+xml" },
			{ "comments", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml" },
			{ "image", "image" },
			// images don't go into Content_Types
			{ "userShape", "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml" },
			{ "activeXBinary", "application/vnd.ms-office.activeX" },
			{ "pivotCacheDefinition", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml" },
			{ "pivotCacheRecords", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml" },
			{ "pivotTable", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml" },
	};

	public static final String[][] relsContentTypes = {
			{ "document", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
			{ "documentMacroEnabled", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
			{ "documentTemplate", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
			{ "documentTemplateMacroEnabled", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
			{ "sheet", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" },
			{ "styles", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" },
			{ "sst", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" },
			{ "drawing", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" },
			{ "image", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" },
			{ "chart", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" },
			{ "theme", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" },
			{ "themeOverride", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" },
			{ "props", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" },
			{ "exprops", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" },
			{ "custprops", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" },
			{ "connections", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" },
			{ "calc", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" },
			{ "printerSettings", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" },
			{ "vba", "http://schemas.microsoft.com/office/2006/relationships/vbaProject" },
			{ "vdependencies", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies" },
			{ "table", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" },
			{ "vmldrawing", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" },
			{ "externalLink", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" },
			{ "externalLinkPath", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" },
			{ "macro", "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" },
			{ "hyperlink", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" },
			{ "activeX", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" },
			{ "oleObject", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" },
			{ "comments", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" },
			{ "userShape", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" },
			{ "activeXBinary", "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" },
			{ "pivotCacheDefinition", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" },
			{ "pivotCacheRecords", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" },
			{ "pivotTable", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" },
	};

	public static final String[] builtInNames = {
			"_xlnm.Consolidate_Area", "",
			//protected static final byte AUTO_OPEN = 0x1;
			"",
			//protected static final byte AUTO_CLOSE = 0x2;
			"_xlnm.Extract", "_xlnm.Database", "_xlnm.Criteria", "_xlnm.Print_Area", "_xlnm.Print_Titles", "",
			//protected static final byte RECORDER = 0x8;
			"",
			//protected static final byte DATA_FORM = 0x9;
			"",
			//protected static final byte AUTO_ACTIVATE = 0xA;
			"",
			//protected static final byte AUTO_DEACTIVATE = 0xB;
			"_xlnm.Sheet_Title", "_xlnm._FilterDatabase",
	};

	public static final String[] patternFill = {
			"none",
			"solid",
			"mediumGray",
			"darkGray",
			"lightGray",
			"darkHorizontal",
			"darkVertical",
			"darkDown",
			"darkUp",
			"darkGrid",
			"darkTrellis",
			"lightHorizontal",
			"lightVertical",
			"lightDown",
			"lightUp",
			"lightGrid",
			"lightTrellis",
			"gray125",
			"gray0625",
	};
	public static final String[] horizontalAlignment = {
			"general", "left", "center", "right", "fill", "justify", "centerContinuous",    // 6
			"distributed",
	};

	public static final String[] verticalAlignment = {
			"top", "center", "bottom", "justify", "distributed",
	};

	public static final String[] fontScheme = {
			"none", "major", "minor",
	};
	public static final String[] cellType = {
			"b",                // boolean
			"n",                // number
			"e",                // error
			"s",                // shared (sst) string
			"str",            // formula string
			"inlineStr",        //
	};
	public static final String[] cellFormulaType = {
			"normal", "array", "dataTable", "shared",
	};
	public static final String[] borderStyle = {
			"none",
			"thin",
			"medium",
			"dashed",
			"dotted",
			"thick",
			"double",
			"hair",
			"mediumDashed",
			"dashDot",
			"mediumDashDot",
			"dashDotDot",
			"mediumDashDotDot",
			"slantDashDot",
	};

	public static final String[] fontFamily = {
			"Roman",    //1
			"Swiss", "Modern", "Script", "Decorative",
	};
	/**
	 * indexed via ChartConstants Chart Types
	 *
	 * @see ChartHandle.getOOXML
	 */
	public static final String[] twoDchartTypes = {
			"barChart",    // including col and all shaped bar types such as pyramid, cone ...
			"barChart", "lineChart",    // 2
			"pieChart",    // 3
			"areaChart",    // 4
			"scatterChart",    // 5
			"radarChart",    // 6
			"surfaceChart",    // 7
			"doughnutChart",    // 8
			"bubbleChart",    // 9
			"ofPieChart",    // 10
			"",    // radararea - not used in OOXML
			"",    // pyrmaid == bar in OOXML
			"",    /// cylinder	""
			"",    // cone			""
			"",    // pyramidbar	""	15
			"",    // cylinderbar	""
			"",    // condebar		""
			"stockChart"
	};
	/**
	 * indexed via ChartConstants Chart Types
	 *
	 * @see ChartHandle.getOOXML
	 */
	public static final String[] threeDchartTypes = {
			"bar3DChart",   // including col and all shaped bar types such as pyramid, cone ...
			"bar3DChart", "line3DChart", "pie3DChart", "area3DChart", "",        // scatter - 5 no 3d
			"",        // radar  - 6	no 3d
			"surface3DChart",    // 7
			"",    // doughnut - 8 - no 3d
			"",    // bubble - 9 - no 3d
			"",    // ofPie	10 - no 3d
	};
	public static final String[] legendPos = {
			"b", "l", "r",        // default
			"t", "tr",
	};
/*	   	   
	   public static String[] indexedColors= { 
		    "00000000",
			"00FF0000",
			"0000FF00",
			"000000FF",
			"00FFFF00",
			"00FF00FF",
			"0000FFFF",
			"00000000",
			"00FF0000",
			"0000FF00",
			"000000FF",
			"00FFFF00",
			"00FF00FF",
			"0000FFFF",
			"00800000",
			"00006600",
			"00000080",
			"00669900",
			"00800080",
			"00008080",
			"00EAEAEA",
			"00808080",
			"009999FF",
			"00993366",
			"00FFFFCC",
			"00CCFFFF",
			"00660066",
			"00FF8080",
			"000066CC",
			"00CCCCFF",
			"00000080",
			"00FF00FF",
			"00FFFF00",
			"0000FFFF",
			"00800080",
			"00800000",
			"00008080",
			"000000FF",
			"0000CCFF",
			"00CCFFFF",
			"00CCFFCC",
			"00FFFF99",
			"0099CCFF",
			"00FF99CC",
			"00CC99FF",
			"00FFCC99",
			"003366FF",
			"0033CCCC",
			"0099CC00",
			"00FFCC00",
			"00FF9900",
			"00FF6600",
			"00666699",
			"00969696",
			"00003366",
			"00339966",
			"00003300",
			"00333300",
			"00993300",
			"00993366",
			"00333399",
			"00333333",
	   };
*/

	/**
	 * System Colors
	 * note some of these are guesswork
	 *
	 * @see
	 */
	public static String[][] systemColors = {
			{ "3dDkShadow", "#A0A0A0" },/*(3D Dark System Color) Specifies a Dark shadow color for three-dimensional display elements.*/
			{ "3dLight", "#FFFFFF" },  /* (3D Light System Color) Specifies a Light color for three-dimensional display elements (for edges facing the light source).*/
			{ "activeBorder", "#B4B4B4" }, /* (Active Border System Color) Specifies an Active Window Border Color.*/
			{ "activeCaption", "#99B4D1" }, /* (Active Caption System Color) Specifies the active window title bar color. In particular the left side color in the color gradient of an active
											window's title bar if the gradient effect is enabled.*/
			{ "appWorkspace", "#ABABAB" }, /* (Application Workspace System Color) Specifies the Background color of multiple document interface (MDI) applications.*/
			{ "background", "#FFFFFF" }, /* (Background System Color) Specifies the desktop background color.*/
			{ "btnFace", "#F0F0F0" }, /* (Button Face System Color) Specifies the face color for three-dimensional display elements and for dialog box backgrounds.*/
			{ "btnHighlight", "#FFFFFF" }, /* (Button Highlight System Color) Specifies the highlight color for three-dimensional display elements (for edges facing the light source).*/
			{ "btnShadow", "#A0A0A0" }, /* (Button Shadow System Color) Specifies the shadow color for three-dimensional display elements (for edges facing away from the light source).*/
			{ "btnText", "#000000" }, /* (Button Text System Color) Specifies the color of text on push buttons.*/
			{ "captionText", "#000000" }, /* (Caption Text System Color) Specifies the color of text in the caption, size box, and scroll bar arrow box.*/
			{ "gradientActiveCaption", "#B9D1EA" }, /* (Gradient Active Caption System Color) 	Specifies the right side color in the color gradient of an active window's title bar.*/
			{ "gradientInactiveCaption", "#D7E4F2" }, /* (Gradient Inactive Caption System Color) Specifies the right side color in the color gradient of an inactive window's title bar.*/
			{ "grayText", "#6D6D6D" }, /* (Gray Text System Color) Specifies a grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color.*/
			{ "highlight", "#0000FF" }, /* (Highlight System Color) Specifies the color of Item(s) selected in a control.*/
			{ "highlightText", "#FFFFFF" }, /* (Highlight Text System Color) Specifies the text color of item(s) selected in a control.*/
			{ "hotLight", "#0066CC" }, /* (Hot Light System Color) Specifies the color for a hyperlink or hot-tracked item.*/
			{ "inactiveBorder", "#F4F7FC" }, /* (Inactive Border System Color) Specifies the color of the Inactive window border.*/
			{ "inactiveCaption", "#BFCDDB" }, /* (Inactive Caption System Color) Specifies the color of the Inactive window caption. Specifies the left side color in the color gradient of an
									inactive window's title bar if the gradient effect is enabled.*/
			{ "inactiveCaptionText", "#434E54" }, /* (Inactive Caption Text System Color) Specifies the color of text in an inactive caption.*/
			{ "infoBk", "#FFFFE1" }, /* (Info Back System Color) Specifies the background color for tooltip controls.*/
			{ "infoText", "#000000" }, /* (Info Text System Color) Specifies the text color for tooltip controls.*/
			{ "menu", "#F0F0F0" }, /* (Menu System Color) Specifies the menu background color. */
			{ "menuBar", "#F0F0F0" }, /* (Menu Bar System Color) Specifies the background color for the menu bar when menus appear as flat menus.*/
			{ "menuHighlight", "#3399FF" }, /* (Menu Highlight System Color) Specifies the color used to highlight menu items when the menu appears as a flat menu.*/
			{ "menuText", "#000000" }, /* (Menu Text System Color) Specifies the color of Text in menus.*/
			{ "scrollBar", "#C8C8C8" }, /* (Scroll Bar System Color) Specifies the scroll bar gray area color.*/
			{ "window", "#FFFFFF" }, /* (Window System Color) Specifies window background color.*/
			{ "windowFrame", "#646464" }, /* (Window Frame System Color) Specifies the window frame color.*/
			{ "windowText", "#000000" } /* (Window Text System Color) Specifies the color of text in windows.*/

	};

}