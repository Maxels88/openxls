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
package org.openxls.formats.XLS.formulas;

/**
 * Function Constants for all Formula Types (PtgFunc, PtgFuncVar - regular and PtgFuncVar - AddIns)
 * Modifications:
 * all xlfXXX constants were originally in FunctionHandler
 * getFunctionString was orignally in PtgFuncVar
 * FUNCTION_STRINGS were originally in FunctionHandler
 * getNumVars was originally in PtgFunc
 *
 * @see
 */
public class FunctionConstants
{
	/* HOW TO USE:
 * 
 * 1- If implementing a formula, MAKE SURE to put it in recArr AND getFunctionString list.  If it is a ptgFunc, also input 
 *     the number of args in getNumArgs.  If it's an Add-in, add xlfXX constant to the end of the Excel function numbers list.  
 *     
 *     All function ID's/number MUST exist in the xlfXXX constants list
 * 
 * 
 */
	public static int FTYPE_PTGFUNC = 0;
	public static int FTYPE_PTGFUNCVAR = 1;
	public static int FTYPE_PTGFUNCVAR_ADDIN = 2;

	/**
	 * *************************************
	 * *
	 * Excel function numbers              *
	 * *
	 * **************************************
	 */

	public static final int XLF_COUNT = 0;
	public static final int XLF_IS = 1;
	public static final int XLF_IS_NA = 2;
	public static final int XLF_IS_ERROR = 3;
	public static final int XLF_SUM = 4;
	public static final int XLF_AVERAGE = 5;
	public static final int XLF_MIN = 6;
	public static final int XLF_MAX = 7;
	public static final int XLF_ROW = 8;
	public static final int xlfColumn = 9;
	public static final int xlfNa = 10;
	public static final int xlfNpv = 11;
	public static final int xlfStdev = 12;
	public static final int xlfDollar = 13;
	public static final int xlfFixed = 14;
	public static final int xlfSin = 15;
	public static final int xlfCos = 16;
	public static final int xlfTan = 17;
	public static final int xlfAtan = 18;
	public static final int xlfPi = 19;
	public static final int xlfSqrt = 20;
	public static final int xlfExp = 21;
	public static final int xlfLn = 22;
	public static final int xlfLog10 = 23;
	public static final int xlfAbs = 24;
	public static final int xlfInt = 25;
	public static final int xlfSign = 26;
	public static final int xlfRound = 27;
	public static final int xlfLookup = 28;
	public static final int xlfIndex = 29;
	public static final int xlfRept = 30;
	public static final int xlfMid = 31;
	public static final int xlfLen = 32;
	public static final int xlfValue = 33;
	public static final int xlfTrue = 34;
	public static final int xlfFalse = 35;
	public static final int xlfAnd = 36;
	public static final int xlfOr = 37;
	public static final int xlfNot = 38;
	public static final int xlfMod = 39;
	public static final int xlfDcount = 40;
	public static final int xlfDsum = 41;
	public static final int xlfDaverage = 42;
	public static final int xlfDmin = 43;
	public static final int xlfDmax = 44;
	public static final int xlfDstdev = 45;
	public static final int xlfVar = 46;
	public static final int xlfDvar = 47;
	public static final int xlfText = 48;
	public static final int xlfLinest = 49;
	public static final int xlfTrend = 50;
	public static final int xlfLogest = 51;
	public static final int xlfGrowth = 52;
	public static final int xlfGoto = 53;
	public static final int xlfHalt = 54;
	public static final int xlfPv = 56;
	public static final int xlfFv = 57;
	public static final int xlfNper = 58;
	public static final int xlfPmt = 59;
	public static final int xlfRate = 60;
	public static final int xlfMirr = 61;
	public static final int xlfIrr = 62;
	public static final int xlfRand = 63;
	public static final int xlfMatch = 64;
	public static final int xlfDate = 65;
	public static final int xlfTime = 66;
	public static final int xlfDay = 67;
	public static final int xlfMonth = 68;
	public static final int xlfYear = 69;
	public static final int xlfWeekday = 70;
	public static final int xlfHour = 71;
	public static final int xlfMinute = 72;
	public static final int xlfSecond = 73;
	public static final int xlfNow = 74;
	public static final int xlfAreas = 75;
	public static final int xlfRows = 76;
	public static final int xlfColumns = 77;
	public static final int xlfOffset = 78;
	public static final int xlfAbsref = 79;
	public static final int xlfRelref = 80;
	public static final int xlfArgument = 81;
	public static final int xlfSearch = 82;
	public static final int xlfTranspose = 83;
	public static final int xlfError = 84;
	public static final int xlfStep = 85;
	public static final int xlfType = 86;
	public static final int xlfEcho = 87;
	public static final int xlfSetName = 88;
	public static final int xlfCaller = 89;
	public static final int xlfDeref = 90;
	public static final int xlfWindows = 91;
	public static final int xlfSeries = 92;
	public static final int xlfDocuments = 93;
	public static final int xlfActiveCell = 94;
	public static final int xlfSelection = 95;
	public static final int xlfResult = 96;
	public static final int xlfAtan2 = 97;
	public static final int xlfAsin = 98;
	public static final int xlfAcos = 99;
	public static final int xlfChoose = 100;
	public static final int xlfHlookup = 101;
	public static final int xlfVlookup = 102;
	public static final int xlfLinks = 103;
	public static final int xlfInput = 104;
	public static final int xlfIsref = 105;
	public static final int xlfGetFormula = 106;
	public static final int xlfGetName = 107;
	public static final int xlfSetValue = 108;
	public static final int xlfLog = 109;
	public static final int xlfExec = 110;
	public static final int xlfChar = 111;
	public static final int xlfLower = 112;
	public static final int xlfUpper = 113;
	public static final int xlfProper = 114;
	public static final int xlfLeft = 115;
	public static final int xlfRight = 116;
	public static final int xlfExact = 117;
	public static final int xlfTrim = 118;
	public static final int xlfReplace = 119;
	public static final int xlfSubstitute = 120;
	public static final int xlfCode = 121;
	public static final int xlfNames = 122;
	public static final int xlfDirectory = 123;
	public static final int xlfFind = 124;
	public static final int xlfCell = 125;
	public static final int xlfIserr = 126;
	public static final int xlfIstext = 127;
	public static final int xlfIsnumber = 128;
	public static final int xlfIsblank = 129;
	public static final int xlfT = 130;
	public static final int xlfN = 131;
	public static final int xlfFopen = 132;
	public static final int xlfFclose = 133;
	public static final int xlfFsize = 134;
	public static final int xlfFreadln = 135;
	public static final int xlfFread = 136;
	public static final int xlfFwriteln = 137;
	public static final int xlfFwrite = 138;
	public static final int xlfFpos = 139;
	public static final int xlfDatevalue = 140;
	public static final int xlfTimevalue = 141;
	public static final int xlfSln = 142;
	public static final int xlfSyd = 143;
	public static final int xlfDdb = 144;
	public static final int xlfGetDef = 145;
	public static final int xlfReftext = 146;
	public static final int xlfTextref = 147;
	public static final int XLF_INDIRECT = 148;
	public static final int xlfRegister = 149;
	public static final int xlfCall = 150;
	public static final int xlfAddBar = 151;
	public static final int xlfAddMenu = 152;
	public static final int xlfAddCommand = 153;
	public static final int xlfEnableCommand = 154;
	public static final int xlfCheckCommand = 155;
	public static final int xlfRenameCommand = 156;
	public static final int xlfShowBar = 157;
	public static final int xlfDeleteMenu = 158;
	public static final int xlfDeleteCommand = 159;
	public static final int xlfGetChartItem = 160;
	public static final int xlfDialogBox = 161;
	public static final int xlfClean = 162;
	public static final int xlfMdeterm = 163;
	public static final int xlfMinverse = 164;
	public static final int xlfMmult = 165;
	public static final int xlfFiles = 166;
	public static final int xlfIpmt = 167;
	public static final int xlfPpmt = 168;
	public static final int xlfCounta = 169;
	public static final int xlfCancelKey = 170;
	public static final int xlfInitiate = 175;
	public static final int xlfRequest = 176;
	public static final int xlfPoke = 177;
	public static final int xlfExecute = 178;
	public static final int xlfTerminate = 179;
	public static final int xlfRestart = 180;
	public static final int xlfHelp = 181;
	public static final int xlfGetBar = 182;
	public static final int xlfProduct = 183;
	public static final int xlfFact = 184;
	public static final int xlfGetCell = 185;
	public static final int xlfGetWorkspace = 186;
	public static final int xlfGetWindow = 187;
	public static final int xlfGetDocument = 188;
	public static final int xlfDproduct = 189;
	public static final int xlfIsnontext = 190;
	public static final int xlfGetNote = 191;
	public static final int xlfNote = 192;
	public static final int xlfStdevp = 193;
	public static final int xlfVarp = 194;
	public static final int xlfDstdevp = 195;
	public static final int xlfDvarp = 196;
	public static final int xlfTrunc = 197;
	public static final int xlfIslogical = 198;
	public static final int xlfDcounta = 199;
	public static final int xlfDeleteBar = 200;
	public static final int xlfUnregister = 201;
	public static final int xlfUsdollar = 204;
	public static final int xlfFindb = 205;
	public static final int xlfSearchb = 206;
	public static final int xlfReplaceb = 207;
	public static final int xlfLeftb = 208;
	public static final int xlfRightb = 209;
	public static final int xlfMidb = 210;
	public static final int xlfLenb = 211;
	public static final int xlfRoundup = 212;
	public static final int xlfRounddown = 213;
	public static final int xlfAsc = 214;
	public static final int xlfDbcs = 215;
	public static final int xlfRank = 216;
	public static final int xlfAddress = 219;
	public static final int xlfDays360 = 220;
	public static final int xlfToday = 221;
	public static final int xlfVdb = 222;
	public static final int xlfMedian = 227;
	public static final int xlfSumproduct = 228;
	public static final int xlfSinh = 229;
	public static final int xlfCosh = 230;
	public static final int xlfTanh = 231;
	public static final int xlfAsinh = 232;
	public static final int xlfAcosh = 233;
	public static final int xlfAtanh = 234;
	public static final int xlfDget = 235;
	public static final int xlfCreateObject = 236;
	public static final int xlfVolatile = 237;
	public static final int xlfLastError = 238;
	public static final int xlfCustomUndo = 239;
	public static final int xlfCustomRepeat = 240;
	public static final int xlfFormulaConvert = 241;
	public static final int xlfGetLinkInfo = 242;
	public static final int xlfTextBox = 243;
	public static final int xlfInfo = 244;
	public static final int xlfGroup = 245;
	public static final int xlfGetObject = 246;
	public static final int xlfDb = 247;
	public static final int xlfPause = 248;
	public static final int xlfResume = 251;
	public static final int xlfFrequency = 252;
	public static final int xlfAddToolbar = 253;
	public static final int xlfDeleteToolbar = 254;
	public static final int xlfADDIN = 255;    // KSC: Added; Excel function ID for add-ins
	public static final int xlfResetToolbar = 256;
	public static final int xlfEvaluate = 257;
	public static final int xlfGetToolbar = 258;
	public static final int xlfGetTool = 259;
	public static final int xlfSpellingCheck = 260;
	public static final int xlfErrorType = 261;
	public static final int xlfAppTitle = 262;
	public static final int xlfWindowTitle = 263;
	public static final int xlfSaveToolbar = 264;
	public static final int xlfEnableTool = 265;
	public static final int xlfPressTool = 266;
	public static final int xlfRegisterId = 267;
	public static final int xlfGetWorkbook = 268;
	public static final int xlfAvedev = 269;
	public static final int xlfBetadist = 270;
	public static final int xlfGammaln = 271;
	public static final int xlfBetainv = 272;
	public static final int xlfBinomdist = 273;
	public static final int xlfChidist = 274;
	public static final int xlfChiinv = 275;
	public static final int xlfCombin = 276;
	public static final int xlfConfidence = 277;
	public static final int xlfCritbinom = 278;
	public static final int xlfEven = 279;
	public static final int xlfExpondist = 280;
	public static final int xlfFdist = 281;
	public static final int xlfFinv = 282;
	public static final int xlfFisher = 283;
	public static final int xlfFisherinv = 284;
	public static final int xlfFloor = 285;
	public static final int xlfGammadist = 286;
	public static final int xlfGammainv = 287;
	public static final int xlfCeiling = 288;
	public static final int xlfHypgeomdist = 289;
	public static final int xlfLognormdist = 290;
	public static final int xlfLoginv = 291;
	public static final int xlfNegbinomdist = 292;
	public static final int xlfNormdist = 293;
	public static final int xlfNormsdist = 294;
	public static final int xlfNorminv = 295;
	public static final int xlfNormsinv = 296;
	public static final int xlfStandardize = 297;
	public static final int xlfOdd = 298;
	public static final int xlfPermut = 299;
	public static final int xlfPoisson = 300;
	public static final int xlfTdist = 301;
	public static final int xlfWeibull = 302;
	public static final int xlfSumxmy2 = 303;
	public static final int xlfSumx2my2 = 304;
	public static final int xlfSumx2py2 = 305;
	public static final int xlfChitest = 306;
	public static final int xlfCorrel = 307;
	public static final int xlfCovar = 308;
	public static final int xlfForecast = 309;
	public static final int xlfFtest = 310;
	public static final int xlfIntercept = 311;
	public static final int xlfPearson = 312;
	public static final int xlfRsq = 313;
	public static final int xlfSteyx = 314;
	public static final int xlfSlope = 315;
	public static final int xlfTtest = 316;
	public static final int xlfProb = 317;
	public static final int xlfDevsq = 318;
	public static final int xlfGeomean = 319;
	public static final int xlfHarmean = 320;
	public static final int xlfSumsq = 321;
	public static final int xlfKurt = 322;
	public static final int xlfSkew = 323;
	public static final int xlfZtest = 324;
	public static final int xlfLarge = 325;
	public static final int xlfSmall = 326;
	public static final int xlfQuartile = 327;
	public static final int xlfPercentile = 328;
	public static final int xlfPercentrank = 329;
	public static final int xlfMode = 330;
	public static final int xlfTrimmean = 331;
	public static final int xlfTinv = 332;
	public static final int xlfMovieCommand = 334;
	public static final int xlfGetMovie = 335;
	public static final int xlfConcatenate = 336;
	public static final int xlfPower = 337;
	public static final int xlfPivotAddData = 338;
	public static final int xlfGetPivotTable = 339;
	public static final int xlfGetPivotField = 340;
	public static final int xlfGetPivotItem = 341;
	public static final int xlfRadians = 342;
	public static final int xlfDegrees = 343;
	public static final int xlfSubtotal = 344;
	public static final int XLF_SUM_IF = 345;
	public static final int xlfCountif = 346;
	public static final int xlfCountblank = 347;
	public static final int xlfScenarioGet = 348;
	public static final int xlfOptionsListsGet = 349;
	public static final int xlfIspmt = 350;
	public static final int xlfDatedif = 351;
	public static final int xlfDatestring = 352;
	public static final int xlfNumberstring = 353;
	public static final int xlfRoman = 354;
	public static final int xlfOpenDialog = 355;
	public static final int xlfSaveDialog = 356;
	public static final int xlfViewGet = 357;
	public static final int xlfGetPivotData = 358;
	public static final int xlfHyperlink = 359;
	public static final int xlfPhonetic = 360;
	public static final int xlfAverageA = 361;
	public static final int xlfMaxA = 362;
	public static final int xlfMinA = 363;
	public static final int xlfStDevPA = 364;
	public static final int xlfVarPA = 365;
	public static final int xlfStDevA = 366;
	public static final int xlfVarA = 367;
	// KSC: ADD-IN formulas - use any index; name must be present in FunctionConstants.addIns
	// Financial Formulas
	public static final int xlfAccrintm = 368;
	public static final int xlfAccrint = 369;
	public static final int xlfCoupDayBS = 370;
	public static final int xlfCoupDays = 371;
	public static final int xlfCumIPmt = 372;
	public static final int xlfCumPrinc = 373;
	public static final int xlfCoupNCD = 374;
	public static final int xlfCoupDaysNC = 375;
	public static final int xlfCoupPCD = 376;
	public static final int xlfCoupNUM = 377;
	public static final int xlfDollarDE = 378;
	public static final int xlfDollarFR = 379;
	public static final int xlfEffect = 380;
	public static final int xlfINTRATE = 381;
	public static final int xlfXIRR = 382;
	public static final int xlfXNPV = 383;
	public static final int xlfYIELD = 384;
	public static final int xlfPRICE = 385;
	public static final int xlfPRICEDISC = 386;
	public static final int xlfPRICEMAT = 387;
	public static final int xlfDURATION = 388;
	public static final int xlfMDURATION = 389;
	public static final int xlfTBillEq = 390;
	public static final int xlfTBillPrice = 391;
	public static final int xlfTBillYield = 392;
	public static final int xlfYieldDisc = 393;
	public static final int xlfYieldMat = 394;
	public static final int xlfFVSchedule = 395;
	public static final int xlfAmorlinc = 396;
	public static final int xlfAmordegrc = 397;
	public static final int xlfOddFPrice = 398;
	public static final int xlfOddLPrice = 399;
	public static final int xlfOddFYield = 400;
	public static final int xlfOddLYield = 401;
	public static final int xlfNOMINAL = 402;
	public static final int xlfDISC = 403;
	public static final int xlfRECEIVED = 404;
	// Engineering Formulas
	public static final int xlfBIN2DEC = 405;
	public static final int xlfBIN2HEX = 406;
	public static final int xlfBIN2OCT = 407;
	public static final int xlfDEC2BIN = 408;
	public static final int xlfDEC2HEX = 409;
	public static final int xlfDEC2OCT = 410;
	public static final int xlfHEX2BIN = 411;
	public static final int xlfHEX2DEC = 412;
	public static final int xlfHEX2OCT = 413;
	public static final int xlfOCT2BIN = 414;
	public static final int xlfOCT2DEC = 415;
	public static final int xlfOCT2HEX = 416;
	public static final int xlfCOMPLEX = 417;
	public static final int xlfGESTEP = 418;
	public static final int xlfDELTA = 419;
	public static final int xlfIMAGINARY = 420;
	public static final int xlfIMABS = 421;
	public static final int xlfIMDIV = 422;
	public static final int xlfIMCONJUGATE = 423;
	public static final int xlfIMCOS = 424;
	public static final int xlfIMSIN = 425;
	public static final int xlfIMREAL = 426;
	public static final int xlfIMEXP = 427;
	public static final int xlfIMSUB = 428;
	public static final int xlfIMSUM = 429;
	public static final int xlfIMPRODUCT = 430;
	public static final int xlfIMLN = 431;
	public static final int xlfIMLOG10 = 432;
	public static final int xlfIMLOG2 = 433;
	public static final int xlfIMPOWER = 434;
	public static final int xlfIMSQRT = 435;
	public static final int xlfIMARGUMENT = 436;
	public static final int xlfCONVERT = 437;
	public static final int xlfERF = 460;
	public static final int xlfERFC = 461;
	// Math Add-in Formulas
	public static final int xlfDOUBLEFACT = 438;
	public static final int xlfGCD = 439;
	public static final int xlfLCM = 440;
	public static final int xlfMROUND = 441;
	public static final int xlfMULTINOMIAL = 442;
	public static final int xlfQUOTIENT = 443;
	public static final int xlfRANDBETWEEN = 444;
	public static final int xlfSERIESSUM = 445;
	public static final int xlfSQRTPI = 446;
	public static final int xlfSUMIFS = 456;
	// Information Add-in Formulas
	public static final int xlfISEVEN = 447;
	public static final int xlfISODD = 448;
	// Date/Time Add-in Formulas
	public static final int xlfNETWORKDAYS = 449;
	public static final int xlfEDATE = 450;
	public static final int xlfEOMONTH = 451;
	public static final int xlfWEEKNUM = 452;
	public static final int xlfWORKDAY = 453;
	public static final int xlfYEARFRAC = 459;
	// Statistical
	public static final int xlfAVERAGEIF = 454;
	public static final int xlfAVERAGEIFS = 457;
	public static final int xlfCOUNTIFS = 458;
	// Logical
	public static final int xlfIFERROR = 455;
	public static final int MAXXLF = 462;

	/**
	 * Japanese Excel contains some different values and string output than US English Excel.
	 * <p/>
	 * This recArr is checked if locale = japan... if null value is returned then the main list is checked
	 */
	public static String[][] jRecArr = {
			{ "YEN", String.valueOf( xlfDollar ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "DOLLAR", String.valueOf( xlfUsdollar ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "JIS", String.valueOf( xlfDbcs ), String.valueOf( FTYPE_PTGFUNC ) },
	};

	/**
	 * Unimplemented records.  This exists to allow writing of functions that are unsupported for calculation
	 */
	public static String[][] unimplRecArr = {
			{ "ASC", String.valueOf( xlfAsc ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"DBCS", String.valueOf( xlfDbcs ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "MDETERM", String.valueOf( xlfMdeterm ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SEARCHB", String.valueOf( xlfSearchb ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TRANSPOSE", String.valueOf( xlfTranspose ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "BETAINV", String.valueOf( xlfBetainv ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "BETADIST", String.valueOf( xlfBetadist ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TIMEVALUE", String.valueOf( xlfTimevalue ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "MINVERSE", String.valueOf( xlfMinverse ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "MDETERM", String.valueOf( xlfMdeterm ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "GETPIVOTDATA", String.valueOf( xlfGetPivotData ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "HYPERLINK", String.valueOf( xlfHyperlink ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "PHONETIC", String.valueOf( xlfPhonetic ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "PERCENTILE", String.valueOf( xlfPercentile ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TRUNC", String.valueOf( xlfTrunc ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "PERCENTRANK", String.valueOf( xlfPercentrank ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "RIGHTB", String.valueOf( xlfRightb ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "REPLACEB", String.valueOf( xlfReplaceb ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FINDB", String.valueOf( xlfFindb ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MIDB", String.valueOf( xlfMidb ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ROWS", String.valueOf( xlfRows ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "COLUMNS", String.valueOf( xlfColumns ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "OFFSET", String.valueOf( xlfOffset ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ISTEXT", String.valueOf( xlfIstext ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LOOKUP", String.valueOf( xlfLookup ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "EXPONDIST", String.valueOf( xlfExpondist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FDIST", String.valueOf( xlfFdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FINV", String.valueOf( xlfFinv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FTEST", String.valueOf( xlfFtest ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FISHER", String.valueOf( xlfFisher ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FISHERINV", String.valueOf( xlfFisherinv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "STANDARDIZE", String.valueOf( xlfStandardize ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "PERMUT", String.valueOf( xlfPermut ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "POISSON", String.valueOf( xlfPoisson ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SUMXMY2", String.valueOf( xlfSumxmy2 ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SUMX2MY2", String.valueOf( xlfSumx2my2 ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SUMX2PY2", String.valueOf( xlfSumx2py2 ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ERFC", String.valueOf( xlfERFC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "CONFIDENCE", String.valueOf( xlfConfidence ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "CRITBINOM", String.valueOf( xlfCritbinom ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DEVSQ", String.valueOf( xlfDevsq ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SERIESSUM", String.valueOf( xlfSERIESSUM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "SUBTOTAL", String.valueOf( xlfSubtotal ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SUMSQ", String.valueOf( xlfSumsq ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "CHIDIST", String.valueOf( xlfChidist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "CHIINV", String.valueOf( xlfChiinv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "CHITEST", String.valueOf( xlfChitest ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "GAMMADIST", String.valueOf( xlfGammadist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "GAMMAINV", String.valueOf( xlfGammainv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "GAMMALN", String.valueOf( xlfGammaln ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "GEOMEAN", String.valueOf( xlfGeomean ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "GROWTH", String.valueOf( xlfGrowth ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "HARMEAN", String.valueOf( xlfHarmean ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "HYPGEOMDIST", String.valueOf( xlfHypgeomdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "KURT", String.valueOf( xlfKurt ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "LOGEST", String.valueOf( xlfLogest ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "LOGINV", String.valueOf( xlfLoginv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LOGNORMDIST", String.valueOf( xlfLognormdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NEGBINOMDIST", String.valueOf( xlfNegbinomdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "PROB", String.valueOf( xlfProb ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SKEW", String.valueOf( xlfSkew ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "STDEVPA", String.valueOf( xlfStDevPA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "STDEVP", String.valueOf( xlfStdevp ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "STDEVA", String.valueOf( xlfStDevA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TDIST", String.valueOf( xlfTdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "TINV", String.valueOf( xlfTinv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "TTEST", String.valueOf( xlfTtest ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "VARA", String.valueOf( xlfVarA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "VARPA", String.valueOf( xlfVarPA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "WEIBULL", String.valueOf( xlfWeibull ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ZTEST", String.valueOf( xlfZtest ), String.valueOf( FTYPE_PTGFUNCVAR ) },
	};

	// Contains function name, id and type of ALL Formulas (PtgFuncs, PtgFuncVars and Add-in PtgFuncVars) 

	// fetch the pattern match from: http://office.microsoft.com/client/helpcategory.aspx?CategoryID=CH100645029990&lcid=1033&NS=EXCEL&Version=12&CTT=4
	public static String[][] recArr = {
			{ "Pi", String.valueOf( xlfPi ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"Round", String.valueOf( xlfRound ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Rept", String.valueOf( xlfRept ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"Mid", String.valueOf( xlfMid ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Mod", String.valueOf( xlfMod ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"MMult", String.valueOf( xlfMmult ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Rand", String.valueOf( xlfRand ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"Date", String.valueOf( xlfDate ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Time", String.valueOf( xlfTime ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"Day", String.valueOf( xlfDay ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Now", String.valueOf( xlfNow ), String.valueOf( FTYPE_PTGFUNC ) },
			{
					"TAN", String.valueOf( xlfTan ), String.valueOf( FTYPE_PTGFUNC )
			},
			{ "Atan2", String.valueOf( xlfAtan2 ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Replace", String.valueOf( xlfReplace ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Exact", String.valueOf( xlfExact ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Trim", String.valueOf( xlfTrim ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Text", String.valueOf( xlfText ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Roundup", String.valueOf( xlfRoundup ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "RoundDown", String.valueOf( xlfRounddown ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "today", String.valueOf( xlfToday ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Combin", String.valueOf( xlfCombin ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Floor", String.valueOf( xlfFloor ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Ceiling", String.valueOf( xlfCeiling ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Power", String.valueOf( xlfPower ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Hour", String.valueOf( xlfHour ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Minute", String.valueOf( xlfMinute ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Month", String.valueOf( xlfMonth ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Year", String.valueOf( xlfYear ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Second", String.valueOf( xlfSecond ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Quartile", String.valueOf( xlfQuartile ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Frequency", String.valueOf( xlfFrequency ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Linest", String.valueOf( xlfLinest ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Correl", String.valueOf( xlfCorrel ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Slope", String.valueOf( xlfSlope ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Intercept", String.valueOf( xlfIntercept ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Pearson", String.valueOf( xlfPearson ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Rsq", String.valueOf( xlfRsq ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Steyx", String.valueOf( xlfSteyx ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Forecast", String.valueOf( xlfForecast ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "Covar", String.valueOf( xlfCovar ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "IsNumber", String.valueOf( xlfIsnumber ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DAVERAGE", String.valueOf( xlfDaverage ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DCOUNT", String.valueOf( xlfDcount ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DCOUNTA", String.valueOf( xlfDcounta ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DGET", String.valueOf( xlfDget ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DMIN", String.valueOf( xlfDmin ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DMAX", String.valueOf( xlfDmax ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DPRODUCT", String.valueOf( xlfDproduct ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DSTDEVP", String.valueOf( xlfDstdevp ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DSTDEV", String.valueOf( xlfDstdev ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DSUM", String.valueOf( xlfDsum ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DVAR", String.valueOf( xlfDvar ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DVARP", String.valueOf( xlfDvarp ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SQRT", String.valueOf( xlfSqrt ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NA", String.valueOf( xlfNa ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "EXP", String.valueOf( xlfExp ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "MIRR", String.valueOf( xlfMirr ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SLN", String.valueOf( xlfSln ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SYD", String.valueOf( xlfSyd ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISPMT", String.valueOf( xlfIspmt ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "UPPER", String.valueOf( xlfUpper ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LOWER", String.valueOf( xlfLower ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LEN", String.valueOf( xlfLen ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISLOGICAL", String.valueOf( xlfIslogical ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISERROR", String.valueOf( XLF_IS_ERROR ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISNONTEXT", String.valueOf( xlfIsnontext ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISBLANK", String.valueOf( xlfIsblank ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ISREF", String.valueOf( xlfIsref ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SIN", String.valueOf( xlfSin ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SINH", String.valueOf( xlfSinh ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ASIN", String.valueOf( xlfAsin ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ASINH", String.valueOf( xlfAsinh ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "COS", String.valueOf( xlfCos ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "COSH", String.valueOf( xlfCosh ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ACOS", String.valueOf( xlfAcos ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ACOSH", String.valueOf( xlfAcosh ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ATAN", String.valueOf( xlfAtan ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ATANH", String.valueOf( xlfAtanh ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "INT", String.valueOf( xlfInt ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ABS", String.valueOf( xlfAbs ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NOT", String.valueOf( xlfNot ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DEGREES", String.valueOf( xlfDegrees ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SIGN", String.valueOf( xlfSign ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "EVEN", String.valueOf( xlfEven ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ODD", String.valueOf( xlfOdd ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LN", String.valueOf( xlfLn ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FACT", String.valueOf( xlfFact ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "RADIANS", String.valueOf( xlfRadians ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "PROPER", String.valueOf( xlfProper ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "CHAR", String.valueOf( xlfChar ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "ERROR.TYPE", String.valueOf( xlfErrorType ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "T", String.valueOf( xlfT ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LOG10", String.valueOf( xlfLog10 ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "VALUE", String.valueOf( xlfValue ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "CODE", String.valueOf( xlfCode ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "N", String.valueOf( xlfN ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "DATEVALUE", String.valueOf( xlfDatevalue ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "SMALL", String.valueOf( xlfSmall ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LARGE", String.valueOf( xlfLarge ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NORMDIST", String.valueOf( xlfNormdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NORMSDIST", String.valueOf( xlfNormsdist ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NORMSINV", String.valueOf( xlfNormsinv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "NORMINV", String.valueOf( xlfNorminv ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LENB", String.valueOf( xlfLenb ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "INFO", String.valueOf( xlfInfo ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "LEFTB", String.valueOf( xlfLeftb ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "TRUE", String.valueOf( xlfTrue ), String.valueOf( FTYPE_PTGFUNC ) },
			{ "FALSE", String.valueOf( xlfFalse ), String.valueOf( FTYPE_PTGFUNC ) },
			// PtgFuncVars
			{ "COUNT", String.valueOf( XLF_COUNT ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "COUNTA", String.valueOf( xlfCounta ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "COUNTIF", String.valueOf( xlfCountif ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "COUNTBLANK", String.valueOf( xlfCountblank ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "IF", String.valueOf( XLF_IS ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ISNA", String.valueOf( XLF_IS_NA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ISERR", String.valueOf( xlfIserr ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SUM", String.valueOf( XLF_SUM ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SUMIF", String.valueOf( XLF_SUM_IF ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "AVERAGE", String.valueOf( XLF_AVERAGE ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MINA", String.valueOf( xlfMinA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MIN", String.valueOf( XLF_MIN ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MAXA", String.valueOf( xlfMaxA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MAX", String.valueOf( XLF_MAX ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ROW", String.valueOf( XLF_ROW ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "COLUMN", String.valueOf( xlfColumn ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "NPV", String.valueOf( xlfNpv ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "PMT", String.valueOf( xlfPmt ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "DB", String.valueOf( xlfDb ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "FIND", String.valueOf( xlfFind ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "DAYS360", String.valueOf( xlfDays360 ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "LEFT", String.valueOf( xlfLeft ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "LOG", String.valueOf( xlfLog ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MEDIAN", String.valueOf( xlfMedian ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MODE", String.valueOf( xlfMode ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "RANK", String.valueOf( xlfRank ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "RIGHT", String.valueOf( xlfRight ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "STDEV", String.valueOf( xlfStdev ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "VAR", String.valueOf( xlfVar ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "VARP", String.valueOf( xlfVarp ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TANH", String.valueOf( xlfTanh ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "VLOOKUP", String.valueOf( xlfVlookup ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "HLOOKUP", String.valueOf( xlfHlookup ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "CONCATENATE", String.valueOf( xlfConcatenate ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "INDEX", String.valueOf( xlfIndex ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "MATCH", String.valueOf( xlfMatch ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "FIXED", String.valueOf( xlfFixed ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "AND", String.valueOf( xlfAnd ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "OR", String.valueOf( xlfOr ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "CHOOSE", String.valueOf( xlfChoose ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ADDRESS", String.valueOf( xlfAddress ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "ROMAN", String.valueOf( xlfRoman ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "DOLLAR", String.valueOf( xlfDollar ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "USDOLLAR", String.valueOf( xlfUsdollar ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "AVEDEV", String.valueOf( xlfAvedev ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SUBSTITUTE", String.valueOf( xlfSubstitute ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "PRODUCT", String.valueOf( xlfProduct ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SEARCH", String.valueOf( xlfSearch ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "AVERAGEA", String.valueOf( xlfAverageA ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "TREND", String.valueOf( xlfTrend ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "SUMPRODUCT", String.valueOf( xlfSumproduct ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "INDIRECT", String.valueOf( XLF_INDIRECT ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			// Add-in Formulas
			// Financial Formulas
			{ "ACCRINTM", String.valueOf( xlfAccrintm ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ACCRINT", String.valueOf( xlfAccrint ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPDAYBS", String.valueOf( xlfCoupDayBS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPDAYS", String.valueOf( xlfCoupDays ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "PV", String.valueOf( xlfPv ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "FV", String.valueOf( xlfFv ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IPMT", String.valueOf( xlfIpmt ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "CUMIPMT", String.valueOf( xlfCumIPmt ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "CUMPRINC", String.valueOf( xlfCumPrinc ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPNCD", String.valueOf( xlfCoupNCD ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPDAYSNC", String.valueOf( xlfCoupDaysNC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPPCD", String.valueOf( xlfCoupPCD ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUPNUM", String.valueOf( xlfCoupNUM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DOLLARDE", String.valueOf( xlfDollarDE ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DOLLARFR", String.valueOf( xlfDollarFR ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "EFFECT", String.valueOf( xlfEffect ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "INTRATE", String.valueOf( xlfINTRATE ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IRR", String.valueOf( xlfIrr ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "XIRR", String.valueOf( xlfXIRR ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "XNPV", String.valueOf( xlfXNPV ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "RATE", String.valueOf( xlfRate ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "YIELD", String.valueOf( xlfYIELD ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "PRICE", String.valueOf( xlfPRICE ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "PRICEDISC", String.valueOf( xlfPRICEDISC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DISC", String.valueOf( xlfDISC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "PRICEMAT", String.valueOf( xlfPRICEMAT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DURATION", String.valueOf( xlfDURATION ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "MDURATION", String.valueOf( xlfMDURATION ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "NPER", String.valueOf( xlfNper ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "TBILLEQ", String.valueOf( xlfTBillEq ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "TBILLPRICE", String.valueOf( xlfTBillPrice ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "TBILLYIELD", String.valueOf( xlfTBillYield ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "YIELDDISC", String.valueOf( xlfYieldDisc ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "YIELDMAT", String.valueOf( xlfYieldMat ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "PPMT", String.valueOf( xlfPpmt ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "FVSCHEDULE", String.valueOf( xlfFVSchedule ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "AMORLINC", String.valueOf( xlfAmorlinc ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "AMORDEGRC", String.valueOf( xlfAmordegrc ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ODDFPRICE", String.valueOf( xlfOddFPrice ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ODDLPRICE", String.valueOf( xlfOddLPrice ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ODDFYIELD", String.valueOf( xlfOddFYield ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ODDLYIELD", String.valueOf( xlfOddLYield ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "NOMINAL", String.valueOf( xlfNOMINAL ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "VDB", String.valueOf( xlfVdb ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DDB", String.valueOf( xlfDdb ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "RECEIVED", String.valueOf( xlfRECEIVED ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Engineering Formulas
			{ "BIN2DEC", String.valueOf( xlfBIN2DEC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "BIN2HEX", String.valueOf( xlfBIN2HEX ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "BIN2OCT", String.valueOf( xlfBIN2OCT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DEC2BIN", String.valueOf( xlfDEC2BIN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DEC2HEX", String.valueOf( xlfDEC2HEX ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DEC2OCT", String.valueOf( xlfDEC2OCT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "HEX2BIN", String.valueOf( xlfHEX2BIN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "HEX2DEC", String.valueOf( xlfHEX2DEC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "HEX2OCT", String.valueOf( xlfHEX2OCT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "OCT2BIN", String.valueOf( xlfOCT2BIN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "OCT2DEC", String.valueOf( xlfOCT2DEC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "OCT2HEX", String.valueOf( xlfOCT2HEX ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COMPLEX", String.valueOf( xlfCOMPLEX ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "GESTEP", String.valueOf( xlfGESTEP ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "DELTA", String.valueOf( xlfDELTA ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMAGINARY", String.valueOf( xlfIMAGINARY ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMREAL", String.valueOf( xlfIMREAL ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMARGUMENT", String.valueOf( xlfIMARGUMENT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMABS", String.valueOf( xlfIMABS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMDIV", String.valueOf( xlfIMDIV ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMCONJUGATE", String.valueOf( xlfIMCONJUGATE ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMCOS", String.valueOf( xlfIMCOS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMSIN", String.valueOf( xlfIMSIN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMEXP", String.valueOf( xlfIMEXP ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMPOWER", String.valueOf( xlfIMPOWER ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMSQRT", String.valueOf( xlfIMSQRT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMSUB", String.valueOf( xlfIMSUB ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMSUM", String.valueOf( xlfIMSUM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMPRODUCT", String.valueOf( xlfIMPRODUCT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMLN", String.valueOf( xlfIMLN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMLOG10", String.valueOf( xlfIMLOG10 ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "IMLOG2", String.valueOf( xlfIMLOG2 ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "CONVERT", String.valueOf( xlfCONVERT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ERF", String.valueOf( xlfERF ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Math Add-In Formulas
			{ "FACTDOUBLE", String.valueOf( xlfDOUBLEFACT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "GCD", String.valueOf( xlfGCD ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "LCM", String.valueOf( xlfLCM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "MROUND", String.valueOf( xlfMROUND ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "MULTINOMIAL", String.valueOf( xlfMULTINOMIAL ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "QUOTIENT", String.valueOf( xlfQUOTIENT ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "RANDBETWEEN", String.valueOf( xlfRANDBETWEEN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "SERIESSUM", String.valueOf( xlfSERIESSUM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "SQRTPI", String.valueOf( xlfSQRTPI ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "SUMIFS", String.valueOf( xlfSUMIFS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Information Add-Ins
			{ "ISEVEN", String.valueOf( xlfISEVEN ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "ISODD", String.valueOf( xlfISODD ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Date/Time Add-in Formulas
			{ "NETWORKDAYS", String.valueOf( xlfNETWORKDAYS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "EDATE", String.valueOf( xlfEDATE ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "EOMONTH", String.valueOf( xlfEOMONTH ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "WEEKNUM", String.valueOf( xlfWEEKNUM ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "WEEKDAY", String.valueOf( xlfWeekday ), String.valueOf( FTYPE_PTGFUNCVAR ) },
			{ "WORKDAY", String.valueOf( xlfWORKDAY ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "YEARFRAC", String.valueOf( xlfYEARFRAC ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Statistical
			{ "AVERAGEIF", String.valueOf( xlfAVERAGEIF ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "AVERAGEIFS", String.valueOf( xlfAVERAGEIFS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			{ "COUNTIFS", String.valueOf( xlfCOUNTIFS ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
			// Logical
			{ "IFERROR", String.valueOf( xlfIFERROR ), String.valueOf( FTYPE_PTGFUNCVAR_ADDIN ) },
	};

	/**
	 * Handles differences
	 * in japanese locale xls
	 *
	 * @param iftb
	 * @return
	 */
	public static String getJFunctionString( short iftb )
	{
		switch( iftb )
		{
			case xlfDollar:
				return "YEN(";
			case xlfUsdollar:
				return "DOLLAR(";
			case xlfDbcs:
				return "JIS(";
		}
		return null;
	}

	public static String getFunctionString( short iftb )
	{
		switch( iftb )
		{
			case xlfADDIN:
				return "";
			case XLF_COUNT:
				return "COUNT(";
			case XLF_IS:
				return "IF(";
			case XLF_IS_NA:
				return "ISNA(";
			case XLF_IS_ERROR:
				return "ISERROR(";
			case XLF_SUM:
				return "SUM(";
			case XLF_AVERAGE:
				return "AVERAGE(";
			case XLF_MIN:
				return "MIN(";
			case XLF_MAX:
				return "MAX(";
			case XLF_ROW:
				return "ROW(";
			case xlfColumn:
				return "COLUMN(";
			case xlfNa:
				return "NA(";
			case xlfNpv:
				return "NPV(";
			case xlfStdev:
				return "STDEV(";
			case xlfDollar:
				return "DOLLAR(";
			case xlfFixed:
				return "FIXED(";
			case xlfSin:
				return "SIN(";
			case xlfCos:
				return "COS(";
			case xlfTan:
				return "TAN(";
			case xlfAtan:
				return "ATAN(";
			case xlfPi:
				return "PI(";
			case xlfSqrt:
				return "SQRT(";
			case xlfExp:
				return "EXP(";
			case xlfLn:
				return "LN(";
			case xlfLog10:
				return "LOG10(";
			case xlfAbs:
				return "ABS(";
			case xlfInt:
				return "INT(";
			case xlfSign:
				return "SIGN(";
			case xlfRound:
				return "ROUND(";
			case xlfLookup:
				return "LOOKUP(";
			case xlfIndex:
				return "INDEX(";
			case xlfRept:
				return "REPT(";
			case xlfMid:
				return "MID(";
			case xlfLen:
				return "LEN(";
			case xlfValue:
				return "VALUE(";
			case xlfTrue:
				return "TRUE(";
			case xlfFalse:
				return "FALSE(";
			case xlfAnd:
				return "AND(";
			case xlfOr:
				return "OR(";
			case xlfNot:
				return "NOT(";
			case xlfMod:
				return "MOD(";
			case xlfDaverage:
				return "DAVERAGE(";
			case xlfDcount:
				return "DCOUNT(";
			case xlfDcounta:
				return "DCOUNTA(";
			case xlfDget:
				return "DGET(";
			case xlfDmax:
				return "DMAX(";
			case xlfDmin:
				return "DMIN(";
			case xlfDproduct:
				return "DPRODUCT(";
			case xlfDstdev:
				return "DSTDEV(";
			case xlfDstdevp:
				return "DSTDEVP(";
			case xlfDsum:
				return "DSUM(";
			case xlfDvar:
				return "DVAR(";
			case xlfDvarp:
				return "DVARP(";
			case xlfVar:
				return "VAR(";
			case xlfText:
				return "TEXT(";
			case xlfLinest:
				return "LINEST(";
			case xlfTrend:
				return "TREND(";
			case xlfLogest:
				return "LOGEST(";
			case xlfGrowth:
				return "GROWTH(";
			case xlfGoto:
				return "GOTO(";
			case xlfHalt:
				return "HALT(";
			case xlfPv:
				return "PV(";
			case xlfFv:
				return "FV(";
			case xlfNper:
				return "NPER(";
			case xlfPmt:
				return "PMT(";
			case xlfRate:
				return "RATE(";
			case xlfMirr:
				return "MIRR(";
			case xlfIrr:
				return "IRR(";
			case xlfRand:
				return "RAND(";
			case xlfMatch:
				return "MATCH(";
			case xlfDate:
				return "DATE(";
			case xlfTime:
				return "TIME(";
			case xlfDay:
				return "DAY(";
			case xlfMonth:
				return "MONTH(";
			case xlfYear:
				return "YEAR(";
			case xlfWeekday:
				return "WEEKDAY(";
			case xlfHour:
				return "HOUR(";
			case xlfMinute:
				return "MINUTE(";
			case xlfSecond:
				return "SECOND(";
			case xlfNow:
				return "NOW(";
			case xlfAreas:
				return "AREAS(";
			case xlfRows:
				return "ROWS(";
			case xlfColumns:
				return "COLUMNS(";
			case xlfOffset:
				return "OFFSET(";
			case xlfAbsref:
				return "ABSREF(";
			case xlfRelref:
				return "RELREF(";
			case xlfArgument:
				return "ARGUMENT(";
			case xlfSearch:
				return "SEARCH(";
			case xlfTranspose:
				return "TRANSPOSE(";
			case xlfError:
				return "ERROR(";
			case xlfStep:
				return "STEP(";
			case xlfType:
				return "TYPE(";
			case xlfEcho:
				return "ECHO(";
			case xlfSetName:
				return "SETNAME(";
			case xlfCaller:
				return "CALLER(";
			case xlfDeref:
				return "DEREF(";
			case xlfWindows:
				return "WINDOWS(";
			case xlfSeries:
				return "SERIES(";
			case xlfDocuments:
				return "DOCUMENTS(";
			case xlfActiveCell:
				return "ACTIVECELL(";
			case xlfSelection:
				return "SELECTION(";
			case xlfResult:
				return "RESULT(";
			case xlfAtan2:
				return "ATAN2(";
			case xlfAsin:
				return "ASIN(";
			case xlfAcos:
				return "ACOS(";
			case xlfChoose:
				return "CHOOSE(";
			case xlfHlookup:
				return "HLOOKUP(";
			case xlfVlookup:
				return "VLOOKUP(";
			case xlfLinks:
				return "LINKS(";
			case xlfInput:
				return "INPUT(";
			case xlfIsref:
				return "ISREF(";
			case xlfGetFormula:
				return "GETFORMULA(";
			case xlfGetName:
				return "GETNAME(";
			case xlfSetValue:
				return "SETVALUE(";
			case xlfLog:
				return "LOG(";
			case xlfExec:
				return "EXEC(";
			case xlfChar:
				return "CHAR(";
			case xlfLower:
				return "LOWER(";
			case xlfUpper:
				return "UPPER(";
			case xlfProper:
				return "PROPER(";
			case xlfLeft:
				return "LEFT(";
			case xlfRight:
				return "RIGHT(";
			case xlfExact:
				return "EXACT(";
			case xlfTrim:
				return "TRIM(";
			case xlfReplace:
				return "REPLACE(";
			case xlfSubstitute:
				return "SUBSTITUTE(";
			case xlfCode:
				return "CODE(";
			case xlfNames:
				return "NAMES(";
			case xlfDirectory:
				return "DIRECTORY(";
			case xlfFind:
				return "FIND(";
			case xlfCell:
				return "CELL(";
			case xlfIserr:
				return "ISERR(";
			case xlfIstext:
				return "ISTEXT(";
			case xlfIsnumber:
				return "ISNUMBER(";
			case xlfIsblank:
				return "ISBLANK(";
			case xlfT:
				return "T(";
			case xlfN:
				return "N(";
			case xlfFopen:
				return "FOPEN(";
			case xlfFclose:
				return "FCLOSE(";
			case xlfFsize:
				return "SIZE(";
			case xlfFreadln:
				return "FREADLN(";
			case xlfFread:
				return "FREAD(";
			case xlfFwriteln:
				return "FWRITELN(";
			case xlfFwrite:
				return "FWRITE(";
			case xlfFpos:
				return "FPOS(";
			case xlfDatevalue:
				return "DATEVALUE(";
			case xlfTimevalue:
				return "TIMEVALUE(";
			case xlfSln:
				return "SLN(";
			case xlfSyd:
				return "SYD(";
			case xlfDdb:
				return "DDB(";
			case xlfGetDef:
				return "GETDEF(";
			case xlfReftext:
				return "REFTEXT(";
			case xlfTextref:
				return "TEXTREF(";
			case XLF_INDIRECT:
				return "INDIRECT(";
			case xlfRegister:
				return "REGISTER(";
			case xlfCall:
				return "CALL(";
			case xlfAddBar:
				return "ADDBAR(";
			case xlfAddMenu:
				return "ADDMENU(";
			case xlfAddCommand:
				return "ADDCOMMAND(";
			case xlfEnableCommand:
				return "ENABLECOMMAND(";
			case xlfCheckCommand:
				return "CHECKCOMMAND(";
			case xlfRenameCommand:
				return "RENAMECOMMAND(";
			case xlfShowBar:
				return "SHOWBAR(";
			case xlfDeleteMenu:
				return "DELETEMENU(";
			case xlfDeleteCommand:
				return "DELETECOMMAND(";
			case xlfGetChartItem:
				return "CHARTITEM(";
			case xlfDialogBox:
				return "DIALOGBOX(";
			case xlfClean:
				return "CLEAN(";
			case xlfMdeterm:
				return "MDETERM(";
			case xlfMinverse:
				return "MINVERSE(";
			case xlfMmult:
				return "MMULT(";
			case xlfFiles:
				return "FILES(";
			case xlfIpmt:
				return "IPMT(";
			case xlfPpmt:
				return "PPMT(";
			case xlfCounta:
				return "COUNTA(";
			case xlfCancelKey:
				return "CANCELKEY(";
			case xlfInitiate:
				return "INITIATE(";
			case xlfRequest:
				return "REQUEST(";
			case xlfPoke:
				return "POKE(";
			case xlfExecute:
				return "EXECUTE(";
			case xlfTerminate:
				return "TERMINATE(";
			case xlfRestart:
				return "RESTART(";
			case xlfHelp:
				return "HELP(";
			case xlfGetBar:
				return "GETBAR(";
			case xlfProduct:
				return "PRODUCT(";
			case xlfFact:
				return "FACT(";
			case xlfGetCell:
				return "GETCELL(";
			case xlfGetWorkspace:
				return "GETWORKSPACE(";
			case xlfGetWindow:
				return "GETWINDOW(";
			case xlfGetDocument:
				return "GETDOCUMENT(";
			case xlfIsnontext:
				return "ISNONTEXT(";
			case xlfGetNote:
				return "GETNOTE(";
			case xlfNote:
				return "NOTE(";
			case xlfStdevp:
				return "STDEVP(";
			case xlfVarp:
				return "VARP(";
			case xlfTrunc:
				return "TRUNC(";
			case xlfIslogical:
				return "ISLOGICAL(";
			case xlfDeleteBar:
				return "DELETEBAR(";
			case xlfUnregister:
				return "UNREGISTER(";
			case xlfUsdollar:
				return "USDOLLAR(";
			case xlfFindb:
				return "FINDB(";
			case xlfSearchb:
				return "SEARCHB(";
			case xlfReplaceb:
				return "REPLACEB(";
			case xlfLeftb:
				return "LEFTB(";
			case xlfRightb:
				return "RIGHTB(";
			case xlfMidb:
				return "MIDB(";
			case xlfLenb:
				return "LENB(";
			case xlfRoundup:
				return "ROUNDUP(";
			case xlfRounddown:
				return "ROUNDDOWN(";
			case xlfAsc:
				return "ASC(";
			case xlfDbcs:
				return "DBCS(";
			case xlfRank:
				return "RANK(";
			case xlfAddress:
				return "ADDRESS(";
			case xlfDays360:
				return "DAYS360(";
			case xlfToday:
				return "TODAY(";
			case xlfVdb:
				return "VDB(";
			case xlfMedian:
				return "MEDIAN(";
			case xlfSumproduct:
				return "SUMPRODUCT(";
			case xlfSinh:
				return "SINH(";
			case xlfCosh:
				return "COSH(";
			case xlfTanh:
				return "TANH(";
			case xlfAsinh:
				return "ASINH(";
			case xlfAcosh:
				return "ACOSH(";
			case xlfAtanh:
				return "ATANH(";
			case xlfCreateObject:
				return "CREATEOBJECT(";
			case xlfVolatile:
				return "VOLATILE(";
			case xlfLastError:
				return "LASTERROR(";
			case xlfCustomUndo:
				return "CUSTOMUNDO(";
			case xlfCustomRepeat:
				return "CUSTOMREPEAT(";
			case xlfFormulaConvert:
				return "FORMULACONVERT(";
			case xlfGetLinkInfo:
				return "GETLINKINFO(";
			case xlfTextBox:
				return "TEXTBOX(";
			case xlfInfo:
				return "INFO(";
			case xlfGroup:
				return "GROUP(";
			case xlfGetObject:
				return "GETOBJECT(";
			case xlfDb:
				return "DB(";
			case xlfPause:
				return "PAUSE(";
			case xlfResume:
				return "RESUME(";
			case xlfFrequency:
				return "FREQUENCY(";
			case xlfAddToolbar:
				return "ADDTOOLBAR(";
			case xlfDeleteToolbar:
				return "DELETETOOLBAR(";
			case xlfResetToolbar:
				return "RESETTOOLBAR(";
			case xlfEvaluate:
				return "EVALUATE(";
			case xlfGetToolbar:
				return "GETTOOLBAR(";
			case xlfGetTool:
				return "GETTOOL(";
			case xlfSpellingCheck:
				return "SPELLINGCHECK(";
			case xlfErrorType:
				return "ERROR.TYPE(";
			case xlfAppTitle:
				return "APPTITLE(";
			case xlfWindowTitle:
				return "WINDOWTITLE(";
			case xlfSaveToolbar:
				return "SAVETOOLBAR(";
			case xlfEnableTool:
				return "ENABLETOOL(";
			case xlfPressTool:
				return "PRESSTOOL(";
			case xlfRegisterId:
				return "REGISTERID(";
			case xlfGetWorkbook:
				return "GETWORKBOOK(";
			case xlfAvedev:
				return "AVEDEV(";
			case xlfBetadist:
				return "BETADIST(";
			case xlfGammaln:
				return "GAMMALN(";
			case xlfBetainv:
				return "BETAINV(";
			case xlfBinomdist:
				return "BINOMDIST(";
			case xlfChidist:
				return "CHIDIST(";
			case xlfChiinv:
				return "CHIINV(";
			case xlfCombin:
				return "COMBIN(";
			case xlfConfidence:
				return "CONFIDENCE(";
			case xlfCritbinom:
				return "CRITBINOM(";
			case xlfEven:
				return "EVEN(";
			case xlfExpondist:
				return "EXPONDIST(";
			case xlfFdist:
				return "FDIST(";
			case xlfFinv:
				return "FINV(";
			case xlfFisher:
				return "FISHER(";
			case xlfFisherinv:
				return "FISHERINV(";
			case xlfFloor:
				return "FLOOR(";
			case xlfGammadist:
				return "GAMMADIST(";
			case xlfGammainv:
				return "GAMMAINV(";
			case xlfCeiling:
				return "CEILING(";
			case xlfHypgeomdist:
				return "HYPGEOMDIST(";
			case xlfLognormdist:
				return "LOGNORMDIST(";
			case xlfLoginv:
				return "LOGINV(";
			case xlfNegbinomdist:
				return "NEGBINOMDIST(";
			case xlfNormdist:
				return "NORMDIST(";
			case xlfNormsdist:
				return "NORMSDIST(";
			case xlfNorminv:
				return "NORMINV(";
			case xlfNormsinv:
				return "NORMSINV(";
			case xlfStandardize:
				return "STANDARDIZE(";
			case xlfOdd:
				return "ODD(";
			case xlfPermut:
				return "PERMUT(";
			case xlfPoisson:
				return "POISSON(";
			case xlfTdist:
				return "TDIST(";
			case xlfWeibull:
				return "WEIBULL(";
			case xlfSumxmy2:
				return "SUMXMY2(";
			case xlfSumx2my2:
				return "SUMX2MY2(";
			case xlfSumx2py2:
				return "SUMX2PY2(";
			case xlfChitest:
				return "CHITEST(";
			case xlfCorrel:
				return "CORREL(";
			case xlfCovar:
				return "COVAR(";
			case xlfForecast:
				return "FORECAST(";
			case xlfFtest:
				return "FTEST(";
			case xlfIntercept:
				return "INTERCEPT(";
			case xlfPearson:
				return "PEARSON(";
			case xlfRsq:
				return "RSQ(";
			case xlfSteyx:
				return "STEYX(";
			case xlfSlope:
				return "SLOPE(";
			case xlfTtest:
				return "TTEST(";
			case xlfProb:
				return "PROB(";
			case xlfDevsq:
				return "DEVSQ(";
			case xlfGeomean:
				return "GEOMEAN(";
			case xlfHarmean:
				return "HARMEAN(";
			case xlfSumsq:
				return "SUMSQ(";
			case xlfKurt:
				return "KURT(";
			case xlfSkew:
				return "SKEW(";
			case xlfZtest:
				return "ZTEST(";
			case xlfLarge:
				return "LARGE(";
			case xlfSmall:
				return "SMALL(";
			case xlfQuartile:
				return "QUARTILE(";
			case xlfPercentile:
				return "PERCENTILE(";
			case xlfPercentrank:
				return "PERCENTRANK(";
			case xlfMode:
				return "MODE(";
			case xlfTrimmean:
				return "TRIMMEAN(";
			case xlfTinv:
				return "TINV(";
			case xlfMovieCommand:
				return "MOVIECOMMAND(";
			case xlfGetMovie:
				return "GETMOVIE(";
			case xlfConcatenate:
				return "CONCATENATE(";
			case xlfPower:
				return "POWER(";
			case xlfPivotAddData:
				return "PIVOTADDDATA(";
			case xlfGetPivotTable:
				return "GETPIVOTTABLE(";
			case xlfGetPivotField:
				return "GETPIVOTFIELD(";
			case xlfGetPivotItem:
				return "GETPIVOTITEM(";
			case xlfRadians:
				return "RADIANS(";
			case xlfDegrees:
				return "DEGREES(";
			case xlfSubtotal:
				return "SUBTOTAL(";
			case XLF_SUM_IF:
				return "SUMIF(";
			case xlfCountif:
				return "COUNTIF(";
			case xlfCountblank:
				return "COUNTBLANK(";
			case xlfScenarioGet:
				return "SCENARIOGET(";
			case xlfOptionsListsGet:
				return "OPTIONSLISTSGET(";
			case xlfIspmt:
				return "ISPMT(";
			case xlfDatedif:
				return "DATEDIF(";
			case xlfDatestring:
				return "DATESTRING(";
			case xlfNumberstring:
				return "NUMBERSTRING(";
			case xlfRoman:
				return "ROMAN(";
			case xlfOpenDialog:
				return "OPENDIALOG(";
			case xlfSaveDialog:
				return "SAVEDIALOG(";
			case xlfViewGet:
				return "VIEWGET(";
			case xlfGetPivotData:
				return "GETPIVOTDATA(";
			case xlfHyperlink:
				return "HYPERLINK(";
			case xlfPhonetic:
				return "PHONETIC(";
			case xlfAverageA:
				return "AVERAGEA(";
			case xlfMaxA:
				return "MAXA(";
			case xlfMinA:
				return "MINA(";
			case xlfStDevPA:
				return "STDEVPA(";
			case xlfVarPA:
				return "VARPA(";
			case xlfStDevA:
				return "STDEVA(";
			case xlfVarA:
				return "VARA(";
			// ADD-IN FORMULAS
			// Financial Formulas AddIns
			case xlfAccrintm:
				return "ACCRINTM(";
			case xlfAccrint:
				return "ACCRINT(";
			case xlfCoupDayBS:
				return "COUPDAYBS(";
			case xlfCoupDays:
				return "COUPDAYS(";
			case xlfCoupDaysNC:
				return "COUPDAYSNC(";
			case xlfCumIPmt:
				return "CUMIPMT(";
			case xlfCumPrinc:
				return "CUMPRINC(";
			case xlfCoupNCD:
				return "COUPNCD(";
			case xlfCoupPCD:
				return "COUPPCD(";
			case xlfCoupNUM:
				return "COUPNUM(";
			case xlfDollarDE:
				return "DOLLARDE(";
			case xlfDollarFR:
				return "DOLLARFR(";
			case xlfEffect:
				return "EFFECT(";
			case xlfINTRATE:
				return "INTRATE(";
			case xlfXIRR:
				return "XIRR(";
			case xlfXNPV:
				return "XNPV(";
			case xlfYIELD:
				return "YIELD(";
			case xlfPRICE:
				return "PRICE(";
			case xlfPRICEDISC:
				return "PRICEDISC(";
			case xlfDISC:
				return "DISC(";
			case xlfPRICEMAT:
				return "PRICEMAT(";
			case xlfDURATION:
				return "DURATION(";
			case xlfMDURATION:
				return "MDURATION(";
			case xlfTBillEq:
				return "TBILLEQ(";
			case xlfTBillPrice:
				return "TBILLPRICE(";
			case xlfTBillYield:
				return "TBILLYIELD(";
			case xlfYieldDisc:
				return "YIELDDISC(";
			case xlfYieldMat:
				return "YIELDMAT(";
			case xlfFVSchedule:
				return "FVSCHEDULE(";
			case xlfAmorlinc:
				return "AMORLINC(";
			case xlfAmordegrc:
				return "AMORDEGRC(";
			case xlfOddFPrice:
				return "ODDFPRICE(";
			case xlfOddFYield:
				return "ODDFYIELD(";
			case xlfOddLPrice:
				return "ODDLPRICE(";
			case xlfOddLYield:
				return "ODDLYIELD(";
			case xlfNOMINAL:
				return "NOMINAL(";
			case xlfRECEIVED:
				return "RECEIVED(";
			// Engineering Formulas AddIns
			case xlfBIN2DEC:
				return "BIN2DEC(";
			case xlfBIN2HEX:
				return "BIN2HEX(";
			case xlfBIN2OCT:
				return "BIN2OCT(";
			case xlfDEC2BIN:
				return "DEC2BIN(";
			case xlfDEC2HEX:
				return "DEC2HEX(";
			case xlfDEC2OCT:
				return "DEC2OCT(";
			case xlfHEX2BIN:
				return "HEX2BIN(";
			case xlfHEX2DEC:
				return "HEX2DEC(";
			case xlfHEX2OCT:
				return "HEX2OCT(";
			case xlfOCT2BIN:
				return "OCT2BIN(";
			case xlfOCT2DEC:
				return "OCT2DEC(";
			case xlfOCT2HEX:
				return "OCT2HEX(";
			case xlfCOMPLEX:
				return "COMPLEX(";
			case xlfGESTEP:
				return "GESTEP(";
			case xlfDELTA:
				return "DELTA(";
			case xlfIMAGINARY:
				return "IMAGINARY(";
			case xlfIMREAL:
				return "IMREAL(";
			case xlfIMARGUMENT:
				return "IMARGUMENT(";
			case xlfIMABS:
				return "IMABS(";
			case xlfIMDIV:
				return "IMDIV(";
			case xlfIMCONJUGATE:
				return "IMCONJUGATE(";
			case xlfIMCOS:
				return "IMCOS(";
			case xlfIMSIN:
				return "IMSIN(";
			case xlfIMEXP:
				return "IMEXP(";
			case xlfIMPOWER:
				return "IMPOWER(";
			case xlfIMSQRT:
				return "IMSQRT(";
			case xlfIMSUB:
				return "IMSUB(";
			case xlfIMSUM:
				return "IMSUM(";
			case xlfIMPRODUCT:
				return "IMPRODUCT(";
			case xlfIMLN:
				return "IMLN(";
			case xlfIMLOG10:
				return "IMLOG10(";
			case xlfIMLOG2:
				return "IMLOG2(";
			case xlfCONVERT:
				return "CONVERT(";
			case xlfDOUBLEFACT:
				return "FACTDOUBLE(";
			case xlfGCD:
				return "GCD(";
			case xlfLCM:
				return "LCM(";
			case xlfMROUND:
				return "MROUND(";
			case xlfMULTINOMIAL:
				return "MULTINOMIAL(";
			case xlfQUOTIENT:
				return "QUOTIENT(";
			case xlfRANDBETWEEN:
				return "RANDBETWEEN(";
			case xlfSERIESSUM:
				return "SERIESSUM(";
			case xlfSQRTPI:
				return "SQRTPI(";
			case xlfERF:
				return "ERF(";
			// information Add-ins
			case xlfISEVEN:
				return "ISEVEN(";
			case xlfISODD:
				return "ISODD(";
			// Date/Time Add-in Formulas
			case xlfNETWORKDAYS:
				return "NETWORKDAYS(";
			case xlfEDATE:
				return "EDATE(";
			case xlfEOMONTH:
				return "EOMONTH(";
			case xlfWEEKNUM:
				return "WEEKNUM(";
			case xlfWORKDAY:
				return "WORKDAY(";
			case xlfYEARFRAC:
				return "YEARFRAC(";
			// Statistical
			case xlfAVERAGEIF:
				return "AVERAGEIF(";
			case xlfAVERAGEIFS:
				return "AVERAGEIFS(";
			case xlfCOUNTIFS:
				return "COUNTIFS(";
			// Logical
			case xlfIFERROR:
				return "IFERROR(";
			// Math
			case xlfSUMIFS:
				return "SUMIFS(";
		}
		return "";
	}

	// Num Params for all PTGFUNCs  - ptgfuncvar's have variable # args (hence the name ...) 
	public static int getNumParams( int iftab )
	{
		if( iftab == xlfNa )
		{
			return 0; // na
		}
		if( iftab == xlfPi )
		{
			return 0; // Pi
		}
		if( iftab == xlfRound )
		{
			return 2; // Round
		}
		if( iftab == xlfRept )
		{
			return 2; // rept
		}
		if( iftab == xlfMid )
		{
			return 3; // Mid
		}
		if( iftab == xlfMod )
		{
			return 2; // Mod
		}
		if( (iftab >= xlfDcount) && (iftab <= xlfDstdev) )
		{
			return 3; // Dxxx formulas
		}
		if( iftab == xlfDvar )
		{
			return 3; // DVar
		}
		if( iftab == xlfRand )
		{
			return 0; // Rand
		}
		if( iftab == xlfDate )
		{
			return 3; // Date
		}
		if( iftab == xlfTime )
		{
			return 3; // Time
		}
		if( iftab == xlfDay )
		{
			return 1; // Day
		}
		if( iftab == xlfNow )
		{
			return 0; // now
		}
		if( iftab == xlfAtan2 )
		{
			return 2; // Atan2
		}
		if( iftab == xlfLog )
		{
			return 2; // Log
		}
		if( iftab == xlfLeft )
		{
			return 2; // Left
		}
		if( iftab == xlfRight )
		{
			return 2; // Right
		}
		if( iftab == xlfTrim )
		{
			return 1; // Trim
		}
		if( iftab == xlfText )
		{
			return 2; // Text
		}
		if( iftab == xlfReplace )
		{
			return 4; // Replace
		}
		if( iftab == xlfExact )
		{
			return 2; // Exact
		}
		if( iftab == 165 )
		{
			return 2;  //TODO:
		}
		if( iftab == xlfDproduct )
		{
			return 3; // DProduct
		}
		if( iftab == xlfDstdevp )
		{
			return 3; // DStdDevp
		}
		if( iftab == xlfDvarp )
		{
			return 3; // DVarP
		}
		if( iftab == xlfDcounta )
		{
			return 3; // DCountA
		}
		if( iftab == xlfRoundup )
		{
			return 2; // Roundup
		}
		if( iftab == xlfRounddown )
		{
			return 2; // Rounddown
		}
		if( iftab == xlfToday )
		{
			return 0; // today
		}
		if( iftab == xlfDget )
		{
			return 3; // DGet
		}
		if( iftab == xlfCombin )
		{
			return 2; // Combin
		}
		if( iftab == xlfFloor )
		{
			return 2; // Floor
		}
		if( iftab == xlfCeiling )
		{
			return 2; // Ceiling
		}
		if( iftab == xlfPower )
		{
			return 2; // Power
		}
		if( iftab == xlfCountif )
		{
			return 2; // CountIf
		}
		if( iftab == xlfQuartile )
		{
			return 2;
		}
		if( iftab == xlfFrequency )
		{
			return 2;
		}
		if( iftab == xlfCorrel )
		{
			return 2;
		}
		if( iftab == xlfCovar )
		{
			return 2;
		}
		if( iftab == xlfSlope )
		{
			return 2;
		}
		if( iftab == xlfIntercept )
		{
			return 2;
		}
		if( iftab == xlfPearson )
		{
			return 2;
		}
		if( iftab == xlfRsq )
		{
			return 2;
		}
		if( iftab == xlfSteyx )
		{
			return 2;
		}
		if( iftab == xlfCritbinom )
		{
			return 3;
		}
		if( iftab == xlfForecast )
		{
			return 3;
		}
		if( iftab == xlfTrend )
		{
			return 2;
		}
		if( iftab == xlfIsnumber )
		{
			return 1;
		}
		if( iftab == xlfMmult )
		{
			return 2;
		}
		if( iftab == xlfHour )
		{
			return 1;
		}
		if( iftab == xlfMinute )
		{
			return 1;
		}
		if( iftab == xlfMonth )
		{
			return 1;
		}
		if( iftab == xlfYear )
		{
			return 1;
		}
		if( iftab == xlfSecond )
		{
			return 1;
		}
		if( iftab == xlfSqrt )
		{
			return 1;
		}
		if( iftab == xlfExp )
		{
			return 1;
		}
		if( iftab == xlfMirr )
		{
			return 3;
		}
		if( iftab == xlfSyd )
		{
			return 4;
		}
		if( iftab == xlfSln )
		{
			return 3;
		}
		if( iftab == xlfIspmt )
		{
			return 4;
		}
		if( iftab == xlfBinomdist )
		{
			return 4;
		}
		if( iftab == xlfChidist )
		{
			return 2;
		}
		if( iftab == xlfChiinv )
		{
			return 2;
		}
		if( iftab == xlfChitest )
		{
			return 2;
		}
		if( iftab == xlfConfidence )
		{
			return 3;
		}
		if( iftab == xlfFtest )
		{
			return 2;
		}
		if( iftab == xlfSumx2my2 )
		{
			return 2;
		}
		if( iftab == xlfSumx2py2 )
		{
			return 2;
		}
		if( iftab == xlfSumxmy2 )
		{
			return 2;
		}
		if( iftab == xlfLookup )
		{
			return 3;
		}
		if( iftab == xlfTrue )
		{
			return 0;
		}
		if( iftab == xlfFalse )
		{
			return 0;
		}
		if( iftab == xlfExpondist )
		{
			return 3;
		}
		if( iftab == xlfFdist )
		{
			return 3;
		}
		if( iftab == xlfFinv )
		{
			return 3;
		}
		if( iftab == xlfLoginv )
		{
			return 2;
		}
		if( iftab == xlfNegbinomdist )
		{
			return 3;
		}
		if( iftab == xlfNormdist )
		{
			return 4;
		}
		if( iftab == xlfNorminv )
		{
			return 3;
		}
		if( iftab == xlfNormsinv )
		{
			return 1;
		}
		if( iftab == xlfStandardize )
		{
			return 3;
		}
		if( iftab == xlfPermut )
		{
			return 2;
		}
		if( iftab == xlfPoisson )
		{
			return 3;
		}
		if( iftab == xlfSumx2my2 )
		{
			return 2;
		}
		if( iftab == xlfSumx2py2 )
		{
			return 2;
		}
		if( iftab == xlfSumxmy2 )
		{
			return 2;
		}
		if( iftab == xlfTdist )
		{
			return 2;
		}
		if( iftab == xlfLarge )
		{
			return 2;
		}
		if( iftab == xlfSmall )
		{
			return 2;
		}
		return 1; //if we are lucky - rest all should be 1 param!
	}
}
