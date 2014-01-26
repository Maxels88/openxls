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

/**
 * WorkBookInstantiationException is thrown when a workbook cannot be parsed for a particular reason.
 * <p/>
 * Error codes can be retrieved with getErrorCode, which map to the static error ints
 */
public class WorkBookException extends org.openxls.formats.XLS.WorkBookException
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5313787084750169461L;
	public final static int DOUBLE_STREAM_FILE = 0;
	public final static int NOT_BIFF8_FILE = 1;
	public final static int LICENSING_FAILED = 2;
	public final static int UNSPECIFIED_INIT_ERROR = 3;
	public final static int RUNTIME_ERROR = 4;
	public final static int SMALLBLOCK_FILE = 5;
	public final static int WRITING_ERROR = 6;
	public final static int DECRYPTION_ERROR = 7;
	public final static int DECRYPTION_INCORRECT_PASSWORD = 8;
	public static final int ENCRYPTION_ERROR = 9;
	public final static int DECRYPTION_INCORRECT_FORMAT = 10;
	public final static int ILLEGAL_INIT_ERROR = 11;
	public final static int READ_ONLY_EXCEPTION = 12;
	public final static int SHEETPROTECT_INCORRECT_PASSWORD = 13;

	public WorkBookException( String n, int x )
	{
		super( n, x );
	}

	public WorkBookException( String string, int x, Exception e )
	{
		super( string, x, e );
	}

}
