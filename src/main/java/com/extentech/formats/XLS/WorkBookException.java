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
 */
public class WorkBookException extends RuntimeException
{
	private static final long serialVersionUID = 6406202417397276014L;

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

	private final int error_code;

	public WorkBookException( String message, int code )
	{
		super( message );
		error_code = code;
	}

	public WorkBookException( String message, int code, Throwable cause )
	{
		super( message, cause );
		error_code = code;
	}

	public int getErrorCode()
	{
		return error_code;
	}

	public String toString()
	{
		return "WorkBook initialization failed: '" + getMessage() + "'";
	}

	/**
	 * Obsolete synonym for <code>getCause()</code>.
	 *
	 * @deprecated Use {@link Throwable#getCause()} instead.
	 */
	public Exception getWrappedException()
	{
		Throwable cause = getCause();
		return cause instanceof Exception ? (Exception) cause : null;
	}
}
