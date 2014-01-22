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
package com.extentech.naming;

import javax.naming.Name;
import javax.naming.NameParser;
import javax.naming.NamingException;

/* 
  This class is used for parsing names from a hierarchical
  namespace. The NameParser contains knowledge of the syntactic 
  information (like left-to-right orientation, name separator, etc.) 
  needed to parse names. 
  
  The equals() method, when used to compare two NameParsers, returns 
  true if and only if they serve the same namespace. 

	
	@see:CompoundName, Name
*/
class NameParserImpl implements NameParser
{

	/* Parses a name into its components.
		Parameters:
		name - The non-null string name to parse.
		Returns:
		A non-null parsed form of the name using the naming convention of this parser.
		Throws:
		InvalidNameException - If name does not conform to syntax defined for the namespace.
		NamingException - If a naming exception was encountered.
	 * @see javax.naming.NameParser#parse(java.lang.String)
	 */
	public Name parse( String arg0 ) throws NamingException
	{
		Name nm = new NameImpl();

		nm.add( arg0 ); // just plop it in for now...

		return nm;
	}

}