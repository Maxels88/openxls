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
package com.extentech.formats.XLS.formulas;

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;

/*   ARRAY CONSTANT followed by 7 reserved bytes.
 * 
 * The token value for ptgArray consists of the array dimensions and the array values
 * 
 * ptgArray differs from most other operand tokens in that the token value doesn't follow the token type.
 * 
 * Instead, the token value is appended to the saved parsed expression, immediately following the last token.
 * 
 * Offset		Name		Size		Contents
    -----------------------------------------------------------
    0				nc		1			number of columns -1 in array constant (0 = 256)
    1				nr		2			number of rows -1  in array constant
    3				rgval		var		the array vals (k+1)*(nr+1) length
 * 
 * 
 * The format of the token value is shown in the following table.
 * 
 * The number of values in the array constant is equal to the product of the array dimensions, (nc+1)*(nr+1_
 * 
 * Each value is either an 8-byte IEEE fp numbr or a string.  The two formats for these values are shown in the following tables.
 * 
 * 
 * IEEE FP Number
 * Offset		Name		Size		Contents
 * -----------------------------------------------------------
 * 0				grbit		1			=01h
 * 1				num		8			IEEE FP number
 * 
 * String
 * Offset		Name		Size		Contents
 * -----------------------------------------------------------
 * 0				grbit		1			=02h
 * 1				cch			1			Length of the String
 * 2				rgch		var		the string.
 * 
 * If a formula contains more than one array constant, the token values for the array constants are appended to the saved
 * parsed expression in order: first the values for the first array constant, 
 * then the values for the second array constant, etc.
 * 
 * If a formula contains very long array constants, the FORMULA, ARRAY, or NAME record contaniing the parsed expression 
 * may overflow into CONTINUE  records.  In such cases, an individual array value is NEVER SPLIT between records, 
 * but record boundaries are established between adjacent array values.
 * 
 * The reference class ptgArray never appears in an Excel formula, only the ptgArrayV and ptgArrayA classes are used.
 * 
 * 
 * @see Ptg
 * @see Formula
    
*/
// 20090119-22 KSC: Many, many changes changes
public class PtgArray extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4416140231168551393L;
	int nc = -1;
	int nr = -1;
	byte[] rgval;
	CompatibleVector arrVals = new CompatibleVector();
	boolean isIntermediary = false;        // 20090824 KSC: true if this PtgArray is only part of a calcualtion process; if so, apparently can have more than 256 columns [BugTracker 2683]

	public boolean getIsOperand()
	{
		return true;
	}

	/**
	 * return the first 8 bytes of the ptgArray record
	 * this represents the id byte and 7 reserved bytes
	 *
	 * @return
	 */
	public byte[] getPreRecord()
	{
		/* 20090820 KSC: now record is always 8 bytes as rgval is now kept separate [BugTracker 2683]
		byte[] retbytes = new byte[8];
		if (record!=null)
			System.arraycopy(record, 0, retbytes, 0, 8);
		return retbytes;		
		*/
		return record;
	}

	/**
	 * these are the bytes appended to the formula token array, after all other ptg's
	 *
	 * @return
	 */
	public byte[] getPostRecord()
	{
		return rgval;
		/*	20090820 KSC: now record and rgval are kept separated (see populateVals) [BugTracker 2683]
		if (record==null) {	// why??
			return new byte[0];	// hits on testScenario[28]
		}
		byte[] retbytes = new byte[record.length -8];			
		System.arraycopy(record, 8, retbytes, 0, (record.length-8));	 
		return retbytes;
		*/
	}

	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	private void populateVals()
	{
/* 20090819 KSC: according to documentation, this isn't correct: [BugTracker 2683]		
		nc= record[1] & 0x00FF;			
		nr = ByteTools.readShort(record[2], record[3]);			
 		* if (record.length>11) {
			rgval = new byte[record.length-11];			
			System.arraycopy(record, 11, rgval, 0, rgval.length);	
			this.parseArrayComponents();
		}
*/
		if( record.length > 8 )
		{ // means that array data has already been appeneded to end of record array; store in rgvals
			rgval = new byte[record.length - 8];
			System.arraycopy( record, 8, rgval, 0, rgval.length );    // save post array= nc, nr + array data
		}
		if( rgval != null )
		{
			// clear out record array:0= id 1-7=reserved
			byte[] b = new byte[8];
			b[0] = record[0];
			record = b;
			this.parseArrayComponents();
		} // otherwise, it's just the initial input of the 1st 8 bytes record - see Formula 
	}

	/**
	 * given "extra info" at end of formula expression, parse array values
	 */
	public void parseArrayComponents()
	{
		int nitems = 0;
		arrVals.clear();    // 20090820 KSC: makes sense to! [BugTracker 2683]
		if( !isIntermediary ) // 20090824 KSC: sometimes an intermediary ptgarry can have more than 256 columns [BugTracker 2683]
		{
			nc = rgval[0] & 0xFF;    // number of columns
		}
		nr = ByteTools.readShort( rgval[1], rgval[2] );    // number of rows
		try
		{
			// (nc+1)*(nr+1) compoments
			for( int i = 3; i < rgval.length; )
			{    // 20090820 KSC: post array contains nc & nr so i should be initially 3 instead of 0 [BugTracker 2683]
				if( rgval[i] == 0 )
				{ // empty value
					i++;
					i += 8;
					arrVals.add( "" );    // TODO: Empty Constant should be null?
				}
				else if( rgval[i] == 0x1 )
				{ // its a number
					i++;
					byte[] barr = new byte[8];
					System.arraycopy( rgval, i, barr, 0, 8 );
					double val = ByteTools.eightBytetoLEDouble( barr );
					Double d = new Double( val );
					arrVals.add( d );
					i = i + 8;
				}
				else if( rgval[i] == 0x2 )
				{ // its a string
					int strLen = ByteTools.readShort( rgval[i + 1], rgval[i + 2] );
					i += 3;
					int grbt = rgval[i++];
					byte[] barr = new byte[strLen];
					System.arraycopy( rgval, i, barr, 0, strLen );
					String strVal = "";
					try
					{
						if( (grbt & 0x1) == 0x1 )
						{
							strVal = new String( barr, UNICODEENCODING );
						}
						else
						{
							strVal = new String( barr, DEFAULTENCODING );
						}
					}
					catch( UnsupportedEncodingException e )
					{
						Logger.logInfo( "decoding formula string in array failed: " + e );
					}
					arrVals.add( strVal );
					i += strLen;
				}
				else if( rgval[i] == 0x4 )
				{ // its a boolean
					if( rgval[++i] == 0 )
					{
						arrVals.add( Boolean.valueOf( false ) );
					}
					else
					{
						arrVals.add( Boolean.valueOf( true ) );
					}
					i = i + 8;
				}
				else if( rgval[i] == 0x10 )
				{ // it's an error value
					int errCode = rgval[++i];
					switch( errCode )
					{
						case 0:
							arrVals.add( "#NULL!" );
							break;
						case 0x7:
							arrVals.add( "#DIV/0!" );
							break;
						case 0x0F:
							arrVals.add( "#VALUE!" );
							break;
						case 0x17:
							arrVals.add( "#REF!" );
							break;
						case 0x1D:
							arrVals.add( "#NAME!" );
							break;
						case 0x24:
							arrVals.add( "#NUM!" );
							break;
						case 0x2A:
							arrVals.add( "#N/A!" );
							break;
					}
					i = i + 8;
				}
				nitems++;
				if( nitems == ((nc + 1) * (nr + 1)) )
				{ // Finished with this array!
					int length = i;
					i = rgval.length;
					// length may be less than rgval.length for cases of more than one array parameter
					// see ExpressionParser.parseExpression
					if( rgval.length != length )
					{// then truncate both record + rgval
						byte[] tmp = new byte[length];
						System.arraycopy( rgval, 0, tmp, 0, length );
						rgval = tmp;
					}
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error Processing Array Formula: " + e.toString() );
			return;
		}
	}

	public int getVal()
	{
		return -1;
	}

	public Object getValue()
	{
		// 20090820 KSC: value = entire array instead of 1st value; desired value is determined by cell position as compared to current formula; see Formula.calculate [BugTracker 2683]
		//return elementAt(0).getValue(-1);	// default: return 1st value
		return getString();
	}

	/*
	 *  returns the string value of the name
		@see com.extentech.formats.XLS.formulas.Ptg#getValue()
	 */
	public String getString()
	{
		Object retVal = null;
		Ptg[] p = this.getComponents();
		String retstr = "";
		if( nc == 0 && nr == 0 )
		{    // if it's a single value, just return val
			for( int i = 0; i < p.length; i++ )
			{
				if( i != 0 )
				{
					retstr += ",";
				}
				retstr += p[i].getValue().toString();
			}
		}
		else
		{
			retstr = "";
			int loc = 0;
			for( int x = 0; x < nr + 1; x++ )
			{
				if( x != 0 )
				{
					retstr += ";";
				}
				for( int i = 0; i < nc + 1; i++ )
				{
					if( i != 0 )
					{
						retstr += ",";
					}
					retstr += p[loc++].getValue().toString();
				}
			}
			//retstr += "}";
			//retVal = retstr.substring(0,retstr.length()-1);
		}
		retVal = retstr;
		return "{" + retVal + "}";
	}

	public String getTextString()
	{
		return getString();
	}

	public void setVal( String arrStr )
	{
		// remove the initial { and ending }
		arrStr = arrStr.substring( 1, arrStr.length() - 1 );
		if( arrStr.indexOf( "{" ) != -1 )
		{ // SHOULDN'T -- see FormulaParser.getPtgsFromFormulaString
			Logger.logErr( "PtgArray.setVal: Multiple Arrays Encountered" );
		}

		// parse all array strings into rows, cols
		String[] rows = null;
		String[][] cols = null;
		// split rows
		rows = arrStr.split( ";" );
		cols = new String[rows.length][];
		for( int i = 0; i < rows.length; i++ )
		{
			String[] s = rows[i].split( ",", -1 );    // include empty strings
			cols[i] = s;
		}
		byte[] databytes = new byte[11];
		databytes[0] = 0x60;                        // 20h=tArrayR, 40h=tArrayV, 60h=tArrayA
		isIntermediary = false;    // init value
		if( cols[0].length >= 255 )
		{    // 20090824 KSC:  apparently sometimes an intermediary calculations step can include > 256 array elements ...
			isIntermediary = true;
			nc = cols[0].length - 1;
		}
		databytes[8] = (byte) ((cols[0].length - 1) & 0xFF);        // nc-1	// 20090819 KSC: placed in wrong pt of record:  was [1]	[BugTracker 2683]
		//databytes[8] = (byte)((cols[0].length-1));		// nc-1	// 20090819 KSC: placed in wrong pt of record:  was [1]	[BugTracker 2683]
		System.arraycopy( ByteTools.shortToLEBytes( (short) (rows.length - 1) ),
		                  0,
		                  databytes,
		                  9,
		                  2 );        // nr-1  // 20090819 KSC: placed in wrong pt of record:  was 2,3 [BugTracker 2683]
		// iterate the array and fill out the data section
		for( int j = 0; j < rows.length; j++ )
		{
			for( int i = 0; i < cols[0].length; i++ )
			{
				byte[] valbytes = this.valuesIntoByteArray( cols[j][i] );
				databytes = ByteTools.append( valbytes, databytes );
			}
		}
		// populate primary values for rec
		record = databytes;
		this.init( databytes );
	}

	/**
	 * Turns a vector of values into a byte array representation for the data section of this record
	 *
	 * @param compVect
	 * @return
	 */
	private byte[] valuesIntoByteArray( String constVal )
	{
		byte[] databytes = new byte[0];
		byte[] thisElement = new byte[9];

		try
		{    // number?
			Double d = new Double( constVal );
			thisElement[0] = 0x1;        // id for number value
			byte[] b = ByteTools.toBEByteArray( d.doubleValue() );
			System.arraycopy( b, 0, thisElement, 1, b.length );
			databytes = ByteTools.append( thisElement, databytes );
		}
		catch( NumberFormatException ee )
		{
			try
			{
				if( constVal.equalsIgnoreCase( "true" ) || constVal.equalsIgnoreCase( "false" ) )
				{
					Boolean bb = Boolean.valueOf( constVal );
					thisElement[0] = 0x4;        // id for boolean value
					thisElement[1] = (byte) (bb.booleanValue() ? 1 : 0);
				}
				else if( constVal == null || constVal.equals( "" ) )
				{    // emtpy or null value
					thisElement[0] = 0x0;        // id for empty value
				}
				else if( constVal.charAt( 0 ) == '#' )
				{ // it's an error value
					thisElement[0] = 0x10;        // id for error value
					int errCode = 0;
					if( constVal.equals( "#NULL!" ) )
					{
						errCode = 0;
					}
					else if( constVal.equals( "#DIV/0!" ) )
					{
						errCode = 0x7;
					}
					else if( constVal.equals( "#VALUE!" ) )
					{
						errCode = 0x0F;
					}
					else if( constVal.equals( "#REF!" ) )
					{
						errCode = 0x17;
					}
					else if( constVal.equals( "#NAME!" ) )
					{
						errCode = 0x1D;
					}
					else if( constVal.equals( "#NUM!" ) )
					{
						errCode = 0x24;
					}
					else if( constVal.equals( "#N/A!" ) || constVal.equals( "#N/A" ) || constVal.equals( "N/A" ) )
					{
						errCode = 0x2A;
					}
					thisElement[1] = (byte) errCode;
				}
				else
				{    // assume string
					thisElement = new byte[3];
					try
					{
						thisElement = new byte[4];
						thisElement[0] = 0x2;        // id for string
						byte[] b = constVal.getBytes( UNICODEENCODING );
						System.arraycopy( ByteTools.shortToLEBytes( (short) b.length ), 0, thisElement, 1, 2 );
						thisElement[3] = 1;    // compressed= 0, uncompressed= 1 (16-bit chars)
						thisElement = ByteTools.append( b, thisElement );
					}
					catch( UnsupportedEncodingException z )
					{
						Logger.logWarn( "encoding formula array:" + z );
					}
				}
				databytes = ByteTools.append( thisElement, databytes );
			}
			catch( Exception ex )
			{
				Logger.logWarn( "PtgArray.valuesIntoByteArray:  error parsing array element:" + ex );
			}
		}
		return databytes;
	}

	/**
	 * Returns the second section of bytes for the PtgArray.
	 * These are the bytes that are split off the end of the
	 * formula

	 public void getComponentBytes(){

	 }
	 //public void updateRecord(){
	 //}*/

	/**
	 * Override due to mystery extra byte
	 * occasionally found in ptgName recs.
	 */
	public int getLength()
	{
		/* 20090820 KSC: really want record length not rgval length, which now is separate [BugTracker 2683]
		 *if (rgval!=null)
            return rgval.length;
        */
		return 8;
	}
    
    /* not used 
    public int getLength(byte[] b){
    	int co = b[1];
    	int rw = ByteTools.readShort(b[2], b[3]);
    	rw++; // appears that rows are not ordinal here...
    	int numrecs = co*rw;
    	int len = 4;
    	int loc = 4;
    	for (int i=0;i<=numrecs;i++){
			if (b[len] == 0x1){ // its a number
				len += 9;
			}else{
				len += b[len+1] + 2;
			}
    	}
    	length = len;
    	return length;
    }
    */

	public String toString()
	{
		return this.getString();
	}

	public Ptg[] getComponents()
	{
		Ptg[] retVals = new Ptg[arrVals.size()];
		for( int i = 0; i < arrVals.size(); i++ )
		{
			Object o = arrVals.elementAt( i );
			if( o instanceof Double )
			{
				Double d = (Double) o;
				PtgNumber pnum = new PtgNumber( d.doubleValue() );
				retVals[i] = pnum;
			}
			else if( o instanceof Boolean )
			{
				PtgBool pb = new PtgBool( ((Boolean) o).booleanValue() );
				retVals[i] = pb;
			}
			else
			{
				if( FormulaParser.isRef( (String) o ) || FormulaParser.isRange( (String) o ) )
				{    // it's a range
					PtgArea3d pa = new PtgArea3d();
					pa.setParentRec( this.getParentRec() );
					pa.setUseReferenceTracker( true );
					pa.setLocation( (String) o );
					Ptg[] pacomps = pa.getComponents();
					Ptg[] temp = new Ptg[retVals.length - 1 + pacomps.length];
					System.arraycopy( retVals, 0, temp, 0, retVals.length - 1 );
					System.arraycopy( pacomps, 0, temp, retVals.length - 1, pacomps.length );
					retVals = temp;
				}
				else
				{
					PtgStr pstr = new PtgStr( (String) o );
					retVals[i] = pstr;
				}
			}
		}
		return retVals;
	}

	/**
	 * returns the 0-based number of rows in this array
	 * if nr>1 then the array is in the form of:
	 * a,b,c;d,e,f; .... where the semicolons delineate rows
	 *
	 * @return
	 */
	public int getNumberOfRows()
	{
		return nr;
	}

	/**
	 * returns the 0-based number of columns in this array
	 * number of columns is the amount of elements before the semicolon (if present)
	 * a,b,c;d,e,f; ....
	 *
	 * @return
	 */
	public int getNumberOfColumns()
	{
		return nc;
	}

	/**
	 * sets the array components values for this PtgArray
	 * returns the actual array components length
	 *
	 * @see ExpressionParser.parseExpression
	 */
	public int setArrVals( byte[] by )
	{
		rgval = by;
		if( rgval != null )
		{
			// clear out record array:0= id 1-7=reserved
			byte[] b = new byte[8];
			b[0] = record[0];
			record = b;
			this.parseArrayComponents();
		}
		return rgval.length;

	}

	public byte[] getArrVals()
	{
		return rgval;
	}

	/**
	 * returns a ptg at the specified location.  Assumes that it is a one-dimensional
	 * array.  If you need a multidimensional array please use the other elementAt(int,int)method
	 *
	 * @param loc
	 * @return
	 */
	public Ptg elementAt( int loc )
	{
		Ptg[] p = this.getComponents();
		return p[loc];
	}

	public Ptg elementAt( int col, int row )
	{
		Ptg[] p = this.getComponents();
		try
		{
			int loc = 0;
			for( int i = 0; i < row; i++ )
			{
				loc += (nc);    // 20090816 KSC: why +1????   +1);  [BugTracker 2683]
			}
			loc += col;
			return elementAt( loc );
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
			Logger.logErr( "PtgArray.elementAt: error retrieving value at [" + row + "," + col + "]: " + e );
		}
		return null;
	}

}
    
    