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

import java.io.Serializable;

/**
 * Coordinates the various records involved in sheet-level protection.
 */
public class SheetProtectionManager extends ProtectionManager implements Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7450088022236591508L;

	/**
	 * The worksheet whose protection state this instance manages.
	 */
	private Boundsheet sheet;
	private ObjProtect objprotect;
	private ScenProtect scenprotect;

	/*
	The PROTECT record in the Worksheet Protection Block indicates that the sheet is protected.
	There may follow a SCENPROTECT record or/and an OBJECTPROTECT record.
	The optional PASSWORD record contains the hash value of the password used to protect the sheet
	In BIFF8, there may occur additional records following the cell records in the Sheet Substream
	Sheet protection with password does not cause to switch on read/write file protection.
	Therefore the file will not be encrypted.
	Structure of the Worksheet Protection Block, BIFF5-BIFF8:
	○ PROTECT Worksheet contents: 1 = protected
	○ OBJECTPROTECT Embedded objects: 1 = protected (if not present, objects are not protected)
	○ SCENPROTECT Scenarios: 1 = protected (if not present, not protected)
	○ PASSWORD Hash value of the password; 0 = no password
	 */
	public SheetProtectionManager( Boundsheet sheet )
	{
		this.sheet = sheet;
	}

	/**
	 * Adds a protection-related record to be managed.
	 * This method will be called automatically by the <code>init()</code>
	 * method of the record where appropriate. It should probably never be
	 * called anywhere else.
	 *
	 * @param record the record to be managed
	 */
	@Override
	public void addRecord( BiffRec record )
	{
		if( record instanceof ObjProtect )
		{
			objprotect = (ObjProtect) record;
		}
		else if( record instanceof ScenProtect )
		{
			scenprotect = (ScenProtect) record;
		}
		else if( record instanceof FeatHeadr )
		{
			enhancedProtection = (FeatHeadr) record;
		}
		else
		{
			super.addRecord( record );
		}
	}

	/**
	 * Sets whether the sheet is protected.
	 */
	@Override
	public void setProtected( boolean value )
	{
		if( value )
		{
			if( protect == null )
			{
				addProtectionRecord();
			}

			// copy legacy values from the EnhancedProtection record
			if( enhancedProtection != null )
			{
				if( enhancedProtection.getProtectionOption( FeatHeadr.ALLOWOBJECTS ) )
				{
					if( objprotect == null )
					{
						insertObjProtect();
					}
				}

				setScenProtect( enhancedProtection.getProtectionOption( FeatHeadr.ALLOWSCENARIOS ) );
			}
		}
		else
		{
			// the Protect record's presence implies protection, so remove it
			if( protect != null )
			{
				sheet.removeRecFromVec( protect );
				protect = null;
			}

			// the ObjProtect record cannot exist if protection is disabled
			if( objprotect != null )
			{
				sheet.removeRecFromVec( objprotect );
				objprotect = null;
			}

			// Note that ScenProtect and EnhancedProtection can and should be
			// retained when disabling protection so that the same settings
			// will be restored if it is re-enabled.
		}
	}

	/**
	 */
	public boolean getProtected( int option )
	{
		if( enhancedProtection != null )
		{
			return enhancedProtection.getProtectionOption( option );
		}

		// fallback if we don't have an EnhancedProtection record
		switch( option )
		{
			case FeatHeadr.ALLOWOBJECTS:
				// if the sheet is protected, check for an ObjProtect
				if( protect != null )
				{
					return objprotect == null;
				}
				return true; // defaults to allowed

			case FeatHeadr.ALLOWSCENARIOS:
				if( scenprotect != null )
				{
					return scenprotect.getIsLocked();
				}
				return true; // defaults to allowed

			// these extended settings default to allowed
			case FeatHeadr.ALLOWSELLOCKEDCELLS:
			case FeatHeadr.ALLOWSELUNLOCKEDCELLS:
				return true;

			// all other extended settings default to prohibited
			default:
				return false;
		}
	}

	/**
	 */
	public void setProtected( int option, boolean value )
	{
		// special cases for legacy records
		switch( option )
		{
			case FeatHeadr.ALLOWOBJECTS:
				if( value )
				{
					if( protect != null )
					{
						insertObjProtect();
					}
				}
				else
				{
					if( objprotect != null )
					{
						sheet.removeRecFromVec( objprotect );
					}
				}
				break;

			case FeatHeadr.ALLOWSCENARIOS:
				if( protect != null )
				{
					setScenProtect( value );
				}
				break;
		}

		// add an enhanced protection record if one does not exist
		if( enhancedProtection == null )
		{
			enhancedProtection = (FeatHeadr) FeatHeadr.getPrototype();
			int i = sheet.getSheetRecs().size() - 1;    // just before EOF;
			sheet.insertSheetRecordAt( enhancedProtection, i );
		}

		// set the value in the enhanced protection record
		enhancedProtection.setProtectionOption( option, value );
	}

	private void insertPassword()
	{
		if( password != null )
		{
			return;
		}

		password = new Password();
		if( protect == null )
		{
			addProtectionRecord();
		}
		int insertIdx = protect.getRecordIndex();

		if( protect != null )
		{
			insertIdx++;
		}
		if( objprotect != null )
		{
			insertIdx++;
		}
		if( scenprotect != null )
		{
			insertIdx++;
		}

		sheet.insertSheetRecordAt( password, insertIdx );
	}

	private void removePassword()
	{
		if( password == null )
		{
			return;
		}

		sheet.removeRecFromVec( password );
		password = null;
	}

	/**
	 * Sets or removes the sheet's protection password.
	 *
	 * @param pass the string password to set or null to remove the password
	 */
	@Override
	public void setPassword( String pass )
	{
		if( (pass != null) && !pass.equals( "" ) )
		{
			insertPassword();
			super.setPassword( pass );
		}
		else
		{
			removePassword();
		}
	}

	/**
	 * Sets or removes the sheet's protection password.
	 *
	 * @param pass the pre-hashed string password to set
	 *             or null to remove the password
	 */
	@Override
	public void setPasswordHashed( String pass )
	{
		if( pass != null )
		{
			insertPassword();
			super.setPasswordHashed( pass );
		}
		else
		{
			removePassword();
		}
	}

	private void setScenProtect( boolean value )
	{
		if( scenprotect == null )
		{
			scenprotect = new ScenProtect();
			sheet.insertSheetRecordAt( scenprotect, protect.getRecordIndex() + 1 );
		}
		scenprotect.setLocked( value );
	}

	private void insertObjProtect()
	{
		if( objprotect == null )
		{
			objprotect = new ObjProtect();
			objprotect.setLocked( true );
			sheet.insertSheetRecordAt( objprotect, ((scenprotect != null) ? scenprotect : protect).getRecordIndex() + 1 );
		}
	}

	/**
	 * create and add new protection record to sheet records
	 */
	private void addProtectionRecord()
	{
		protect = new Protect();
		// baseRecordIndex= 18 
		// no! is: PROTECTIONBLOCK, [DEFCOLW], [COLINFO], [SORT], DIMENSIONS
		int i = 0;
		while( ++i < sheet.getSheetRecs().size() )
		{
			int opc = ((BiffRec) sheet.getSheetRecs().get( i )).getOpcode();
			if( opc == XLSConstants.PROTECT )
			{
				protect = ((Protect) sheet.getSheetRecs().get( i ));
				i++;
				if( ((BiffRec) sheet.getSheetRecs().get( i )).getOpcode() == XLSConstants.SCENPROTECT )
				{
					scenprotect = ((ScenProtect) sheet.getSheetRecs().get( i ));
					i++;
				}
				if( ((BiffRec) sheet.getSheetRecs().get( i )).getOpcode() == XLSConstants.OBJPROTECT )
				{
					objprotect = ((ObjProtect) sheet.getSheetRecs().get( i ));
				}
				break;
			}
			if( (opc == XLSConstants.DEFCOLWIDTH) || (opc == XLSConstants.COLINFO) || (opc == XLSConstants.DIMENSIONS) )
			{
				break;
			}
		}
		sheet.insertSheetRecordAt( protect, i );
		protect.setLocked( true );
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		super.close();
		sheet = null;
		if( objprotect != null )
		{
			objprotect.close();
		}
		if( scenprotect != null )
		{
			scenprotect.close();
		}
	}
}
