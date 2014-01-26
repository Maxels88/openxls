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
package docs.samples.Comments;

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.CommentHandle;
import org.openxls.ExtenXLS.DocumentObjectNotFoundException;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TestComments
{
	private static final Logger log = LoggerFactory.getLogger( TestComments.class );
	public WorkBookHandle book;
	public WorkSheetHandle sheet;
	String workingdir = System.getProperty( "user.dir" ) + "/docs/samples/Comments/";

	public static void main( String[] args )
	{
		TestComments t = new TestComments();
		t.addNotes();        // illustrate adding notes/comments to a new workbook
		t.manipulateComments();    // illustrate basic note/comment manipulation
		t.editNotes();
		t.showNotes();        // illustrate how to manipulate note/comment visible state (hidden upon loss of focus/always shown)
		t.hideNotes();        // ""
		t.removeNotes();    // illustrates how to remove comments/notes
		t.editFormats();    // illustrate how to add notes/comments with formatting information
	}

	/**
	 * shows how to adding notes or Comments to a new workbook
	 * <br>Comments can be added via WorkSheetHandle.createNote or
	 * <br>CellHandle.createComment
	 */
	public void addNotes()
	{
		book = new WorkBookHandle();
		try
		{
			sheet = book.getWorkSheet( 0 );
			sheet.add( "Cell A1", "A1" );
			sheet.add( "Cell D2", "D2" );
			CellHandle cell = sheet.add( "Cell F7", "F7" );
			// add a note to cell A1, specifying note text and author name
			sheet.createNote( "A1", "This is a note attached to Cell A1", "cagney" );
			// add a note to cell D2, specifying note text and author name
			CommentHandle nh = sheet.createNote( "D2", "this is a note attached to D2\nshown", "maya" );
			nh.show();    // make note at D2 always shown
			cell.createComment( "Another Note attached to F7.  This is very very long\nMultiLine\nRambling ...", "elmer" );
			book.write( workingdir + "testAddNotesOUT.xls" );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}

	/**
	 * shows how to manipulate notes via a CommentHandle
	 */
	public void manipulateComments()
	{
		book = new WorkBookHandle();
		try
		{
			sheet = book.getWorkSheet( 0 );
			// create and modify a CommentHandle
			CellHandle cell = sheet.add( "Cell A1", "A1" );
			// CommentHandle can be used to manipulate comments
			CommentHandle comment = cell.createComment( "This is my comment", "noteAuthor" );
			comment.setAuthor( "newAuthor" );
			comment.setCommentText( "New comment text" );
			// retrieve CommentHandle from a worksheet and delete it.
			comment = null;
			comment = cell.getComment();
			cell.removeComment();

			// get all comments from the worksheet
			CommentHandle[] comments = sheet.getCommentHandles();
			book.write( workingdir + "testNotesOUT.xls" );
		}
		catch( WorkSheetNotFoundException e )
		{

		}
		catch( DocumentObjectNotFoundException e )
		{
		}
	}

	/**
	 * shows altering the text of existing notes
	 */
	public void editNotes()
	{
		// contains 3 notes at locations: A1, D2 and D8
		book = new WorkBookHandle( workingdir + "testNotes-SIMPLEII.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			CommentHandle[] nhs = sheet.getCommentHandles();
			// assert initial state - NOTE: getCommentText includes Author Name if present
			log.info( "A1: " + nhs[0].getCommentText() );    // "Ted:\nA note for cell A1"
			log.info( "Location" + nhs[0].getAddress() );    // "Sheet1!A1"
			log.info( "D2: " + nhs[1].getCommentText() );        // "Maya:\na note attached to cell D2";
			log.info( "Location" + nhs[1].getAddress() );    // "Sheet1!D2"
			log.info( "D8: " + nhs[2].getCommentText() );        // "James:\nThis is a longer comment attached to D8, it is also, along with D2, NOT hidden");
			log.info( "Location" + nhs[2].getAddress() );    // "Sheet1!D8"
			// change 'em
			nhs[0].setCommentText( "A new note for A1" );
			nhs[2].setCommentText( "A considerably shorter note for D8" );
			book.write( workingdir + "testNotesEditOUT.xls" );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}

	/**
	 * shows how to remove notes from an exisitng wb
	 */
	public void removeNotes()
	{

		book = new WorkBookHandle( workingdir + "testNotes-SIMPLEII.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			CommentHandle[] nhs = sheet.getCommentHandles();
			// has 3 notes - remove 1
			nhs[1].remove();
			book.write( workingdir + "testNotesRemoveOUT.xls" );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}

	/**
	 * shows how to set a comment to always be displayed i.e not hidden upon loss of focus
	 */
	public void showNotes()
	{
		// SIMPLE has one hidden note at A1
		book = new WorkBookHandle( workingdir + "testNotes-SIMPLE.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			CommentHandle[] nhs = sheet.getCommentHandles();
			nhs[0].show();    // make always shown - getIsHidden() should return false
			log.info( "Is Hidden? " + nhs[0].getIsHidden() );    // getIsHidden indicataes whether it is hidden - normal state - or always shown (false)
			book.write( workingdir + "testNotesShowOUT.xls" );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}

	/**
	 * shows how to set a note to be hidden on loss of focus (the default setting)
	 */
	public void hideNotes()
	{
		// SIMPLEII has one hidden note at A1 and two shown notes at D2 and D8
		book = new WorkBookHandle( workingdir + "testNotes-SIMPLEII.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			CommentHandle[] nhs = sheet.getCommentHandles();
			// assert initial state
			log.info( "Location of Comment= " + nhs[0].getAddress() ); //"Sheet1!A1"
			log.info( "Is Hidden? " + nhs[0].getIsHidden() );    // true
			log.info( "Location of Comment= " + nhs[1].getAddress() ); // "Sheet1!D2"
			log.info( "Is Hidden? " + nhs[1].getIsHidden() );    // false i.e it's always shown
			log.info( "Location of Comment= " + nhs[2].getAddress() ); // "Sheet1!D8"
			log.info( "Is Hidden? " + nhs[2].getIsHidden() );    // false
			nhs[0].show();    // show A1
			nhs[0].setCommentText( nhs[0].getCommentText() + "\nSHOW" );
			nhs[1].hide();
			nhs[1].setCommentText( nhs[1].getCommentText() + "\nHIDDEN" );
			nhs[2].hide();
			nhs[2].setCommentText( nhs[2].getCommentText() + "\nHIDDEN" );
			book.write( workingdir + "testNotesHideOUT.xls" );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}

	/**
	 * shows how to format comments via embedded formatting information within the comment text
	 * <p>the format of the embedded information is as follows:
	 * <p>&lt;font specifics>text segment&lt;font specifics for next segment>text segment...
	 * <br>where font specifics can be one or more of:
	 * <ul>b		if present, bold
	 * <br>i		if present, italic
	 * <br>s		if present, strikethru
	 * <br>u		if present, underlined
	 * <br>f=""		font name e.g. "Arial"
	 * <br>sz=""	font size in points e.g. "10"
	 * <br>delimited by ;'s
	 * <br>For Example:
	 * <br>&lt;f="Tahoma";b;sz="16">Note: < f="Tahoma";sz="12">This is an important point
	 */
	public void editFormats()
	{
		book = new WorkBookHandle();
		String outputfile = "testEditFormatsOUT.xls";
		try
		{
			sheet = book.getWorkSheet( 0 );
			sheet.add( "CellA1", "A1" );
			String noteTextWithFormats = "<f=\"Tahoma\";b;sz=\"14\">Note: <f=\"Tahoma\";i;sz=\"10\">Testing: <f=\"Tahoma\";u;sz=\"10\">A new note for A1";
			String noteTextWithoutFormats = "Note: Testing: A new note for A1";
			CommentHandle nh = sheet.createNote( "A1", noteTextWithFormats, "an author" );
			if( !noteTextWithoutFormats.equals( nh.getCommentText() ) )
			{
				log.info( "Note text does not equal expected" );
			}
			book.write( workingdir + outputfile );
		}
		catch( Exception e )
		{
			log.info( e.toString() );
		}
	}
}
