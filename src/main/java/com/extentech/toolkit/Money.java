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
package com.extentech.toolkit;

import java.io.Serializable;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;

/**
 * <b>Class Money</b>
 * <p/>
 * <p>
 * The Money class represents a United States
 * monetary value, expressed in dollars and cents. Internally,
 * the
 * value is represented using Java's BigDecimal class.
 * </p>
 * <p/>
 * <p>
 * Methods are provided to perform all the usual arithmetic
 * manipulations required when dealing with monetary data,
 * including add, subtract, multiply, and divide.
 * </p>
 * <p/>
 * <p>
 * <b>Rounding</b>
 * </p>
 * <p>
 * Rounding does not occur during intermediate computations;
 * maximum precision
 * (and accuracy) is thus preserved throughout all computations.
 * Round-off to an integral cent occurs only when the
 * monetary value is externalized (when formatted for display as
 * a String or
 * converted to a long integer). One of several different
 * rounding modes
 * can be specified. The default rounding mode is to discard any
 * fractional
 * cent and truncate the monetary value to 2 decimal places.
 * </p>
 * <p/>
 * <p>
 * <b>Currency Format</b>
 * </p>
 * <p>
 * A Currency Format (an instance of DecimalFormat) is used to
 * control formatting of
 * monetary values for display as well as the parsing of strings
 * which represent
 * monetary values.
 * By default, the Currency Format for the current locale is
 * used. For the
 * United States, the default Currency Symbol is the Dollar Sign
 * ("$").
 * Negative amounts are enclosed in parentheses. A Decimal Point
 * (".")
 * separates the dollars and cents, and a comma (",") separates
 * each group
 * of 3 consecutive digits in the dollar amount.
 * </p>
 * <p>
 * Examples: $1,234.56   ($1,234.56)
 * </p>
 * <p/>
 * <p>
 * <b>Immutability</b>
 * </p>
 * <p>
 * Money objects, like String objects, are <b>immutable</b>.
 * An operation on a Money object (such as add, subtract, etc.)
 * does not alter the object in any way.
 * Rather, a new Money object is returned whose state reflects
 * the result of the operation.
 * Thus, a statement like
 * </p>
 * <p>
 * <tt>money1.add(money2);</tt>
 * </p>
 * <p>
 * has no effect; it does not modify money1 in any way, and the
 * result is effectively discarded.
 * If the intent is to modify money1, then you should code
 * </p>
 * <p>
 * money1 = money1.add(money2);
 * </p>
 * <p>
 * which effectively replaces money1 with the result.
 * </p>
 *
 * @see BigDecimal
 * @see DecimalFormat
 */
public class Money // Money Class
		implements Cloneable, // Money objects can be cloned
		           Serializable // Money objects are serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2055249101317798370L;
	/**
	 * The monetary value.
	 */
	protected BigDecimal value = null; // The monetary value
	/**
	 * The Rounding Mode. Specifies if and how the monetary value is
	 * to be rounded off to an integral cent.
	 */
	protected int roundingMode = BigDecimal.ROUND_DOWN; // Rounding Mode
	/**
	 * The Currency Format, used for formatting and parsing a
	 * monetary value. Refer to the Java API
	 * documentation for the DecimalFormat class for information on
	 * formats.
	 */
	protected DecimalFormat currencyFormat = (DecimalFormat) NumberFormat.getCurrencyInstance();
	// Currency format
	/**
	 * The special monetary value of zero ($0.00).
	 */
	protected static final BigDecimal ZERO = new BigDecimal( "0.00" );
	// The value $0.00

	/**
	 * <b>Class InvalidScaleFactorException</b>
	 * <p/>
	 * <p>
	 * The InvalidScaleFactorException is thrown if an invalid scale
	 * factor is specified (valid scale factors are 0, 1, and 2).
	 * This is a non-checked exception and will be detected at
	 * runtime only.
	 * </p>
	 */
	public static class InvalidScaleFactorException extends RuntimeException // Non-checked exception
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = -8038085965896123803L;

		/**
		 * Default constructor for a InvalidScaleFactorException object.
		 */
		public InvalidScaleFactorException()
		{
			super(); // Invoke super class constructor.
		}

		/**
		 * Constructor for a InvalidScaleFactorException object.
		 * <p/>
		 * <p/>
		 *
		 * @param info Descriptive information
		 *             </p>
		 */
		public InvalidScaleFactorException( String info ) // Descriptive info
		{
			super( info ); // Invoke super class constructor.
		}
	} // Class InvalidScaleFactorException

	/**
	 * <b>Class InvalidRoundingModeException</b>
	 * <p/>
	 * <p>
	 * The InvalidRoundingModeException is thrown if an invalid
	 * Rounding
	 * Mode is specified (all rounding modes except
	 * ROUND_UNNECESSARY are valid).
	 * This is a non-checked exception and will be detected at
	 * runtime only.
	 * </p>
	 *
	 * @see BigDecimal#ROUND_UP
	 * @see BigDecimal#ROUND_DOWN
	 * @see BigDecimal#ROUND_CEILING
	 * @see BigDecimal#ROUND_FLOOR
	 * @see BigDecimal#ROUND_HALF_UP
	 * @see BigDecimal#ROUND_HALF_DOWN
	 * @see BigDecimal#ROUND_HALF_EVEN
	 */
	public static class InvalidRoundingModeException extends RuntimeException // Non-checked exception
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = 5658836125641516151L;

		/**
		 * Default constructor for a InvalidRoundingModeException object.
		 */
		public InvalidRoundingModeException()
		{
			super(); // Invoke super class constructor.
		} // Default Constructor InvalidRoundingModeException()

		/**
		 * Constructor for a InvalidRoundingModeException object.
		 * <p/>
		 * <p/>
		 *
		 * @param info Descriptive information
		 *             </p>
		 */
		public InvalidRoundingModeException( String info ) // Descriptive info
		{
			super( info ); // Invoke super class constructor.
		}
	} // Class InvalidRoundingModeException

	/**
	 * Default Constructor for a Money object; creates an object
	 * whose value is $0.00.
	 */
	public Money()
	{
		value = ZERO; // Initialize the monetary value
		// to $0.00.
	} // Default Constructor Money()

	/**
	 * Constructs a Money object from a double-precision,
	 * floating-point value.
	 * The Currency Format is set to the default format for the
	 * current locale.
	 * <p/>
	 * <p>
	 * The integral part of the value represents whole dollars, and
	 * the fractional
	 * part of the value represents fractional dollars (cents). As
	 * an example,
	 * the value 19.95 would represent $19.95.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>amount</b> The monetary amount, in dollars
	 *                      and cents
	 *                      </p>
	 */
	public Money( double amount ) // Monetary Value, dollars and cents
	{
		// In general, a double floating point value cannot represent a
		// decimal value exactly, and
		// therefore is only a very close approximation of the actual decimal
		// value. Fortunately,
		// the Double.toString() method is able to "recognize" an approximated
		// decimal value, and
		// will return the original (approximated) decimal value, rather than
		// the literal floating
		// point value. Here, we take advantage of this fact to obtain a
		// string representation of
		// the monetary value (without formatting, except for the decimal
		// point), which we then use
		// to create a BigDecimal object having the required value.
		// Were we not to make this simplifying assumption, the only
		// alternative would be to parse the
		// string ourselves, which can be quite complicated, and would
		// unnecessarily duplicate code
		// already implemented by the format's parse() method.
		// Convert the parsed value to a simple string (decimal point only)
		// and use it to set the monetary value.
		value = new BigDecimal( Double.toString( amount ) );
	} // Constructor Money(double amount)

	/**
	 * Constructs a Money object from a long integer value. The
	 * specified
	 * value represents whole dollars only (that is, cents are
	 * implicity assumed
	 * to be 00). For example, the integer 25 would represent a
	 * monetary value of $25.00.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>amount</b> The monetary amount, in whole
	 *                      dollars (no cents)
	 *                      </p>
	 */
	public Money( long amount ) // Monetary amount, whole dollars (no cents)
	{
		value = new BigDecimal( Long.toString( amount ) ); // Set monetary value.
	} // Constructor Money(long amount)

	/**
	 * Constructs a Money object from a long integer value with a
	 * specified scale factor.
	 * The scale factor (0, 1, or 2) specifies the number of digits
	 * to the right of an
	 * implied decimal point.
	 * <p/>
	 * <p>
	 * For example, the value 1995 would be interpreted as a
	 * monetary value of
	 * $1995.00, $199.50, and $19.95 for scale factors of 0, 1, or
	 * 2, respectively.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>amount</b> The monetary amount, in dollars
	 *                      and cents
	 * @param <b>scale</b>  Scale factor (must be 0, 1, or 2)
	 *                      </p>
	 *                      <p/>
	 *                      <p/>
	 * @throws InvalidScaleFactorException The scale factor
	 *                                     specified is not valid (must be 0, 1, or 2)
	 *                                     </p>
	 */
	public Money( long amount, // Montetary amount
	              int scale ) // Scale Factor (0, 1, 2)
			throws InvalidScaleFactorException
	{
		if(// If the Scale Factor is not
				(scale < 0) // 0, 1, or
						|| (scale > 2) // 2,
				)
		{
			throw new InvalidScaleFactorException( "Invalid scale factor: " + scale + " (must be 0, 1, or 2)" );
		}
		// Set the monetary value and scale as specified.
		value = (new BigDecimal( Long.toString( amount ) )).movePointLeft( scale );
	} // Constructor Money(long amount, int scale)

	/**
	 * Constructs a Money object from a string representation of a
	 * monetary value. The
	 * format of the string must be consistent with the Currency
	 * Format; otherwise,
	 * a ParseException is recognized.
	 * <p/>
	 * Refer to the Java API documentation for the DecimalFormat
	 * class for information on
	 * Decimal Formats in general, and Currency Formats in
	 * particular.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>string</b> A String representing a monetary
	 *                      value
	 *                      </p>
	 *                      <p/>
	 *                      <p/>
	 * @throws ParseException The string is inconsistent with
	 *                        the Currency Format
	 *                        </p>
	 * @see DecimalFormat
	 * @see #setCurrencyFormat
	 */
	public Money( String string ) // String representing a Monetary Value
			throws ParseException // If string is inconsistent with
	// Currency Format
	{
		// We make use of the format's parse() method, which parses the string
		// according to the format and returns either a Double or a Long
		// object representing the value.
		Number number;
		// Attempt to parse the string as a monetary value. May throw
		// ParseException if the string is not consistent with the Currency
		// Format.
		number = currencyFormat.parse( string );
		// In general, a double floating point value cannot represent a
		// decimal value exactly, and
		// therefore is only a very close approximation of the actual decimal
		// value. Fortunately,
		// the Double.toString() method is able to "recognize" an approximated
		// decimal value, and
		// will return the original (approximated) decimal value, rather than
		// the literal floating
		// point value. Here, we take advantage of this fact to obtain a
		// string representation of
		// the monetary value (without formatting, except for the decimal
		// point), which we then use
		// to create a BigDecimal object having the required value.
		// Were we not to make this simplifying assumption, the only
		// alternative would be to parse the
		// string ourselves, which can be quite complicated, and would
		// unnecessarily duplicate code
		// already implemented by the format's parse() method.
		// Convert the parsed value to a simple string (decimal point only)
		// and use it to set the monetary value.
		value = new BigDecimal( number.toString() );
	} // Constructor Money(String)

	/**
	 * Constructs a Money object from a BigDecimal object.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>amount</b> A BigDecimal value representing a
	 *                      monetary amount, in dollars and
	 *                      cents
	 *                      </p>
	 */
	public Money( BigDecimal amount ) // A BigDecimal object representing a
	// monetary amount
	{
		value = new BigDecimal( amount.toString() ); // Set the monetary value.
	} // Constructor Money(BigDecimal amount)

	/**
	 * Copy Constructor; constructs a Money object from another
	 * Money object.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>amount</b> A Money object
	 *                      </p>
	 */
	public Money( Money amount ) // Money Object
	{
		roundingMode = amount.roundingMode; // Copy the Rounding Mode.
		// Clone the Currency Format.
		currencyFormat = (DecimalFormat) amount.currencyFormat.clone();
		value = amount.value; // Copy the Monetary Value.
	} // Copy Constructor Money(Money money)

	/**
	 * Adds a specified monetary value to this monetary value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>money</b> The monetary value to be added
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money add( Money money ) // The Monetary Value to be added
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.value = value.add( money.value ); // Add the two monetary values.
		return result; // Return the result.
	} // Method Money.add()

	/**
	 * Subtracts a specified monetary value from this monetary value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>money</b> The monetary value to be subtracted
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money subtract( Money money ) // The monetary value to be
	// subtracted
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.value = value.subtract( money.value ); // Subtract the specified
		// monetary value.
		return result; // Return the result.
	} // Method Money.subtract()

	/**
	 * Multiplies this monetary value by a specified value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>mult</b> The multiplier value
	 *                    </p>
	 *                    <p/>
	 *                    <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money multiply( double mult ) // The multiplier value
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		// Multiply by the specified value.
		result.value = value.multiply( new BigDecimal( mult ) );
		return result; // Return the result.
	} // Method Money.multiply()

	/**
	 * Multiplies this monetary value by a specified value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>mult</b> The multiplier value
	 *                    </p>
	 *                    <p/>
	 *                    <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money multiply( long mult ) // The multiplier value
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		// Multiply by the specified value.
		result.value = value.multiply( new BigDecimal( Long.toString( mult ) ) );
		return result; // Return the result.
	} // Method Money.multiply()

	/**
	 * Divides this monetary value by a specified value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>div</b> The divisor value
	 *                   </p>
	 *                   <p/>
	 *                   <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money divide( double div ) // The divisor value
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		// Divide the monetary value by the specified value and round if
		// necessary.
		result.value = value.divide( new BigDecimal( div ), BigDecimal.ROUND_HALF_UP );
		return result; // Return the result.
	} // Method Money.divide()

	/**
	 * Divides this monetary value by a specified value.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>div</b> The divisor value
	 *                   </p>
	 *                   <p/>
	 *                   <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money divide( long div ) // The divisor value
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		// Divide the monetary value by the specified value and round if
		// necessary.
		result.value = value.divide( new BigDecimal( Long.toString( div ) ), BigDecimal.ROUND_HALF_UP );
		return result; // Return the result.
	} // Method Money.divide()

	/**
	 * Negates this monetary value.
	 * <p/>
	 * Positive values become negative, and negative values become
	 * positive. The effect is
	 * the same as if the monetary value were multiplied by -1.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money negate()
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.value = value.negate(); // Negate the monetary value.
		return result; // Return the result.
	} // Method Money.negate()

	/**
	 * Returns the absolute monetary value.
	 * <p/>
	 * A positive value is returned, irrespective of whether the
	 * monetary
	 * value is positive or negative.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @return A new Money object representing the
	 * result
	 * </p>
	 */
	public Money abs()
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.value = value.abs(); // Get the abolute monetary value.
		return result; // Return the result.
	} // Method Money.abs()

	/**
	 * Returns the monetary value as a long integer with 2 decimal
	 * digits to the
	 * right of an implicit decimal point. For example, if the
	 * monetary value were $19.95,
	 * a value of 1995 would be returned. Note that the monetary
	 * value is
	 * rounded, if necessary, according to the specified Rounding
	 * Mode.
	 * <p/>
	 * <p/>
	 *
	 * @return The monetary value
	 * </p>
	 */
	public long toLong()
	{
		// Round off the monetary value to 2 decimal places.
		BigDecimal result = value.setScale( 2, roundingMode );
		result = result.movePointRight( 2 ); // Move decimal point 2 places to
		// the right to preserve cents.
		return result.longValue(); // Return the result.
	} // Method Money.toLong()

	/**
	 * Returns the monetary value as a double-precision,
	 * floating-point value.
	 * <p/>
	 * <p>
	 * Note: Exercise care when converting monetary values to
	 * floating point
	 * values, because floating-point arithmetic is not
	 * well-suited for
	 * use with monetary data.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @return The monetary value
	 * </p>
	 */
	public double toDouble()
	{
		return value.doubleValue(); // Return the monetary value as a
		// double floating point value.
	} // Method Money.toDouble()

	/**
	 * Returns the monetary value as a BigDecimal object.
	 * <p/>
	 * <p/>
	 *
	 * @return The monetary value
	 * </p>
	 */
	public BigDecimal getValue()
	{
		return value; // Return the monetary value.
	} // Method Money.getValue()

	/**
	 * Returns the Rounding Mode.
	 * <p/>
	 * Refer to Java's BigDecimal object for a description of the
	 * possible
	 * rounding modes.
	 * <p/>
	 * <p/>
	 *
	 * @return The Rounding Mode
	 * </p>
	 * @see BigDecimal#ROUND_UP
	 * @see BigDecimal#ROUND_DOWN
	 * @see BigDecimal#ROUND_CEILING
	 * @see BigDecimal#ROUND_FLOOR
	 * @see BigDecimal#ROUND_HALF_UP
	 * @see BigDecimal#ROUND_HALF_DOWN
	 * @see BigDecimal#ROUND_HALF_EVEN
	 */
	public int getRoundingMode()
	{
		return roundingMode; // Return the Rounding Mode.
	} // Method Money.getRoundingMode()

	/**
	 * Sets the Rounding Mode.
	 * <p/>
	 * Refer to the Java API documentation for the BigDecimal class
	 * for a description of
	 * the possible Rounding Modes.
	 * <p/>
	 * The default Rounding Mode is
	 * BigDecimal.ROUND_DOWN, which effectively discards any
	 * fractional cent amount
	 * and truncates the monetary value to 2 decimal places.
	 * <p/>
	 * A Rounding Mode of BigDecimal.ROUND_UNNECESSARY is not valid
	 * for use with
	 * monetary data since certain operations result in a loss of
	 * precision and
	 * therefore require that a rounding mode be explicitly
	 * specified.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>int</b> The Rounding Mode
	 *                   </p>
	 *                   <p/>
	 *                   <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 * <p/>
	 * <p/>
	 * @throws InvalidRoundingModeException The Rounding Mode
	 *                                      specified is not valid for monetary data
	 *                                      </p>
	 * @see BigDecimal#ROUND_UP
	 * @see BigDecimal#ROUND_DOWN
	 * @see BigDecimal#ROUND_CEILING
	 * @see BigDecimal#ROUND_FLOOR
	 * @see BigDecimal#ROUND_HALF_UP
	 * @see BigDecimal#ROUND_HALF_DOWN
	 * @see BigDecimal#ROUND_HALF_EVEN
	 */
	public Money setRoundingMode( int mode ) // Rounding Mode
			throws InvalidRoundingModeException
	{
		if( mode == BigDecimal.ROUND_UNNECESSARY ) // If Rounding Mode is not
		// valid,
		{
			throw new InvalidRoundingModeException( "Rounding mode not valid for monetary data: " + mode );
		}
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.roundingMode = mode; // Set the Rounding Mode.
		return result; // Return the result.
	} // Method Money.setRoundingMode(int mode)

	/**
	 * Returns the Currency Format, used to format and parse
	 * monetary values.
	 * <p/>
	 * Refer to the Java API documentation for the DecimalFormat
	 * class for information on
	 * Decimal Formats in general, and Currency Formats in
	 * particular.
	 * <p/>
	 * <p/>
	 *
	 * @return The Currency Format
	 * </p>
	 * @see DecimalFormat
	 */
	public DecimalFormat getCurrencyFormat()
	{
		return currencyFormat; // Return the Currency Format.
	} // Method Money.getCurrencyFormat()

	/**
	 * Sets the Currency Format, used for formatting and parsing
	 * monetary values.
	 * <p/>
	 * Refer to the Java API documentation for the DecimalFormat
	 * class for information on
	 * Decimal Formats in general, and Currency Formats in particular.
	 * <p/>
	 * <p>
	 * By default, the Currency Format for the current locale is
	 * used. For the
	 * United States, the default Currency Symbol is the Dollar Sign
	 * ("$").
	 * Negative amounts are enclosed in parentheses. A Decimal Point
	 * (".")
	 * separates the dollars and cents, and a comma (",") separates
	 * each group
	 * of 3 consecutive digits in the dollar amount.
	 * </p>
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>format</b> The Currency Format
	 *                      </p>
	 *                      <p/>
	 *                      <p/>
	 * @return A new Money object representing the
	 * result
	 * </p>
	 * @see DecimalFormat
	 */
	public Money setCurrencyFormat( DecimalFormat format ) // Currency Format
	{
		Money result = new Money( this ); // Create a new Money object for
		// the result.
		result.currencyFormat = format; // Set the Currency Format.
		return result; // Return the result.
	} // Method Money.setCurrencyFormat()

	/**
	 * Returns a formatted string representation of the monetary
	 * value. The format
	 * of the string is determined by the Currency Format.
	 * <p/>
	 * Refer to the Java API documentation for the DecimalFormat
	 * class for information on
	 * Decimal Formats in general, and Currency Formats in
	 * particular.
	 * <p/>
	 * <p>
	 * By default, the Currency Format for the current locale is
	 * used.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @return A string representation of the
	 * monetary value
	 * </p>
	 * @see DecimalFormat
	 * @see #setCurrencyFormat
	 */
	public String toString()
	{
		// Round off the monetary value to 2 decimal places.
		BigDecimal result = value.setScale( 2, roundingMode );
		return currencyFormat.format( result ); // Format the result according
		// to the Currency Format.
	}

	/**
	 * Parses a string representation of a monetary value, returning
	 * a new Money
	 * object with the specified value. The format of the string
	 * must be consistent
	 * with the Currency Format; otherwise, a ParseException is
	 * recognized.
	 * <p/>
	 * Refer to the Java API documentation for the DecimalFormat
	 * class for information on
	 * Decimal Formats in general, and Currency Formats in
	 * particular.
	 * <p/>
	 * <p>
	 * The value of this object is not affected in any way. A new
	 * object is returned reflecting the result of the operation.
	 * </p>
	 * <p/>
	 * <p/>
	 *
	 * @param <b>string</b> A String representing a monetary
	 *                      value
	 *                      </p>
	 *                      <p/>
	 *                      <p/>
	 * @return A new Money object representating
	 * the parsed monetary value
	 * </p>
	 * <p/>
	 * <p/>
	 * @throws ParseException The string is inconsistent with
	 *                        the Currency Format
	 *                        </p>
	 * @see DecimalFormat
	 * @see #setCurrencyFormat
	 */
	public Money parse( String string ) // String representing a Monetary
	// Value
			throws ParseException // If string is inconsistent with
	// Currency Format
	{
		Money result = new Money( this ); // Create a Money object for the
		// result.
		// We make use of the format's parse() method, which parses the string
		// according to the format and returns either a Double or a Long
		// object representing the value.
		Number number;
		// Attempt to parse the string as a monetary value. May throw a
		// ParseException if the string is not consistent with the Currency
		// Format.
		number = currencyFormat.parse( string );
		// In general, a double floating point value cannot represent a
		// decimal value exactly, and
		// therefore is only a very close approximation of the actual decimal
		// value. Fortunately,
		// the Double.toString() method is able to "recognize" an approximated
		// decimal value, and
		// will return the original (approximated) decimal value, rather than
		// the literal floating
		// point value. Here, we take advantage of this fact to obtain a
		// string representation of
		// the monetary value (without formatting, except for the decimal
		// point), which we then use
		// to create a BigDecimal object having the required value.
		// Were we not to make this simplifying assumption, the only
		// alternative would be to parse the
		// string ourselves, which can be quite complicated, and would
		// unnecessarily duplicate code
		// already implemented by the format's parse() method.
		// Convert the parsed value to a simple string (decimal point only)
		// and use it to set the monetary value.
		result.value = new BigDecimal( number.toString() );
		return result; // Return the new Money object.
	} // Method Money.parse()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is zero (equal to $0.00).
	 * <p/>
	 * <p/>
	 *
	 * @return <b>true </b>  The monetary value is zero
	 * <br>
	 * <b>false</b>  The monetary value is not zero
	 * </p>
	 */
	public boolean isZero()
	{
		return (value.compareTo( ZERO ) == 0); // Return true if value is zero.
	} // Method Money.isZero()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is negative (less than $0.00).
	 * <p/>
	 * <p/>
	 *
	 * @return <b>true </b>  The monetary value is negative
	 * <br>
	 * <b>false</b>  The monetary value is not negative
	 * </p>
	 */
	public boolean isNegative()
	{
		return (value.compareTo( ZERO ) < 0); // Return true if value is less
		// than zero.
	} // Method Money.isNegative()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is positive (greater than or equal to $0.00).
	 * <p/>
	 * <p/>
	 *
	 * @return <b>true </b>  The monetary value is positive
	 * <br>
	 * <b>false</b>  The monetary value is not positive
	 * </p>
	 */
	public boolean isPositive()
	{
		return (value.compareTo( ZERO ) >= 0); // Return true if value is
		// greater than zero.
	} // Method Money.isPositive()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is equal to
	 * another monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>other</b> A monetary value with which this
	 *                     monetary value is to be compared
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return <b>true </b>  This monetary values are equal
	 * <br>
	 * <b>false</b>  This monetary values are not equal
	 * </p>
	 */
	public boolean isEqual( Money other ) // Monetary value for comparison
	{
		return (value.compareTo( other.value ) == 0); // Return true if equal to
		// the other amount.
	} // Method Money.isEqual()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is less than
	 * another monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>other</b> A monetary value with which this
	 *                     monetary value is to be compared
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return <b>true </b>  This monetary value is less than
	 * the specified monetary value
	 * <br>
	 * <b>false</b>  This monetary value is not less
	 * than the specified monetary value
	 * </p>
	 */
	public boolean isLessThan( Money other ) // Monetary value for
	// comparison
	{
		return (value.compareTo( other.value ) < 0); // Return true if less than
		// the other amount.
	} // Method Money.isLessThan()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is less than
	 * or equal to another monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>other</b> A monetary value with which this
	 *                     monetary value is to be compared
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return <b>true </b>  This monetary value is less than
	 * or equal to the specified monetary
	 * value
	 * <br>
	 * <b>false</b>  This monetary value is not less
	 * than or equal to the specified
	 * monetary value
	 * </p>
	 */
	public boolean isLessThanOrEqual( Money other ) // Monetary value for
	// comparison
	{
		// Return true if less than or equal to the other amount.
		return (value.compareTo( other.value ) <= 0);
	} // Method Money.isLessThanOrEqual()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is greater than
	 * another monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>other</b> A monetary value with which this
	 *                     monetary value is to be compared
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return <b>true </b>  This monetary value is greater
	 * than the specified monetary value
	 * <br>
	 * <b>false</b>  This monetary value is not greater
	 * than the specified monetary value
	 * </p>
	 */
	public boolean isGreaterThan( Money other ) // Monetary value for
	// comparison
	{
		// Return true if greater than the other amount.
		return (value.compareTo( other.value ) > 0);
	} // Method Money.isGreaterThan()

	/**
	 * Returns an indication of whether or not this monetary value
	 * is greater than
	 * or equal to another monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>other</b> A monetary value with which this
	 *                     monetary value is to be compared
	 *                     </p>
	 *                     <p/>
	 *                     <p/>
	 * @return <b>true </b>  This monetary value is greater
	 * than or equal to the specified
	 * monetary value
	 * <br>
	 * <b>false</b>  This monetary value is not greater
	 * than or equal to the specified
	 * monetary value
	 * </p>
	 */
	public boolean isGreaterThanOrEqual( Money other ) // Monetary value for
	// comparison
	{
		// Return true if greater than or equal to the other amount.
		return (value.compareTo( other.value ) >= 0);
	} // Method Money.isGreaterThanOrEqual()

	/**
	 * Compares this object with the specified object. The objects
	 * are equal
	 * if and only if the specified object is not null, is a Money
	 * object, and has
	 * the same monetary value as this object.
	 * <p/>
	 * <p/>
	 *
	 * @param <b>object</b> Some object
	 *                      </p>
	 *                      <p/>
	 *                      <p/>
	 * @return <b>true </b>  The objects are equal
	 * <br>
	 * <b>false</b>  The objects are not equal
	 * </p>
	 */
	public boolean equals( Object object ) // Object to compare
	{
		if( object == this ) // If the object is this object,
		{
			return true; // the objects are equal by
		}
		// definition.
		if( object == null ) // If the object is null,
		{
			return false; // the objects are not equal by
		}
		// definition.
		if( !(object instanceof Money) ) // If the object is not an instance
		// of Money,
		{
			return false; // the objects are not equal by
		}
		// definition.
		// Return true if monetary values are the same.
		return (value.compareTo( ((Money) object).value ) == 0);
	} // Method Money.equals()

	/**
	 * Returns a hashcode for this object. The hashcode is identical
	 * to that for the BigDecimal
	 * object that represents the monetary value.
	 * <p/>
	 * <p/>
	 *
	 * @return The hashcode
	 * </p>
	 */
	public int hashCode()
	{
		return value.hashCode(); // Return the hashcode for the
	}

	/**
	 * Clones a Money object. The new object is an exact copy of
	 * this object,  and inherits the object's monetary value and Currency Format.
	 * <p/>
	 * <p/>
	 *
	 * @return <b>Money</b> The cloned object
	 * </p>
	 */
	public Object clone()
	{
		Money result = new Money( this ); // Create a copy of this Money object.
		return result; // Return the cloned object.
	}
} // Class Money
