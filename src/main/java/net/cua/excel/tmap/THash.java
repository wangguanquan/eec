///////////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001, Eric D. Friedman All Rights Reserved.
// Copyright (c) 2009, Rob Eden All Rights Reserved.
// Copyright (c) 2009, Jeff Randall All Rights Reserved.
//
// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either
// version 2.1 of the License, or (at your option) any later version.
//
// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public
// License along with this program; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
///////////////////////////////////////////////////////////////////////////////

package net.cua.excel.tmap;

/**
        * Base class for hashtables that use open addressing to resolve
        * collisions.
        *
        * Created: Wed Nov 28 21:11:16 2001
        *
        * @author Eric D. Friedman
        * @author Rob Eden (auto-compaction)
        * @author Jeff Randall
        *
        * @version $Id: THash.java,v 1.1.2.4 2010/03/02 00:55:34 robeden Exp $
        */
abstract public class THash {

    /** the load above which rehashing occurs. */
    protected static final float DEFAULT_LOAD_FACTOR = 0.85f;

    /**
     * the default initial capacity for the hash table.  This is one
     * less than a prime value because one is added to it when
     * searching for a prime capacity to account for the free slot
     * required by open addressing. Thus, the real default capacity is
     * 11.
     */
    protected static final int DEFAULT_CAPACITY = 10;


    /** the current number of occupied slots in the hash. */
    protected transient int _size;

    /** the current number of free slots in the hash. */
    protected transient int _free;

    /**
     * Determines how full the internal table can become before
     * rehashing is required. This must be a value in the range: 0.0 <
     * loadFactor < 1.0.  The default value is 0.5, which is about as
     * large as you can get in open addressing without hurting
     * performance.  Cf. Knuth, Volume 3., Chapter 6.
     */
    protected float _loadFactor;

    /**
     * The maximum number of elements allowed without allocating more
     * space.
     */
    protected int _maxSize;


    /** The number of removes that should be performed before an auto-compaction occurs. */
    protected int _autoCompactRemovesRemaining;

    /**
     * The auto-compaction factor for the table.
     *
     */
    protected float _autoCompactionFactor;


    /**
     * Creates a new <code>THash</code> instance with the default
     * capacity and load factor.
     */
    public THash() {
        this( DEFAULT_CAPACITY, DEFAULT_LOAD_FACTOR );
    }


    /**
     * Creates a new <code>THash</code> instance with a prime capacity
     * at or near the specified capacity and with the default load
     * factor.
     *
     * @param initialCapacity an <code>int</code> value
     */
    public THash( int initialCapacity ) {
        this( initialCapacity, DEFAULT_LOAD_FACTOR );
    }


    /**
     * Creates a new <code>THash</code> instance with a prime capacity
     * at or near the minimum needed to hold <tt>initialCapacity</tt>
     * elements with load factor <tt>loadFactor</tt> without triggering
     * a rehash.
     *
     * @param initialCapacity an <code>int</code> value
     * @param loadFactor      a <code>float</code> value
     */
    public THash( int initialCapacity, float loadFactor ) {
        super();
        _loadFactor = loadFactor;

        // Through testing, the load factor (especially the default load factor) has been
        // found to be a pretty good starting auto-compaction factor.
        _autoCompactionFactor = loadFactor;

        setUp( fastCeil( initialCapacity / loadFactor ) );
    }

    public static int fastCeil( float v ) {
        int possible_result = ( int ) v;
        if ( v - possible_result > 0 ) possible_result++;
        return possible_result;
    }

    /**
     * Tells whether this set is currently holding any elements.
     *
     * @return a <code>boolean</code> value
     */
    public boolean isEmpty() {
        return 0 == _size;
    }


    /**
     * Returns the number of distinct elements in this collection.
     *
     * @return an <code>int</code> value
     */
    public int size() {
        return _size;
    }


    /** @return the current physical capacity of the hash table. */
    abstract public int capacity();

    /** Empties the collection. */
    public void clear() {
        _size = 0;
        _free = capacity();
    }


    /**
     * initializes the hashtable to a prime capacity which is at least
     * <tt>initialCapacity + 1</tt>.
     *
     * @param initialCapacity an <code>int</code> value
     * @return the actual capacity chosen
     */
    protected int setUp( int initialCapacity ) {
        int capacity;

        capacity = PrimeFinder.nextPrime( initialCapacity );
        computeMaxSize( capacity );
        computeNextAutoCompactionAmount( initialCapacity );

        return capacity;
    }


    /**
     * Rehashes the set.
     *
     * @param newCapacity an <code>int</code> value
     */
    protected abstract void rehash( int newCapacity );


    /**
     * Computes the values of maxSize. There will always be at least
     * one free slot required.
     *
     * @param capacity an <code>int</code> value
     */
    protected void computeMaxSize( int capacity ) {
        // need at least one free slot for open addressing
        _maxSize = Math.min( capacity - 1, (int) ( capacity * _loadFactor ) );
        _free = capacity - _size; // reset the free element count
    }


    /**
     * Computes the number of removes that need to happen before the next auto-compaction
     * will occur.
     *
     * @param size an <tt>int</tt> that sets the auto-compaction limit.
     */
    protected void computeNextAutoCompactionAmount( int size ) {
        if ( _autoCompactionFactor != 0 ) {
            // NOTE: doing the round ourselves has been found to be faster than using
            //       Math.round.
            _autoCompactRemovesRemaining =
                    (int) ( ( size * _autoCompactionFactor ) + 0.5f );
        }
    }


    /**
     * After an insert, this hook is called to adjust the size/free
     * values of the set and to perform rehashing if necessary.
     *
     * @param usedFreeSlot the slot
     */
    protected final void postInsertHook( boolean usedFreeSlot ) {
        if ( usedFreeSlot ) {
            _free--;
        }

        // rehash whenever we exhaust the available space in the table
        if ( ++_size > _maxSize || _free == 0 ) {
            // choose a new capacity suited to the new state of the table
            // if we've grown beyond our maximum size, double capacity;
            // if we've exhausted the free spots, rehash to the same capacity,
            // which will free up any stale removed slots for reuse.
            int newCapacity = _size > _maxSize ? PrimeFinder.nextPrime( capacity() << 1 ) : capacity();
            rehash( newCapacity );
            computeMaxSize( capacity() );
        }
    }

}// THash