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

package cn.ttzero.excel.tmap;

abstract public class TIntIntHash extends TPrimitiveHash {

    /** the set of ints */
    public transient int[] _set;


    /**
     * key that represents null
     *
     * NOTE: should not be modified after the Hash is created, but is
     *       not final because of Externalization
     *
     */
    protected int no_entry_key;


    /**
     * value that represents null
     *
     * NOTE: should not be modified after the Hash is created, but is
     *       not final because of Externalization
     *
     */
    protected int no_entry_value;

    protected boolean consumeFreeSlot;

    /**
     * Creates a new <code>T#E#Hash</code> instance with the default
     * capacity and load factor.
     */
    public TIntIntHash() {
        super();
        no_entry_key = 0;
        no_entry_value = 0;
    }


    /**
     * Creates a new <code>T#E#Hash</code> instance whose capacity
     * is the next highest prime above <tt>initialCapacity + 1</tt>
     * unless that value is already prime.
     *
     * @param initialCapacity an <code>int</code> value
     */
    public TIntIntHash( int initialCapacity ) {
        super( initialCapacity );
        no_entry_key = 0;
        no_entry_value = 0;
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
        capacity = super.setUp( initialCapacity );
        _set = new int[capacity];
        return capacity;
    }

    /**
     * Locates the index of <tt>val</tt>.
     *
     * @param key an <code>int</code> value
     * @return the index of <tt>val</tt> or -1 if it isn't in the set.
     */
    protected int index( int key ) {
        int hash, index, length;

        final byte[] states = _states;
        final int[] set = _set;
        length = states.length;
        hash = key & 0x7fffffff;
        index = hash % length;
        byte state = states[index];

        if (state == FREE)
            return -1;

        if (state == FULL && set[index] == key)
            return index;

        return indexRehashed(key, index, hash, state);
    }

    int indexRehashed(int key, int index, int hash, byte state) {
        // see Knuth, p. 529
        int length = _set.length;
        int probe = 1 + (hash % (length - 2));
        final int loopIndex = index;

        do {
            index -= probe;
            if (index < 0) {
                index += length;
            }
            state = _states[index];
            //
            if (state == FREE)
                return -1;

            //
            if (key == _set[index] && state != REMOVED)
                return index;
        } while (index != loopIndex);

        return -1;
    }


    /**
     * Locates the index at which <tt>val</tt> can be inserted.  if
     * there is already a value equal()ing <tt>val</tt> in the set,
     * returns that value as a negative integer.
     *
     * @return an <code>int</code> value
     */
    protected int insertKey( int val ) {
        int hash, index;

        hash = val & 0x7fffffff;
        index = hash % _states.length;
        byte state = _states[index];

        consumeFreeSlot = false;

        if (state == FREE) {
            consumeFreeSlot = true;
            insertKeyAt(index, val);

            return index;       // empty, all done
        }

        if (state == FULL && _set[index] == val) {
            return -index - 1;   // already stored
        }

        // already FULL or REMOVED, must probe
        return insertKeyRehash(val, index, hash, state);
    }

    int insertKeyRehash(int val, int index, int hash, byte state) {
        // compute the double hash
        final int length = _set.length;
        int probe = 1 + (hash % (length - 2));
        final int loopIndex = index;
        int firstRemoved = -1;

        /**
         * Look until FREE slot or we start to loop
         */
        do {
            // Identify first removed slot
            if (state == REMOVED && firstRemoved == -1)
                firstRemoved = index;

            index -= probe;
            if (index < 0) {
                index += length;
            }
            state = _states[index];

            // A FREE slot stops the search
            if (state == FREE) {
                if (firstRemoved != -1) {
                    insertKeyAt(firstRemoved, val);
                    return firstRemoved;
                } else {
                    consumeFreeSlot = true;
                    insertKeyAt(index, val);
                    return index;
                }
            }

            if (state == FULL && _set[index] == val) {
                return -index - 1;
            }

            // Detect loop
        } while (index != loopIndex);

        // We inspected all reachable slots and did not find a FREE one
        // If we found a REMOVED slot we return the first one found
        if (firstRemoved != -1) {
            insertKeyAt(firstRemoved, val);
            return firstRemoved;
        }

        // Can a resizing strategy be found that resizes the set?
        throw new IllegalStateException("No free or removed slots available. Key set full?!!");
    }

    void insertKeyAt(int index, int val) {
        _set[index] = val;  // insert value
        _states[index] = FULL;
    }
} // TIntIntHash