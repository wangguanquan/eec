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

package net.cua.export.tmap;


import java.util.Arrays;

/**
 * An open addressed Map implementation for int keys and int values.
 *
 * @author Eric D. Friedman
 * @author Rob Eden
 * @author Jeff Randall
 * @version $Id: _K__V_HashMap.template,v 1.1.2.16 2010/03/02 04:09:50 robeden Exp $
 */
public class TIntIntHashMap extends TIntIntHash {
//    static final long serialVersionUID = 1L;

    /** the values of the map */
    protected transient int[] _values;


    /**
     * Creates a new <code>TIntIntHashMap</code> instance with the default
     * capacity and load factor.
     */
    public TIntIntHashMap() {
        super();
    }


    /**
     * Creates a new <code>TIntIntHashMap</code> instance with a prime
     * capacity equal to or greater than <tt>initialCapacity</tt> and
     * with the default load factor.
     *
     * @param initialCapacity an <code>int</code> value
     */
    public TIntIntHashMap( int initialCapacity ) {
        super( initialCapacity );
    }

//
//    /**
//     * Creates a new <code>TIntIntHashMap</code> instance with a prime
//     * capacity equal to or greater than <tt>initialCapacity</tt> and
//     * with the specified load factor.
//     *
//     * @param initialCapacity an <code>int</code> value
//     * @param loadFactor a <code>float</code> value
//     */
//    public TIntIntHashMap( int initialCapacity, float loadFactor ) {
//        super( initialCapacity, loadFactor );
//    }
//
//
//    /**
//     * Creates a new <code>TIntIntHashMap</code> instance with a prime
//     * capacity equal to or greater than <tt>initialCapacity</tt> and
//     * with the specified load factor.
//     *
//     * @param initialCapacity an <code>int</code> value
//     * @param loadFactor a <code>float</code> value
//     * @param noEntryKey a <code>int</code> value that represents
//     *                   <tt>null</tt> for the Key set.
//     * @param noEntryValue a <code>int</code> value that represents
//     *                   <tt>null</tt> for the Value set.
//     */
//    public TIntIntHashMap( int initialCapacity, float loadFactor,
//                           int noEntryKey, int noEntryValue ) {
//        super( initialCapacity, loadFactor, noEntryKey, noEntryValue );
//    }
//
//
//    /**
//     * Creates a new <code>TIntIntHashMap</code> instance containing
//     * all of the entries in the map passed in.
//     *
//     * @param keys a <tt>int</tt> array containing the keys for the matching values.
//     * @param values a <tt>int</tt> array containing the values.
//     */
//    public TIntIntHashMap( int[] keys, int[] values ) {
//        super( Math.max( keys.length, values.length ) );
//
//        int size = Math.min( keys.length, values.length );
//        for ( int i = 0; i < size; i++ ) {
//            this.put( keys[i], values[i] );
//        }
//    }


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
        _values = new int[capacity];
        return capacity;
    }


    /**
     * rehashes the map to the new capacity.
     *
     * @param newCapacity an <code>int</code> value
     */
    /** {@inheritDoc} */
    protected void rehash( int newCapacity ) {
        int oldCapacity = _set.length;

        int oldKeys[] = _set;
        int oldVals[] = _values;
        byte oldStates[] = _states;

        _set = new int[newCapacity];
        _values = new int[newCapacity];
        _states = new byte[newCapacity];

        for ( int i = oldCapacity; i-- > 0; ) {
            if( oldStates[i] == FULL ) {
                int o = oldKeys[i];
                int index = insertKey( o );
                _values[index] = oldVals[i];
            }
        }
    }


    /** {@inheritDoc} */
    public int put( int key, int value ) {
        int index = insertKey( key );
        return doPut( key, value, index );
    }


//    /** {@inheritDoc} */
//    public int putIfAbsent( int key, int value ) {
//        int index = insertKey( key );
//        if (index < 0)
//            return _values[-index - 1];
//        return doPut( key, value, index );
//    }


    private int doPut( int key, int value, int index ) {
        int previous = no_entry_value;
        boolean isNewMapping = true;
        if ( index < 0 ) {
            index = -index -1;
            previous = _values[index];
            isNewMapping = false;
        }
        _values[index] = value;

        if (isNewMapping) {
            postInsertHook( consumeFreeSlot );
        }

        return previous;
    }


//    /** {@inheritDoc} */
//    public void putAll( Map<? extends Integer, ? extends Integer> map ) {
//        ensureCapacity( map.size() );
//        // could optimize this for cases when map instanceof THashMap
//        for ( Map.Entry<? extends Integer, ? extends Integer> entry : map.entrySet() ) {
//            this.put( entry.getKey().intValue(), entry.getValue().intValue() );
//        }
//    }


    /** {@inheritDoc} */
    public int get( int key ) {
        int index = index( key );
        return index < 0 ? no_entry_value : _values[index];
    }


    /** {@inheritDoc} */
    public void clear() {
        super.clear();
        Arrays.fill( _set, 0, _set.length, no_entry_key );
        Arrays.fill( _values, 0, _values.length, no_entry_value );
        Arrays.fill( _states, 0, _states.length, FREE );
    }


    /** {@inheritDoc} */
//    public boolean isEmpty() {
//        return 0 == _size;
//    }
//
//
//    /** {@inheritDoc} */
//    public int remove( int key ) {
//        int prev = no_entry_value;
//        int index = index( key );
//        if ( index >= 0 ) {
//            prev = _values[index];
//            removeAt( index );    // clear key,state; adjust size
//        }
//        return prev;
//    }
//
//
//    /** {@inheritDoc} */
//    protected void removeAt( int index ) {
//        _values[index] = no_entry_value;
//        super.removeAt( index );  // clear key, state; adjust size
//    }


    /** {@inheritDoc} */
    public int[] keys() {
        int[] keys = new int[size()];
        int[] k = _set;
        byte[] states = _states;

        for ( int i = k.length, j = 0; i-- > 0; ) {
            if ( states[i] == FULL ) {
                keys[j++] = k[i];
            }
        }
        return keys;
    }


//    /** {@inheritDoc} */
//    public int[] keys( int[] array ) {
//        int size = size();
//        if ( array.length < size ) {
//            array = new int[size];
//        }
//
//        int[] keys = _set;
//        byte[] states = _states;
//
//        for ( int i = keys.length, j = 0; i-- > 0; ) {
//            if ( states[i] == FULL ) {
//                array[j++] = keys[i];
//            }
//        }
//        return array;
//    }


    /** {@inheritDoc} */
    public int[] values() {
        int[] vals = new int[size()];
        int[] v = _values;
        byte[] states = _states;

        for ( int i = v.length, j = 0; i-- > 0; ) {
            if ( states[i] == FULL ) {
                vals[j++] = v[i];
            }
        }
        return vals;
    }


//    /** {@inheritDoc} */
//    public int[] values( int[] array ) {
//        int size = size();
//        if ( array.length < size ) {
//            array = new int[size];
//        }
//
//        int[] v = _values;
//        byte[] states = _states;
//
//        for ( int i = v.length, j = 0; i-- > 0; ) {
//            if ( states[i] == FULL ) {
//                array[j++] = v[i];
//            }
//        }
//        return array;
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean containsValue( int val ) {
//        byte[] states = _states;
//        int[] vals = _values;
//
//        for ( int i = vals.length; i-- > 0; ) {
//            if ( states[i] == FULL && val == vals[i] ) {
//                return true;
//            }
//        }
//        return false;
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean containsKey( int key ) {
//        return contains( key );
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean forEachKey( TIntProcedure procedure ) {
//        return forEach( procedure );
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean forEachValue( TIntProcedure procedure ) {
//        byte[] states = _states;
//        int[] values = _values;
//        for ( int i = values.length; i-- > 0; ) {
//            if ( states[i] == FULL && ! procedure.execute( values[i] ) ) {
//                return false;
//            }
//        }
//        return true;
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean forEachEntry( TIntIntProcedure procedure ) {
//        byte[] states = _states;
//        int[] keys = _set;
//        int[] values = _values;
//        for ( int i = keys.length; i-- > 0; ) {
//            if ( states[i] == FULL && ! procedure.execute( keys[i], values[i] ) ) {
//                return false;
//            }
//        }
//        return true;
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean increment( int key ) {
//        return adjustValue( key, 1 );
//    }
//
//
//    /** {@inheritDoc} */
//    public boolean adjustValue( int key, int amount ) {
//        int index = index( key );
//        if (index < 0) {
//            return false;
//        } else {
//            _values[index] += amount;
//            return true;
//        }
//    }


    /** {@inheritDoc} */
    public int adjustOrPutOne( int key ) {
        int index = insertKey( key );
        final boolean isNewMapping;
        final int newValue;
        if ( index < 0 ) {
            index = -index -1;
            newValue = ( _values[index] += 1 );
            isNewMapping = false;
        } else {
            newValue = ( _values[index] = 1 );
            isNewMapping = true;
        }

//        byte previousState = _states[index];

        if ( isNewMapping ) {
            postInsertHook(consumeFreeSlot);
        }

        return newValue;
    }


//    /** {@inheritDoc} */
//    @Override
//    public boolean equals( Object other ) {
//        if ( ! ( other instanceof TIntIntHashMap ) ) {
//            return false;
//        }
//        TIntIntHashMap that = ( TIntIntHashMap ) other;
//        if ( that.size() != this.size() ) {
//            return false;
//        }
//        int[] values = _values;
//        byte[] states = _states;
//        int this_no_entry_value = getNoEntryValue();
//        int that_no_entry_value = that.getNoEntryValue();
//        for ( int i = values.length; i-- > 0; ) {
//            if ( states[i] == FULL ) {
//                int key = _set[i];
//                int that_value = that.get( key );
//                int this_value = values[i];
//                if ( ( this_value != that_value ) &&
//                        ( this_value != this_no_entry_value ) &&
//                        ( that_value != that_no_entry_value ) ) {
//                    return false;
//                }
//            }
//        }
//        return true;
//    }
//
//
//    /** {@inheritDoc} */
//    @Override
//    public int hashCode() {
//        int hashcode = 0;
//        byte[] states = _states;
//        for ( int i = _values.length; i-- > 0; ) {
//            if ( states[i] == FULL ) {
//                hashcode += _set[i] ^ _values[i];
//            }
//        }
//        return hashcode;
//    }
//
//
//    /** {@inheritDoc} */
//    @Override
//    public String toString() {
//        final StringBuilder buf = new StringBuilder( "{" );
//        forEachEntry( new TIntIntProcedure() {
//            private boolean first = true;
//            public boolean execute( int key, int value ) {
//                if ( first ) first = false;
//                else buf.append( ", " );
//
//                buf.append(key);
//                buf.append("=");
//                buf.append(value);
//                return true;
//            }
//        });
//        buf.append( "}" );
//        return buf.toString();
//    }
//
//
//    /** {@inheritDoc} */
//    public void writeExternal(ObjectOutput out) throws IOException {
//        // VERSION
//        out.writeByte( 0 );
//
//        // SUPER
//        super.writeExternal( out );
//
//        // NUMBER OF ENTRIES
//        out.writeInt( _size );
//
//        // ENTRIES
//        for ( int i = _states.length; i-- > 0; ) {
//            if ( _states[i] == FULL ) {
//                out.writeInt( _set[i] );
//                out.writeInt( _values[i] );
//            }
//        }
//    }
//
//
//    /** {@inheritDoc} */
//    public void readExternal(ObjectInput in) throws IOException, ClassNotFoundException {
//        // VERSION
//        in.readByte();
//
//        // SUPER
//        super.readExternal( in );
//
//        // NUMBER OF ENTRIES
//        int size = in.readInt();
//        setUp( size );
//
//        // ENTRIES
//        while (size-- > 0) {
//            int key = in.readInt();
//            int val = in.readInt();
//            put(key, val);
//        }
//    }
} // TIntIntHashMap