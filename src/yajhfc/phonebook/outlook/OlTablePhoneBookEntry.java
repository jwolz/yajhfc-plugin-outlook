package yajhfc.phonebook.outlook;
/*
 * YAJHFC - Yet another Java Hylafax client
 * Copyright (C) 2011 Jonas Wolz <info@yajhfc.de>
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License as published by
 *  the Free Software Foundation, either version 3 of the License, or
 *  (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
import java.util.Map;

import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.SimplePhoneBookEntry;

import com.jacob.com.SafeArray;

public class OlTablePhoneBookEntry extends SimplePhoneBookEntry {	
	protected final OutlookPhonebook parent;
	protected String suffix;

	
	protected OlTablePhoneBookEntry(OutlookPhonebook parent, SafeArray table, int row, Map<PBEntryField,Integer> indexMap) {
		super();
		this.parent = parent;
		loadEntry(table, row, indexMap);
	}
	
	protected void loadEntry(SafeArray table, int row, Map<PBEntryField,Integer> indexMap) {
		for (PBEntryField field : indexMap.keySet()) {
			int col = indexMap.get(field);
			String value = table.getString(row, col);
			
			setFieldUndirty(field, value);
		}
		setDirty(false);
	}
	
	public String getSuffix() {
		return suffix;
	}
	
	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}
	
	public boolean hasAddress() {
		for (PBEntryField field : parent.addressFields) {
			String s = getField(field);
			if (s != null && s.length() > 0) {
				return true;
			}
		}
		return false;
	}
	
	
	@Override
	public PhoneBook getParent() {
		return parent;
	}
	
	@Override
	public String toString() {
		if (suffix != null) {
			return super.toString() + " (" + suffix + ")";
		} else {
			return super.toString();
		}
	}

	@Override
	public void delete() {
		// Not modifiable
	}

	@Override
	public void commit() {
		//  Not modifiable
	}

	

}
