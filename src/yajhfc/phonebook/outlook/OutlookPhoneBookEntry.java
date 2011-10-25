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

import com.jacob.com.Dispatch;
import com.jacobgen.ms.outlook._ContactItem;

public class OutlookPhoneBookEntry extends SimplePhoneBookEntry {
	
	
	protected final OutlookPhonebook parent;
	protected final String suffix;

	protected OutlookPhoneBookEntry(OutlookPhonebook parent,
			_ContactItem contactItem, Map<PBEntryField,String> propertyMap, String suffix) {
		super();
		this.parent = parent;
		this.suffix = suffix;
		readContact(contactItem, propertyMap);
	}

	protected void readContact(_ContactItem contact, Map<PBEntryField,String> propertyMap) {
		//System.out.println(contact.getFullName() + ": " + contact.getEntryID());
		for (PBEntryField field : PBEntryField.values()) {
			String olProp = propertyMap.get(field);
			if (olProp != null) {
				setFieldUndirty(field, Dispatch.get(contact, olProp).toString());
			} else {
				setFieldUndirty(field, "");
			}
		}
		setDirty(false);
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
