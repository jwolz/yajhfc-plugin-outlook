package yajhfc.phonebook.outlook;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import yajhfc.Utils;
import yajhfc.phonebook.DefaultPhoneBookEntry;
import yajhfc.phonebook.DistributionList;
import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.PhoneBookEntry;
import yajhfc.phonebook.PhonebookEventListener;
import yajhfc.phonebook.convrules.PBEntryFieldContainer;

import com.jacobgen.ms.outlook.Recipient;
import com.jacobgen.ms.outlook._ContactItem;
import com.jacobgen.ms.outlook._DistListItem;

public class OutlookDistList extends DefaultPhoneBookEntry implements
		DistributionList {
	private static final Logger log = Logger.getLogger(OutlookDistList.class.getName());
	
	protected final OutlookPhonebook parent;
	protected String dlName;
	protected List<PhoneBookEntry> entries = new ArrayList<PhoneBookEntry>();
	protected List<PhoneBookEntry> entryView = Collections.<PhoneBookEntry>unmodifiableList(entries);
	
	public OutlookDistList(OutlookPhonebook parent, _DistListItem distList) {
		super();
		this.parent = parent;
		loadFromDistListItem(distList);
	}
	
	
	protected void loadFromDistListItem(_DistListItem distList) {
		dlName = distList.getDLName();
		if (Utils.debugMode) {
			log.fine("Distribution list name: " + dlName);
		}
		for (int i=1; i<=distList.getMemberCount(); i++) {
			Recipient rec = distList.getMember(i);
//			System.out.println("Name=" +rec.getName() + "; Address=" + rec.getAddress() + "; DisplayType=" + rec.getDisplayType() + "; Type=" + rec.getType());
			
			if (Utils.debugMode) {
				log.fine(dlName + ": recipient #" + i + ": " + rec.getName() + ": " + rec.getAddress());
			}

			if (parent.settings.resolveDistributionLists && rec.resolve()) {
				String entryID = rec.getEntryID();
				if (Utils.debugMode) {
					log.fine(dlName + ": recipient #" + i + " resolves; ID=" + entryID);
				}

				if (entryID.length() >=48) {
					String ciID = entryID.substring(entryID.length()-48);
					try {
						_ContactItem ci = new _ContactItem(parent.nameSpace.getItemFromID(ciID).toDispatch());
						// Add contact item
						entries.add(createEntryForContactAndRecipient(ci, rec));
					} catch (Exception e) {
						log.log(Level.WARNING, "Resolution failed for ID " + ciID + " (recipient ID was: " + entryID + ")", e);
						entries.add(new OlRecipientPhoneBookEntry(parent, rec));
					}
				} else {
					if (Utils.debugMode) {
						log.fine(dlName + ": recipient #" + i + " ID is too short");
					}
					entries.add(new OlRecipientPhoneBookEntry(parent, rec));
				}
			} else {
				if (Utils.debugMode) {
					log.fine(dlName + ": recipient #" + i + " does not resolve");
				}
				entries.add(new OlRecipientPhoneBookEntry(parent, rec));
			}
		}
	}
	
	protected PhoneBookEntry createEntryForContactAndRecipient(_ContactItem ci, Recipient rec) {
		// !Important: Also update logic in OlRecipientPhoneBookEntry when you change this!
		String address = rec.getAddress();
		int atPos = address.indexOf('@');
		if (atPos >= 0 && atPos < address.length() - 1) {
			String domain = address.substring(atPos+1);
			if (domain.startsWith("+")) {
				// If the domain starts with a + (as in +49 12345), assume it is a fax number
				if (domain.endsWith(ci.getBusinessFaxNumber())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Business, null);
				} else if (domain.endsWith(ci.getHomeFaxNumber())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Home, null);
				} else if (domain.endsWith(ci.getOtherFaxNumber())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Other, null);
				} else {
					log.info("Could not find fax number '" + domain + "', using no mapping.");
					return new OlRecipientPhoneBookEntry(parent, rec);
				}
			} else {
				// Assume it is a mail address
				// If the domain starts with a + (as in +49 12345), assume it is a fax number
				if (address.endsWith(ci.getEmail1Address())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Business, null);
				} else if (address.endsWith(ci.getEmail2Address())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Home, null);
				} else if (address.endsWith(ci.getEmail3Address())) {
					return new OutlookPhoneBookEntry(parent, ci, parent.propertyMap_Other, null);
				} else {
					log.info("Could not find email '" + address + "', using no mapping.");
					return new OlRecipientPhoneBookEntry(parent, rec);
				}
			}
		} else { // Something unknown..
			log.info("Unknown address in recipient, using no mapping.");
			return new OlRecipientPhoneBookEntry(parent, rec);
		}
	}
	
	@Override
	public List<PhoneBookEntry> getEntries() {
		return entryView;
	}

	@Override
	public void addEntries(Collection<? extends PBEntryFieldContainer> items) {
		// Read only
	}

	@Override
	public PhoneBookEntry addNewEntry() {
		// Read only
		return null;
	}

	@Override
	public PhoneBookEntry addNewEntry(PBEntryFieldContainer item) {
		// Read only
		return null;
	}

	@Override
	public void addPhonebookEventListener(PhonebookEventListener pel) {
		// Not necessary because this list is static
	}

	@Override
	public void removePhonebookEventListener(PhonebookEventListener pel) {
		// Not necessary because this list is static
	}

	@Override
	public boolean isReadOnly() {
		return true;
	}

	@Override
	public PhoneBook getParent() {
		return parent;
	}

	@Override
	public String getField(PBEntryField field) {
		switch (field) {
		case Name:
			return dlName;
		default:
			return null;
		}
	}

	@Override
	public void setField(PBEntryField field, String value) {
		// Read only
	}

	@Override
	public void delete() {
		// Read only
	}

	@Override
	public void commit() {
		// Read only
	}

}
