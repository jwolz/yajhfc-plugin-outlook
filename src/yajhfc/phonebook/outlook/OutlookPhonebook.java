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
import static yajhfc.phonebook.outlook.EntryPoint._;

import java.awt.Dialog;
import java.util.ArrayList;
import java.util.Collections;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import yajhfc.Utils;
import yajhfc.phonebook.PBEntryField;
import yajhfc.phonebook.PhoneBook;
import yajhfc.phonebook.PhoneBookEntry;
import yajhfc.phonebook.PhoneBookException;
import yajhfc.util.ExceptionDialog;

import com.jacob.com.Dispatch;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;
import com.jacobgen.ms.outlook.Columns;
import com.jacobgen.ms.outlook.MAPIFolder;
import com.jacobgen.ms.outlook.OlObjectClass;
import com.jacobgen.ms.outlook.Table;
import com.jacobgen.ms.outlook._Application;
import com.jacobgen.ms.outlook._ContactItem;
import com.jacobgen.ms.outlook._DistListItem;
import com.jacobgen.ms.outlook._Items;
import com.jacobgen.ms.outlook._NameSpace;

public class OutlookPhonebook extends PhoneBook {
	
	private static final Logger log = Logger.getLogger(OutlookPhonebook.class.getName());
	protected static final Object outlookLock = new Object();
	
	public static String PB_Prefix = "outlook";      // The prefix of this Phonebook type's descriptor
	public static String PB_DisplayName = _("Microsoft Outlook contacts"); // A user-readable name for this Phonebook type
	public static String PB_Description = _("Phone book reading it entries from Microsoft Outlook contacts"); // A user-readable description of this Phonebook type
	
	/**
	 * Which fields must be present for an address to be loaded
	 */
	protected PBEntryField[] addressFields; 
	
	/*
	 * Mappings PBEntryField <-> Outlook property name
	 */
	protected final Map<PBEntryField,String> propertyMap_Home  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	protected final Map<PBEntryField,String> propertyMap_Business  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	protected final Map<PBEntryField,String> propertyMap_Other  = new EnumMap<PBEntryField,String>(PBEntryField.class);
	
	protected List<PhoneBookEntry> entries = new ArrayList<PhoneBookEntry>();
	protected List<PhoneBookEntry> entryView = Collections.<PhoneBookEntry>unmodifiableList(entries);
	
	protected OutlookSettings settings;
	protected boolean open;
	protected String folderName;

	protected _NameSpace nameSpace;
	protected _Application app;
	
	private boolean useOl2007Api;
	
	public OutlookPhonebook(Dialog parent) {
		super(parent);
	}
	
	
	protected void loadBusinessAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody && !useOl2007Api)
			propertyMap.put(PBEntryField.Comment, "Body");
		propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "BusinessAddressCountry");
		propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email1Address");
		propertyMap.put(PBEntryField.FaxNumber, "BusinessFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "BusinessAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "BusinessAddressState");
		propertyMap.put(PBEntryField.Street, "BusinessAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "BusinessTelephoneNumber");
		propertyMap.put(PBEntryField.WebSite, "BusinessHomePage");
		propertyMap.put(PBEntryField.ZIPCode, "BusinessAddressPostalCode");
	}
	
	protected void loadHomeAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody && !useOl2007Api)
			propertyMap.put(PBEntryField.Comment, "Body");
		//propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "HomeAddressCountry");
		//propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email2Address");
		propertyMap.put(PBEntryField.FaxNumber, "HomeFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "HomeAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		//propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "HomeAddressState");
		propertyMap.put(PBEntryField.Street, "HomeAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "HomeTelephoneNumber");
		propertyMap.put(PBEntryField.WebSite, "WebPage");
		propertyMap.put(PBEntryField.ZIPCode, "HomeAddressPostalCode");
	}

	protected void loadOtherAddressMapping(Map<PBEntryField,String> propertyMap) {
		propertyMap.clear();
		
		if (settings.accessEMailAndBody && !useOl2007Api)
			propertyMap.put(PBEntryField.Comment, "Body");
		//propertyMap.put(PBEntryField.Company, "CompanyName");
		propertyMap.put(PBEntryField.Country, "OtherAddressCountry");
		//propertyMap.put(PBEntryField.Department, "Department");
		if (settings.accessEMailAndBody)
			propertyMap.put(PBEntryField.EMailAddress, "Email3Address");
		propertyMap.put(PBEntryField.FaxNumber, "OtherFaxNumber");
		propertyMap.put(PBEntryField.GivenName, "FirstName");
		propertyMap.put(PBEntryField.Location, "OtherAddressCity");
		propertyMap.put(PBEntryField.Name, "LastName");
		//propertyMap.put(PBEntryField.Position, "JobTitle"); 
		propertyMap.put(PBEntryField.State, "OtherAddressState");
		propertyMap.put(PBEntryField.Street, "OtherAddressStreet");
		propertyMap.put(PBEntryField.Title, "Title");
		propertyMap.put(PBEntryField.VoiceNumber, "OtherTelephoneNumber");
		//propertyMap.put(PBEntryField.WebSite, "OtherHomePage");
		propertyMap.put(PBEntryField.ZIPCode, "OtherAddressPostalCode");
	}
	
	@Override
	public PhoneBookEntry addNewEntry() {
		return null;
	}

	@Override
	public List<PhoneBookEntry> getEntries() {
		return entryView;
	}

	@Override
	public String browseForPhoneBook(boolean exportMode) {
		try {
			OutlookSettings newSettings = new OutlookSettings();
			if (settings != null)
				newSettings.copyFrom(settings);
			ConnectionDialog cd = new ConnectionDialog(parentDialog);
			if (cd.promptForNewSettings(newSettings)) {
				return PB_Prefix + ":" + newSettings.saveToString();
			} else {
				return null;
			}
		} catch (Exception e) {
			ExceptionDialog.showExceptionDialog(parentDialog, _("Cannot connect to Outlook"), e);
			return null;
		}
	}

	@Override
	public boolean isOpen() {
		return open;
	}

	@Override
	public String getDisplayCaption() {
		return "Outlook - " + folderName;
	}
	
	@Override
	public void resort() {
		Collections.sort(entries);
	}

	@Override
	protected void openInternal(String descriptorWithoutPrefix)
			throws PhoneBookException {
		log.fine("Loading settings...");
		settings = new OutlookSettings();
		settings.loadFromString(descriptorWithoutPrefix);
		
		if (settings.loadOnlyFaxContacts) {
			addressFields = new PBEntryField[] {
					PBEntryField.FaxNumber,
			};
		} else {
			addressFields = new PBEntryField[] {
					PBEntryField.Street,
					PBEntryField.Location,
					PBEntryField.VoiceNumber,
					PBEntryField.FaxNumber,
			};
		}

		try {
			synchronized (outlookLock) { // Serialize access to Outlook to avoid message filter error
				log.fine("Creating Application...");
				app = new _Application("Outlook.Application");
				
				String appVersion = app.getVersion();
				log.info("Outlook version is " + appVersion);
				
				int pos = appVersion.indexOf('.');
				int major = -1;
				if (pos >= 0) {
					major = Integer.parseInt(appVersion.substring(0, pos));
				}
				useOl2007Api = (major >= 12); // Outlook 2007 has 12.0.0.6562; 2003 is 11....
				
				log.fine("Use Outlook 2007+ API: " + useOl2007Api);
				
				log.fine("Loading mappings...");
				loadBusinessAddressMapping(propertyMap_Business);
				loadHomeAddressMapping(propertyMap_Home);
				loadOtherAddressMapping(propertyMap_Other);
				
				log.fine("Got Application, getting MAPI namespace");
				nameSpace = app.getNamespace("MAPI");

				if (Utils.debugMode)
					log.fine("Getting contact folder with folderID=" + settings.folderID + " and storeID=" + settings.storeID);
				MAPIFolder contacts = nameSpace.getFolderFromID(settings.folderID, new Variant(settings.storeID));
				folderName = contacts.getName();

				if (Utils.debugMode)
					log.fine("Got folder name \"" + folderName + "\", reading items now...");
				
				if (useOl2007Api) {
					loadOutlook2007(contacts);
				} else {
					loadPreOutlook2007(contacts);
				}
			}
			log.fine("Successfully loaded phone book");
			open = true;
		} catch (UnsatisfiedLinkError ule) {
			throw new PhoneBookException("Cannot initialize COM connection to Outlook: " + ule.getMessage(), ule, false);
		}
		nameSpace = null;
		app = null;
	}


	private void loadOutlook2007(MAPIFolder contacts) {
		log.info("Loading contacts using Outlook 2007+ API...");
		entries.clear();
		
		// Build the list of needed fields
		Map<String,Integer> columnMap = new HashMap<String,Integer>();
		Integer dummy = Integer.valueOf(-1);
		for (String col : propertyMap_Business.values()) {
			columnMap.put(col, dummy);
		}
		for (String col : propertyMap_Home.values()) {
			columnMap.put(col, dummy);
		}
		for (String col : propertyMap_Other.values()) {
			columnMap.put(col, dummy);
		}
		
		String[] columns = columnMap.keySet().toArray(new String[columnMap.keySet().size()]);
		Table tbl = contacts.getTable(new Variant("[MessageClass] = \"IPM.Contact\""));
		Columns cols = tbl.getColumns();
		cols.removeAll();
		for (int i = 0; i < columns.length; i++) {
			cols.add(columns[i]);
			columnMap.put(columns[i], i);
		}
		
		if (Utils.debugMode) {
			log.fine("Union of used columns is: " + columnMap);
		}
		
		// Build index maps:
		Map<PBEntryField,Integer> indexMap_Business = buildIndexMap(propertyMap_Business, columnMap);
		Map<PBEntryField,Integer> indexMap_Home = buildIndexMap(propertyMap_Home, columnMap);
		Map<PBEntryField,Integer> indexMap_Other = buildIndexMap(propertyMap_Other, columnMap);
		
		int rowCount = tbl.getRowCount();
		if (Utils.debugMode) {
			log.fine("Row count: " + rowCount);
		}
		SafeArray tableArray = tbl.getArray(rowCount).toSafeArray(true);
		if (Utils.debugMode) {
			log.fine("Got an array: rows=["  + tableArray.getLBound(1) + ".." + tableArray.getUBound(1) + "]; cols=["  + tableArray.getLBound(2) + ".." + tableArray.getUBound(2) + "]");
		}
		for (int row=0; row<tableArray.getUBound(1); row++) {
			int numAddresses = 0;
			OlTablePhoneBookEntry pbeOther = new OlTablePhoneBookEntry(this, tableArray, row, indexMap_Other);
			if (pbeOther.hasAddress()) {
				entries.add(pbeOther);
				numAddresses++;
			}
			OlTablePhoneBookEntry pbeHome = new OlTablePhoneBookEntry(this, tableArray, row, indexMap_Home);
			if (pbeHome.hasAddress()) {
				entries.add(pbeHome);
				numAddresses++;
			}
			OlTablePhoneBookEntry pbeBusiness = new OlTablePhoneBookEntry(this, tableArray, row, indexMap_Business);
			if (numAddresses == 0 || pbeBusiness.hasAddress()) {
				entries.add(pbeBusiness);
				numAddresses++;
			}
			if (numAddresses > 1) {
				pbeOther.setSuffix(_("Other"));
				pbeHome.setSuffix(_("Home"));
				pbeBusiness.setSuffix(_("Business"));
			}
		}
		
		//long startTime = System.currentTimeMillis();/
		// Used 2 secs
//		SafeArray table = tbl.getArray(tbl.getRowCount()).toSafeArray(true);
//		for (int i=table.getLBound(1); i<table.getUBound(1); i++) {
//			//System.out.print("[");
//			for (int j=table.getLBound(2); j<table.getUBound(2); j++) {
//				table.getString(new int[] {i,j});
//			}
//			//System.out.println("]");
//		}
		
		// Used 13 secs
//		while (!tbl.getEndOfTable()) {
//			Row row = tbl.getNextRow();
//			SafeArray arr = row.getValues().toSafeArray();
//			for (int i=arr.getLBound(); i<arr.getUBound(); i++) {
//				arr.getString(i);
//			}
//		}
//		System.out.println(System.currentTimeMillis()-startTime);
		
		_Items cl = contacts.getItems();
		cl = cl.restrict("[MessageClass] = \"IPM.DistList\"");
		loadOutlookItems(cl);
	}
	
	private Map<PBEntryField,Integer> buildIndexMap(Map<PBEntryField,String> propertyMap, Map<String,Integer> columnMap) {
		// Coalesce the two maps for performance
		Map<PBEntryField,Integer> res  = new EnumMap<PBEntryField,Integer>(PBEntryField.class);
		for (PBEntryField field : propertyMap.keySet()) {
			Integer i = columnMap.get(propertyMap.get(field));
			if (i != null) {
				res.put(field, i);
			}
		}
		return res;
	}
	
	private void loadPreOutlook2007(MAPIFolder contacts) {
		log.info("Loading contacts using pre-Outlook 2007 API...");
		entries.clear();
		
		_Items cl = contacts.getItems();
		loadOutlookItems(cl);
	}


	private void loadOutlookItems(_Items cl) {
		final int dispIDClass = cl.getIDOfName("Class");
		final int itemCount = cl.getCount();
		for (int i=1; i<=itemCount; i++) {
			Dispatch item = cl.item(i).toDispatch();
			int itemClass = Dispatch.get(item, dispIDClass).getInt();
			if (Utils.debugMode) {
				log.fine("Item #" + i + ": itemClass=" + itemClass);
			}
			switch (itemClass) {
			case OlObjectClass.olContact:
				if (Utils.debugMode) {
					log.fine("Item #" + i + ": is a ContactItem");
				}

				_ContactItem ci = new _ContactItem(item);
				int numAddresses = 0;
				
				OutlookPhoneBookEntry pbeOther = new OutlookPhoneBookEntry(this, ci, propertyMap_Other);
				if (pbeOther.hasAddress()) {
					pbeOther.loadFully();
					entries.add(pbeOther);
					numAddresses++;
				}
				OutlookPhoneBookEntry pbeHome = new OutlookPhoneBookEntry(this, ci, propertyMap_Home);
				if (pbeHome.hasAddress()) {
					pbeHome.loadFully();
					entries.add(pbeHome);
					numAddresses++;
				}
				OutlookPhoneBookEntry pbeBusiness = new OutlookPhoneBookEntry(this, ci, propertyMap_Business);
				if (numAddresses == 0 || pbeBusiness.hasAddress()) {
					pbeBusiness.loadFully();
					entries.add(pbeBusiness);
					numAddresses++;
				}
				if (numAddresses > 1) {
					pbeOther.setSuffix(_("Other"));
					pbeHome.setSuffix(_("Home"));
					pbeBusiness.setSuffix(_("Business"));
				}
				break;
			case OlObjectClass.olDistributionList:
				if (Utils.debugMode) {
					log.fine("Item #" + i + ": is a DistListItem");
				}
				if (settings.accessDistributionLists) {
					_DistListItem dl = new _DistListItem(item);
					entries.add(new OutlookDistList(this, dl));
				} else {
					log.info(folderName + ": Ignoring item #" + i + " because it is a DistListItem");
				}
				break;
			default:
				log.info(folderName + " item #" + i + ": has a unhandled class: " + itemClass);
			}
		}
	}

	
	@Override
	public void close() {
		open = false;
		entries.clear();
		app = null;
		nameSpace = null;
	}

	@Override
	public boolean isReadOnly() {
		return true;
	}
	
	
}
