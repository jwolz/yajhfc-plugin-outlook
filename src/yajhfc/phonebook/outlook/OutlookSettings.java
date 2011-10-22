package yajhfc.phonebook.outlook;

import yajhfc.phonebook.GeneralConnectionSettings;
import yajhfc.phonebook.PBEntryField;

public class OutlookSettings extends GeneralConnectionSettings {
	public String folderID;
	public boolean readOnly;
	
	public void loadBusinessAddressMapping() {
		setMappingFor(PBEntryField.Comment, "Body");
		setMappingFor(PBEntryField.Company, "CompanyName");
		setMappingFor(PBEntryField.Country, "BusinessAddressCountry");
		setMappingFor(PBEntryField.Department, "Department");
		setMappingFor(PBEntryField.EMailAddress, "Email1Address");
		setMappingFor(PBEntryField.FaxNumber, "BusinessFaxNumber");
		setMappingFor(PBEntryField.GivenName, "FirstName");
		setMappingFor(PBEntryField.Location, "BusinessAddressCity");
		setMappingFor(PBEntryField.Name, "LastName");
		setMappingFor(PBEntryField.Position, "JobTitle"); 
		setMappingFor(PBEntryField.State, "BusinessAddressState");
		setMappingFor(PBEntryField.Street, "BusinessAddressStreet");
		setMappingFor(PBEntryField.Title, "Title");
		setMappingFor(PBEntryField.VoiceNumber, "BusinessTelephoneNumber");
		setMappingFor(PBEntryField.WebSite, "BusinessHomePage");
		setMappingFor(PBEntryField.ZIPCode, "BusinessAddressPostalCode");
	}
	
	public void loadHomeAddressMapping() {
		setMappingFor(PBEntryField.Comment, "Body");
		setMappingFor(PBEntryField.Company, "CompanyName");
		setMappingFor(PBEntryField.Country, "HomeAddressCountry");
		setMappingFor(PBEntryField.Department, "Department");
		setMappingFor(PBEntryField.EMailAddress, "Email1Address");
		setMappingFor(PBEntryField.FaxNumber, "HomeFaxNumber");
		setMappingFor(PBEntryField.GivenName, "FirstName");
		setMappingFor(PBEntryField.Location, "HomeAddressCity");
		setMappingFor(PBEntryField.Name, "LastName");
		setMappingFor(PBEntryField.Position, "JobTitle"); 
		setMappingFor(PBEntryField.State, "HomeAddressState");
		setMappingFor(PBEntryField.Street, "HomeAddressStreet");
		setMappingFor(PBEntryField.Title, "Title");
		setMappingFor(PBEntryField.VoiceNumber, "HomeTelephoneNumber");
		setMappingFor(PBEntryField.WebSite, "HomeHomePage");
		setMappingFor(PBEntryField.ZIPCode, "HomeAddressPostalCode");
	}
	
	public OutlookSettings() {
		loadBusinessAddressMapping();
	}
}
