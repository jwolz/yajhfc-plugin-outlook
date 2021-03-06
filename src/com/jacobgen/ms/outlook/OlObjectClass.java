/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.jacobgen.ms.outlook;

import com.jacob.com.*;

public interface OlObjectClass {

	public static final int olApplication = 0;
	public static final int olNamespace = 1;
	public static final int olFolder = 2;
	public static final int olRecipient = 4;
	public static final int olAttachment = 5;
	public static final int olAddressList = 7;
	public static final int olAddressEntry = 8;
	public static final int olFolders = 15;
	public static final int olItems = 16;
	public static final int olRecipients = 17;
	public static final int olAttachments = 18;
	public static final int olAddressLists = 20;
	public static final int olAddressEntries = 21;
	public static final int olAppointment = 26;
	public static final int olMeetingRequest = 53;
	public static final int olMeetingCancellation = 54;
	public static final int olMeetingResponseNegative = 55;
	public static final int olMeetingResponsePositive = 56;
	public static final int olMeetingResponseTentative = 57;
	public static final int olRecurrencePattern = 28;
	public static final int olExceptions = 29;
	public static final int olException = 30;
	public static final int olAction = 32;
	public static final int olActions = 33;
	public static final int olExplorer = 34;
	public static final int olInspector = 35;
	public static final int olPages = 36;
	public static final int olFormDescription = 37;
	public static final int olUserProperties = 38;
	public static final int olUserProperty = 39;
	public static final int olContact = 40;
	public static final int olDocument = 41;
	public static final int olJournal = 42;
	public static final int olMail = 43;
	public static final int olNote = 44;
	public static final int olPost = 45;
	public static final int olReport = 46;
	public static final int olRemote = 47;
	public static final int olTask = 48;
	public static final int olTaskRequest = 49;
	public static final int olTaskRequestUpdate = 50;
	public static final int olTaskRequestAccept = 51;
	public static final int olTaskRequestDecline = 52;
	public static final int olExplorers = 60;
	public static final int olInspectors = 61;
	public static final int olPanes = 62;
	public static final int olOutlookBarPane = 63;
	public static final int olOutlookBarStorage = 64;
	public static final int olOutlookBarGroups = 65;
	public static final int olOutlookBarGroup = 66;
	public static final int olOutlookBarShortcuts = 67;
	public static final int olOutlookBarShortcut = 68;
	public static final int olDistributionList = 69;
	public static final int olPropertyPageSite = 70;
	public static final int olPropertyPages = 71;
	public static final int olSyncObject = 72;
	public static final int olSyncObjects = 73;
	public static final int olSelection = 74;
	public static final int olLink = 75;
	public static final int olLinks = 76;
	public static final int olSearch = 77;
	public static final int olResults = 78;
	public static final int olViews = 79;
	public static final int olView = 80;
	public static final int olItemProperties = 98;
	public static final int olItemProperty = 99;
	public static final int olReminders = 100;
	public static final int olReminder = 101;
	public static final int olConflict = 102;
	public static final int olConflicts = 103;
	public static final int olSharing = 104;
	public static final int olAccount = 105;
	public static final int olAccounts = 106;
	public static final int olStore = 107;
	public static final int olStores = 108;
	public static final int olSelectNamesDialog = 109;
	public static final int olExchangeUser = 110;
	public static final int olExchangeDistributionList = 111;
	public static final int olPropertyAccessor = 112;
	public static final int olStorageItem = 113;
	public static final int olRules = 114;
	public static final int olRule = 115;
	public static final int olRuleActions = 116;
	public static final int olRuleAction = 117;
	public static final int olMoveOrCopyRuleAction = 118;
	public static final int olSendRuleAction = 119;
	public static final int olTable = 120;
	public static final int olRow = 121;
	public static final int olAssignToCategoryRuleAction = 122;
	public static final int olPlaySoundRuleAction = 123;
	public static final int olMarkAsTaskRuleAction = 124;
	public static final int olNewItemAlertRuleAction = 125;
	public static final int olRuleConditions = 126;
	public static final int olRuleCondition = 127;
	public static final int olImportanceRuleCondition = 128;
	public static final int olFormRegion = 129;
	public static final int olCategoryRuleCondition = 130;
	public static final int olFormNameRuleCondition = 131;
	public static final int olFromRuleCondition = 132;
	public static final int olSenderInAddressListRuleCondition = 133;
	public static final int olTextRuleCondition = 134;
	public static final int olAccountRuleCondition = 135;
	public static final int olClassTableView = 136;
	public static final int olClassIconView = 137;
	public static final int olClassCardView = 138;
	public static final int olClassCalendarView = 139;
	public static final int olClassTimeLineView = 140;
	public static final int olViewFields = 141;
	public static final int olViewField = 142;
	public static final int olOrderField = 144;
	public static final int olOrderFields = 145;
	public static final int olViewFont = 146;
	public static final int olAutoFormatRule = 147;
	public static final int olAutoFormatRules = 148;
	public static final int olColumnFormat = 149;
	public static final int olColumns = 150;
	public static final int olCalendarSharing = 151;
	public static final int olCategory = 152;
	public static final int olCategories = 153;
	public static final int olColumn = 154;
	public static final int olClassNavigationPane = 155;
	public static final int olNavigationModules = 156;
	public static final int olNavigationModule = 157;
	public static final int olMailModule = 158;
	public static final int olCalendarModule = 159;
	public static final int olContactsModule = 160;
	public static final int olTasksModule = 161;
	public static final int olJournalModule = 162;
	public static final int olNotesModule = 163;
	public static final int olNavigationGroups = 164;
	public static final int olNavigationGroup = 165;
	public static final int olNavigationFolders = 166;
	public static final int olNavigationFolder = 167;
	public static final int olClassBusinessCardView = 168;
	public static final int olAttachmentSelection = 169;
	public static final int olAddressRuleCondition = 170;
	public static final int olUserDefinedProperty = 171;
	public static final int olUserDefinedProperties = 172;
	public static final int olFromRssFeedRuleCondition = 173;
	public static final int olClassTimeZone = 174;
	public static final int olClassTimeZones = 175;
}
