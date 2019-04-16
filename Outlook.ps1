#constants for Outlook based on https://docs.microsoft.com/en-us/office/vba/api/Outlook
$ol = [Ordered]@{

    #values from outlook.olaccounttype
    #****************************
    olEas                                  =	4		# An account that uses Exchange ActiveSync (EAS) on mobile devices.
    olExchange                             =	0		# An Exchange account.
    olHttp                                 =	3		# An HTTP account.
    olImap                                 =	1		# An IMAP account.
    olOtherAccount                         =	5		# Other or unknown account.
    olPop3                                 =	2		# A POP3 account.

    #values from outlook.olactioncopylike
    #****************************
    olForward                              =	2		# Properties of the new item will be set such that the new item is a forward of the original item. If creating a new  MailItem, the value of the To and CC properties in the new item will be empty and the Subject property of the new item will be the original Subject with a prefix such as &quot;FW:&quot; (see Prefix property) added. The attachments on the original item will be copied to the new item.
    olReply                                =	0		# Properties of the new item will be set such that the new item is a reply to the original item. If creating a new  MailItem, the value of the original To field will be copied to the SenderEmailAddress property of the new item, the CC field will be blank and the Subject field of the new item will be the original Subject with a prefix such as &quot;RE:&quot; (see Prefix property) added.
    olReplyAll                             =	1		# Properties of the new item will be set such that the new item is a reply to all of the senders of the original item. If creating a new  MailItem, the value of the SenderEmailAddress and CC properties will be copied to the To property of the new item and the Subject property of the new item will be the Subject of the original item with a prefix such as &quot;RE:&quot; (see Prefix property) added.
    olReplyFolder                          =	3		# If creating a new  PostItem based on an old one, the Post To property of the new item will contain the active folder address, the Subject property of the original item will be copied to the ConversationTopic property of the new item, and the Subject property of the new item will be empty.
    olRespond                              =	4		# Used exclusively for voting button actions.

    #values from outlook.olactionreplystyle
    #****************************
    olEmbedOriginalItem                    =	1		# The reply will include the original item embedded in it.
    olIncludeOriginalText                  =	2		# The reply will include the text of the original item.
    olIndentOriginalText                   =	3		# The reply will include the indented text of the original item.
    olLinkOriginalItem                     =	4		# The reply will include a link to the original item.
    olOmitOriginalText                     =	0		# The reply will not include any references to the original item or its text.
    olReplyTickOriginalText                =	1000		# The reply will include the original text with each line preceded by a symbol such as &quot;&gt;&quot;.
    olUserPreference                       =	5		# The reply style will be set based on the user's preference.

    #values from outlook.olactionresponsestyle
    #****************************
    olOpen                                 =	0		# Indicates that a form will be opened.
    olPrompt                               =	2		# Indicates that the user will be prompted to open or send the form.
    olSend                                 =	1		# Indicates that the form will be sent immediately.

    #values from outlook.olactionshowon
    #****************************
    olDontShow                             =	0		# Indicates that the action will not be displayed on the menu or toolbar.
    olMenu                                 =	1		# Indicates that the action will be displayed as an available action on the menu.
    olMenuAndToolbar                       =	2		# Indicates that the action will be displayed as an available action on the menu and the toolbar.

    #values from outlook.oladdressentryusertype
    #****************************
    olExchangeAgentAddressEntry            =	3		# An address entry that is an Exchange agent.
    olExchangeDistributionListAddressEntry	=	1		# An address entry that is an Exchange distribution list.
    olExchangeOrganizationAddressEntry     =	4		# An address entry that is an Exchange organization.
    olExchangePublicFolderAddressEntry     =	2		# An address entry that is an Exchange public folder.
    olExchangeRemoteUserAddressEntry       =	5		# An Exchange user that belongs to a different Exchange forest.
    olExchangeUserAddressEntry             =	0		# An Exchange user that belongs to the same Exchange forest.
    olLdapAddressEntry                     =	20		# An address entry that uses the Lightweight Directory Access Protocol (LDAP).
    olOtherAddressEntry                    =	40		# A custom or some other type of address entry such as FAX.
    olOutlookContactAddressEntry           =	10		# An address entry in an Outlook Contacts folder.
    olOutlookDistributionListAddressEntry  =	11		# An address entry that is an Outlook distribution list.
    olSmtpAddressEntry                     =	30		# An address entry that uses the Simple Mail Transfer Protocol (SMTP).

    #values from outlook.oladdresslisttype
    #****************************
    olCustomAddressList                    =	4		# A custom address book provider.
    olExchangeContainer                    =	1		# A container for address lists on an Exchange server.
    olExchangeGlobalAddressList            =	0		# An Exchange Global Address List.
    olOutlookAddressList                   =	2		# An address list that corresponds to the Outlook Contacts Address Book.
    olOutlookLdapAddressList               =	3		# An address list that uses the Lightweight Directory Access Protocol (LDAP).

    #values from outlook.olalign
    #****************************
    olAlignCenter                          =	1		# Indicates that the label for the specified column should be centered.
    olAlignLeft                            =	0		# Indicates that the label for the specified column should be left-aligned.
    olAlignRight                           =	2		# Indicates that the label for the specified column should be right-aligned.

    #values from outlook.olalignment
    #****************************
    olAlignmentLeft                        =	0		# Places the caption to the left of the control.
    olAlignmentRight                       =	1		# Places the caption to the right of the control.

    #values from outlook.olalwaysdeleteconversation
    #****************************
    olAlwaysDelete                         =	1		# New items of the conversation are always moved to the Deleted Items folder for the store that contains the items
    olAlwaysDeleteUnsupported              =	2		# The specified store does not support the action of always moving items to the Deleted Items folder of that store.
    olDoNotDelete                          =	0		# New items joining the conversation are not moved to the Deleted Items folder on the specified delivery store, and existing conversation items in the Deleted Items folder are moved to the Inbox.

    #values from outlook.olappointmentcopyoptions
    #****************************
    olCopyAsAccept                         =	2		# Creates an appointment in the destination folder and accepts the meeting request automatically.
    olCreateAppointment                    =	1		# Creates an appointment in the destination folder without defaulting to a response or prompting for a response.
    olPromptUser                           =	0		# Copies the appointment to the destination folder and prompts the user to accept the request before completing the copy operation.

    #values from outlook.olappointmenttimefield
    #****************************
    olAppointmentTimeFieldEnd              =	3		# The control is bound to the end time of the appointment.
    olAppointmentTimeFieldNone             =	1		# The control is not bound.
    olAppointmentTimeFieldStart            =	2		# The control is bound to the start time of the appointment.

    #values from outlook.olattachmentblocklevel
    #****************************
    olAttachmentBlockLevelNone             =	0		# There is no restriction on the type of the attachment based on its file extension.
    olAttachmentBlockLevelOpen             =	1		# There is a restriction on the type of the attachment based on its file extension such that users must first save the attachment to disk before opening it.

    #values from outlook.olattachmenttype
    #****************************
    olByReference                          =	4		# This value is no longer supported since Microsoft Outlook 2007. Use  olByValue to attach a copy of a file in the file system.
    olByValue                              =	1		# The attachment is a copy of the original file and can be accessed even if the original file is removed.
    olEmbeddeditem                         =	5		# The attachment is an Outlook message format file (.msg) and is a copy of the original message.
    olOLE                                  =	6		# The attachment is an OLE document.

    #values from outlook.olautodiscoverconnectionmode
    #****************************
    olAutoDiscoverConnectionExternal       =	1		# Connection is over the Internet.
    olAutoDiscoverConnectionInternal       =	2		# Connection is over the Intranet.
    olAutoDiscoverConnectionInternalDomain	=	3		# Connection is in the same domain over the Intranet.
    olAutoDiscoverConnectionUnknown        =	0		# Other or unknown connection, or no connection.

    #values from outlook.olautopreview
    #****************************
    olAutoPreviewAll                       =	0		# All items are automatically previewed.
    olAutoPreviewNone                      =	2		# No items are automatically previewed.
    olAutoPreviewUnread                    =	1		# Only unread items are automatically previewed.

    #values from outlook.olbackstyle
    #****************************
    olBackStyleOpaque                      =	1		# Indicates the background color of the control is rendered.
    olBackStyleTransparent                 =	0		# Indicates that the background color of the control is the background color of the parent control. It is not a truly transparent display.

    #values from outlook.olbodyformat
    #****************************
    olFormatHTML                           =	2		# HTML format
    olFormatPlain                          =	1		# Plain format
    olFormatRichText                       =	3		# Rich text format
    olFormatUnspecified                    =	0		# Unspecified format

    #values from outlook.olborderstyle
    #****************************
    olBorderStyleNone                      =	0		# The control is rendered without any border.
    olBorderStyleSingle                    =	1		# The control is rendered with a single-line border.

    #values from outlook.olbusinesscardtype
    #****************************
    olBusinessCardTypeInterConnect         =	1		# Indicates that the  ContactItem uses the Microsoft Office InterConnect format for the associated Electronic Business Card.
    olBusinessCardTypeOutlook              =	0		# Indicates that the  ContactItem object uses the Microsoft Outlook format for the associated Electronic Business Card.

    #values from outlook.olbusystatus
    #****************************
    olBusy                                 =	2		# The user is busy.
    olFree                                 =	0		# The user is available.
    olOutOfOffice                          =	3		# The user is out of office.
    olTentative                            =	1		# The user has a tentative appointment scheduled.
    olWorkingElsewhere                     =	4		# The user is working in a location away from the office.

    #values from outlook.olcalendardetail
    #****************************
    olFreeBusyAndSubject                   =	1		# Free/busy information and the appointment subjects are exported to the iCalendar file.
    olFreeBusyOnly                         =	0		# Only free/busy information is exported to the iCalendar file.
    olFullDetails                          =	2		# Full details of each appointment item are exported to the iCalendar file.

    #values from outlook.olcalendarmailformat
    #****************************
    olCalendarMailFormatDailySchedule      =	0		# The calendar information is formatted as a daily schedule of appointments, containing an hour-by-hour breakdown of the calendar, showing both free and busy time blocks along with working-hour information. This layout is intended to help show recipients which times you are available.
    olCalendarMailFormatEventList          =	1		# The calendar information is formatted as a list of events, containing a list of the calendar appointments without showing any time blocks. This layout is intended to help show recipients the events scheduled for a given time period.

    #values from outlook.olcalendarviewmode
    #****************************
    olCalendarView5DayWeek                 =	4		# Displays a 5-day week.
    olCalendarViewDay                      =	0		# Displays a single day.
    olCalendarViewMonth                    =	2		# Displays a month.
    olCalendarViewMultiDay                 =	3		# Displays a number of days equal to the  DaysInMultiDayMode property value of the CalendarView object.
    olCalendarViewWeek                     =	1		# Displays a 7-day week.

    #values from outlook.olcategorycolor
    #****************************
    olCategoryColorBlack                   =	15		# Black
    olCategoryColorBlue                    =	8		# Blue
    olCategoryColorDarkBlue                =	23		# Dark Blue
    olCategoryColorDarkGray                =	14		# Dark Gray
    olCategoryColorDarkGreen               =	20		# Dark Green
    olCategoryColorDarkMaroon              =	25		# Dark Maroon
    olCategoryColorDarkOlive               =	22		# Dark Olive
    olCategoryColorDarkOrange              =	17		# Dark Orange
    olCategoryColorDarkPeach               =	18		# Dark Peach
    olCategoryColorDarkPurple              =	24		# Dark Purple
    olCategoryColorDarkRed                 =	16		# Dark Red
    olCategoryColorDarkSteel               =	12		# Dark Steel
    olCategoryColorDarkTeal                =	21		# Dark Teal
    olCategoryColorDarkYellow              =	19		# Dark Yellow
    olCategoryColorGray                    =	13		# Gray
    olCategoryColorGreen                   =	5		# Green
    olCategoryColorMaroon                  =	10		# Maroon
    olCategoryColorNone                    =	0		# No color assigned.
    olCategoryColorOlive                   =	7		# Olive
    olCategoryColorOrange                  =	2		# Orange
    olCategoryColorPeach                   =	3		# Peach
    olCategoryColorPurple                  =	9		# Purple
    olCategoryColorRed                     =	1		# Red
    olCategoryColorSteel                   =	11		# Steel
    olCategoryColorTeal                    =	6		# Teal
    olCategoryColorYellow                  =	4		# Yellow

    #values from outlook.olcategoryshortcutkey
    #****************************
    olCategoryShortcutKeyCtrlF10           =	10		# CTRL+F10
    olCategoryShortcutKeyCtrlF11           =	11		# CTRL+F11
    olCategoryShortcutKeyCtrlF12           =	12		# CTRL+F12
    olCategoryShortcutKeyCtrlF2            =	2		# CTRL+F2
    olCategoryShortcutKeyCtrlF3            =	3		# CTRL+F3
    olCategoryShortcutKeyCtrlF4            =	4		# CTRL+F4
    olCategoryShortcutKeyCtrlF5            =	5		# CTRL+F5
    olCategoryShortcutKeyCtrlF6            =	6		# CTRL+F6
    olCategoryShortcutKeyCtrlF7            =	7		# CTRL+F7
    olCategoryShortcutKeyCtrlF8            =	8		# CTRL+F8
    olCategoryShortcutKeyCtrlF9            =	9		# CTRL+F9
    olCategoryShortcutKeyNone              =	0		# No shortcut key specified.

    #values from outlook.olcolor
    #****************************
    olAutoColor                            =	0		# Color is based on system preferences
    olColorAqua                            =	15		# Aqua
    olColorBlack                           =	1		# Black
    olColorBlue                            =	13		# Blue
    olColorFuchsia                         =	14		# Fuchsia
    olColorGray                            =	8		# Gray
    olColorGreen                           =	3		# Green
    olColorLime                            =	11		# Lime
    olColorMaroon                          =	2		# Maroon
    olColorNavy                            =	5		# Navy
    olColorOlive                           =	4		# Olive
    olColorPurple                          =	6		# Purple
    olColorRed                             =	10		# Red
    olColorSilver                          =	9		# Silver
    olColorTeal                            =	7		# Teal
    olColorWhite                           =	16		# White
    olColorYellow                          =	12		# Yellow

    #values from outlook.olcomboboxstyle
    #****************************
    olComboBoxStyleCombo                   =	0		# Indicates that the combo box behaves like a traditional combo box in which the user can type a value in the edit box or select a value from the drop-down list.
    olComboBoxStyleList                    =	1		# Indicates that the combo box behaves like a drop-down list from which the user can only select a value.

    #values from outlook.olcontactphonenumber
    #****************************
    olContactPhoneAssistant                =	0		# Telephone number of the person who is the assistant for the contact
    olContactPhoneBusiness                 =	1		# Business telephone number
    olContactPhoneBusiness2                =	2		# Second business telephone number
    olContactPhoneBusinessFax              =	3		# Business fax number
    olContactPhoneCallback                 =	4		# Callback telephone number
    olContactPhoneCar                      =	5		# Car telephone number
    olContactPhoneCompany                  =	6		# Main company telephone number
    olContactPhoneHome                     =	7		# Home telephone number
    olContactPhoneHome2                    =	8		# Second home telephone number
    olContactPhoneHomeFax                  =	9		# Home fax number
    olContactPhoneISDN                     =	10		# Integrated Services Digital Network (ISDN) phone number
    olContactPhoneMobile                   =	11		# Mobile telephone number
    olContactPhoneOther                    =	12		# Other telephone number
    olContactPhoneOtherFax                 =	13		# Other fax number
    olContactPhonePager                    =	14		# Pager telephone number
    olContactPhonePrimary                  =	15		# Primary telephone number
    olContactPhoneRadio                    =	16		# Radio telephone number
    olContactPhoneTelex                    =	17		# Telex telephone number
    olContactPhoneTTYTTD                   =	18		# TTD/TTY (Teletypewriting Device for the Deaf/Teletypewriter) telephone number

    #values from outlook.oldaysofweek
    #****************************
    olFriday                               =	32		# Friday
    olMonday                               =	2		# Monday
    olSaturday                             =	64		# Saturday
    olSunday                               =	1		# Sunday
    olThursday                             =	16		# Thursday
    olTuesday                              =	4		# Tuesday
    olWednesday                            =	8		# Wednesday

    #values from outlook.oldayweektimescale
    #****************************
    olTimeScale10Minutes                   =	2		# Indicates that each time period represents 10 minutes.
    olTimeScale15Minutes                   =	3		# Indicates that each time period represents 15 minutes.
    olTimeScale30Minutes                   =	4		# Indicates that each time period represents 30 minutes.
    olTimeScale5Minutes                    =	0		# Indicates that each time period represents 5 minutes.
    olTimeScale60Minutes                   =	5		# Indicates that each time period represents 60 minutes.
    olTimeScale6Minutes                    =	1		# Indicates that each time period represents 6 minutes.

    #values from outlook.oldefaultexpandcollapsesetting
    #****************************
    olAllCollapsed                         =	1		# All groups are collapsed when the view is initially displayed.
    olAllExpanded                          =	0		# All groups are expanded when the view is initially displayed.
    olLastViewed                           =	2		# Groups are displayed expanded or collapsed depending on the state the groups were in when the view was last closed.

    #values from outlook.oldefaultfolders
    #****************************
    olFolderCalendar                       =	9		# The Calendar folder.
    olFolderConflicts                      =	19		# The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderContacts                       =	10		# The Contacts folder.
    olFolderDeletedItems                   =	3		# The Deleted Items folder.
    olFolderDrafts                         =	16		# The Drafts folder.
    olFolderInbox                          =	6		# The Inbox folder.
    olFolderJournal                        =	11		# The Journal folder.
    olFolderJunk                           =	23		# The Junk E-Mail folder.
    olFolderLocalFailures                  =	21		# The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderManagedEmail                   =	29		# The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
    olFolderNotes                          =	12		# The Notes folder.
    olFolderOutbox                         =	4		# The Outbox folder.
    olFolderSentMail                       =	5		# The Sent Mail folder.
    olFolderServerFailures                 =	22		# The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderSuggestedContacts              =	30		# The Suggested Contacts folder.
    olFolderSyncIssues                     =	20		# The Sync Issues folder. Only available for an Exchange account.
    olFolderTasks                          =	13		# The Tasks folder.
    olFolderToDo                           =	28		# The To Do folder.
    olPublicFoldersAllPublicFolders        =	18		# The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
    olFolderRssFeeds                       =	25		# The RSS Feeds folder.

    #values from outlook.oldefaultselectnamesdisplaymode
    #****************************
    olDefaultDelegates                     =	6		# Displays one edit box for To recipients, uses localized string representing &quot;Add&quot; for the To button, and localized string representing &quot;Add Users&quot; for the caption.  CcLabel and BccLabel are set to an empty string. Sets SelectNamesDialog.AllowMultipleSelection to True and SelectNamesDialog.NumberOfRecipientSelectors to olTo.
    olDefaultMail                          =	1		# Displays three edit boxes for To, Cc, and Bcc recipients, uses localized strings representing &quot;To&quot;, &quot;Cc&quot;, and &quot;Bcc&quot; for To, Cc, and Bcc buttons, and localized string representing &quot;Select Names&quot; for the caption. Sets  AllowMultipleSelection to True and NumberOfRecipientSelectors to olToCcBcc.
    olDefaultMeeting                       =	2		# Displays three edit boxes for Required, Optional, and Resource recipients, uses localized strings representing &quot;Required&quot;, &quot;Optional&quot;, and &quot;Resources&quot; for the To, Cc, and Bcc buttons, and localized string representing &quot;Select Attendees and Resources&quot; for the caption. Sets  AllowMultipleSelection to True and NumberOfRecipientSelectors to olToCcBcc.
    olDefaultMembers                       =	5		# Displays one edit box for To recipients, uses localized string representing &quot;To&quot; for the To button, and localized string representing &quot;Select Members&quot; for caption.  CcLabel and BccLabel are set to an empty string. Sets AllowMultipleSelection to True and NumberOfRecipientSelectors to olTo.
    olDefaultPickRooms                     =	8		# Displays one edit box for Resource recipients, uses localized string representing &quot;Rooms&quot; for To button, and localized string representing &quot;Select Rooms&quot; for caption.  CcLabel and BccLabel are set to an empty string. Sets AllowMultipleSelection to True and NumberOfRecipientSelectors to ****olShowTo.  InitialDisplayList is set to the Global Address List.
    olDefaultSharingRequest                =	4		# Displays one edit box for To recipients, uses localized string representing &quot;To&quot; for To button, and localized string representing &quot;Select Names&quot; for caption.  CcLabel and BccLabel are set to an empty string. Sets AllowMultipleSelection to True and NumberOfRecipientSelectors to olTo.
    olDefaultSingleName                    =	7		# Displays no edit boxes for recipients, uses localized string representing &quot;Select Name&quot; for caption.  ToLabel,  CcLabel, and  Bcclabel are set to an empty string. Sets AllowMultipleSelection to False and NumberOfRecipientSelectors to olNone.
    olDefaultTask                          =	3		# Displays one edit box for To recipients, uses localized string representing &quot;To&quot; for To button, and localized string representing &quot;Select Task Recipient&quot; for caption.  CcLabel and BccLabel are set to an empty string. Sets AllowMultipleSelection to True and NumberOfRecipientSelectors to olTo.

    #values from outlook.oldisplaymode
    #****************************
    olDisplayModeNormal                    =	0		# Displays Normal mode.
    olDisplayModePortraitReadingPane       =	2		# Displays Portrait Reading Pane mode.
    olDisplayModePortraitView              =	1		# Displays Portrait View mode.

    #values from outlook.oldisplaytype
    #****************************
    olAgent                                =	3		# Agent address
    olDistList                             =	1		# Exchange distribution list
    olForum                                =	2		# Forum address
    olOrganization                         =	4		# Organization address
    olPrivateDistList                      =	5		# Outlook private distribution list
    olRemoteUser                           =	6		# Remote user address
    olUser                                 =	0		# User address

    #values from outlook.oldownloadstate
    #****************************
    olFullItem                             =	1		# Full item has been downloaded.
    olHeaderOnly                           =	0		# Only the header has been downloaded.

    #values from outlook.oldragbehavior
    #****************************
    olDragBehaviorDisabled                 =	0		# The control does not support drag-and-drop activities. It will always display the &quot;no&quot; cursor when an item is dragged over the control.
    olDragBehaviorEnabled                  =	1		# The control can support drag-and-drop activities. Use the drag and drop events to control this behavior.

    #values from outlook.oleditortype
    #****************************
    olEditorHTML                           =	2		# HTML editor
    olEditorRTF                            =	3		# Real Text Format (RTF) editor
    olEditorText                           =	1		# Text editor
    olEditorWord                           =	4		# Microsoft Office Word editor

    #values from outlook.olenterfieldbehavior
    #****************************
    olEnterFieldBehaviorRecallSelection    =	1		# The previous selection is displayed.
    olEnterFieldBehaviorSelectAll          =	0		# The contents of the control are selected and any previous selection is ignored.

    #values from outlook.olexchangeconnectionmode
    #****************************
    olCachedConnectedDrizzle               =	600		# The account is using cached Exchange mode such that headers are downloaded first, followed by the bodies and attachments of full items.
    olCachedConnectedFull                  =	700		# The account is using cached Exchange mode on a Local Area Network or a fast connection with the Exchange server. The user can also select this state manually, disabling auto-detect logic and always downloading full items regardless of connection speed.
    olCachedConnectedHeaders               =	500		# The account is using cached Exchange mode on a dial-up or slow connection with the Exchange server, such that only headers are downloaded. Full item bodies and attachments remain on the server. The user can also select this state manually regardless of connection speed.
    olCachedDisconnected                   =	400		# The account is using cached Exchange mode with a disconnected connection to the Exchange server.
    olCachedOffline                        =	200		# The account is using cached Exchange mode and the user has selected  Work Offline from the File menu.
    olDisconnected                         =	300		# The account has a disconnected connection to the Exchange server.
    olNoExchange                           =	0		# The account does not use an Exchange server.
    olOffline                              =	100		# The account is not connected to an Exchange server and is in the classic offline mode. This also occurs when the user selects  Work Offline from the File menu.
    olOnline                               =	800		# The account is connected to an Exchange server and is in the classic online mode.

    #values from outlook.olexchangestoretype
    #****************************
    olExchangeMailbox                      =	1		# Specifies an Exchange delegate mailbox store.
    olExchangePublicFolder                 =	2		# Specifies an Exchange Public Folder store.
    olNotExchange                          =	3		# Specifies that the store is not an Exchange store.
    olPrimaryExchangeMailbox               =	0		# Specifies a primary Exchange mailbox store.
    olAdditionalExchangeMailbox            =	4		# Specifies an additional Exchange mailbox store.

    #values from outlook.olfolderdisplaymode
    #****************************
    olFolderDisplayFolderOnly              =	1		# Only the contents of the selected folder are displayed.
    olFolderDisplayNoNavigation            =	2		# Folder contents are displayed but no navigation pane is shown.
    olFolderDisplayNormal                  =	0		# Folder is displayed with navigation pane on the left and folder contents on the right.

    #values from outlook.olformatcurrency
    #****************************
    olFormatCurrencyDecimal                =	1		# Displays currency values with decimal places.
    olFormatCurrencyNonDecimal             =	2		# Displays currency values without decimal places, rounding the currency value to the nearest integer.

    #values from outlook.olformatdatetime
    #****************************
    olFormatDateTimeBestFit                =	17		# Displays the date/time value using the best fit for the data contained in the column.
    olFormatDateTimeLongDate               =	6		# Displays a long date, without the day name, according to your locale's format.
    olFormatDateTimeLongDateReversed       =	7		# Displays a long date, reversing the day number and month name, according to your locale's format.
    OlFormatDateTimeLongDayDate            =	5		# Displays a long date, with the day name, according to your locale's format.
    olFormatDateTimeLongDayDateTime        =	1		# Displays a long date and long time according to your locale's format.
    olFormatDateTimeLongTime               =	15		# Displays a long time according to your locale's format.
    olFormatDateTimeShortDate              =	8		# Displays a short date according to your locale's format.
    olFormatDateTimeShortDateNumOnly       =	9		# Displays a short date, using numeric representations of the day, month, and year, according to your locale's format.
    olFormatDateTimeShortDateTime          =	2		# Displays a short date and short time according to your locale's format.
    olFormatDateTimeShortDayDate           =	13		# Displays a short date, with the day abbreviation, according to your locale's format.
    olFormatDateTimeShortDayDateTime       =	3		# Displays a day abbreviation, short date, and short time according to your locale's format.
    olFormatDateTimeShortDayMonth          =	10		# Displays the month and the day of a date, with the day abbreviation, according to your locale's format.
    olFormatDateTimeShortDayMonthDateTime  =	4		# Displays a day abbreviation, day number, month number, and short time according to your locale's format.
    olFormatDateTimeShortMonthYear         =	11		# Displays the month and year of a date, according to your locale's format.
    olFormatDateTimeShortMonthYearNumOnly  =	12		# Displays the month and year of a date, using numeric representations of the month and year, according to your locale's format.
    olFormatDateTimeShortTime              =	16		# Displays a short time according to your locale's format.

    #values from outlook.olformatduration
    #****************************
    olFormatDurationLong                   =	2		# Displays duration values using full names for time period descriptions.
    olFormatDurationLongBusiness           =	4		# Displays duration values, taking into consideration specified calendar work week settings and using full names for time period descriptions.
    olFormatDurationShort                  =	1		# Displays duration values using abbreviations for time period descriptions.
    olFormatDurationShortBusiness          =	3		# Displays duration values, taking into consideration specified calendar work week settings and using abbreviations for time period descriptions.

    #values from outlook.olformatenumeration
    #****************************
    olFormatEnumBitmap                     =	1		# Display as a bitmap with a pop-up list.
    olFormatEnumText                       =	2		# Display as text with a pop-up list.

    #values from outlook.olformatinteger
    #****************************
    olFormatIntegerComputer1               =	2		# Displays integer values, representing bytes, as kilobytes (with the abbreviation &quot;K&quot;) depending on the value. For example, the integer value of 1048576 is displayed as &quot;1,024 K&quot;.
    olFormatIntegerComputer2               =	3		# Displays integer values, representing bytes, as either kilobytes (with the abbreviation &quot;K&quot;), megabytes (with the abbreviation &quot;M&quot;), or gigabytes (with the abbreviation &quot;G&quot;), depending on the value. For example, the integer value of 2048 is displayed as &quot;2 K&quot;.
    olFormatIntegerComputer3               =	4		# Displays integer values, representing bytes, as either bytes (with the abbreviation &quot;B&quot;), kilobytes (with the abbreviation &quot;KB&quot;), megabytes (with the abbreviation &quot;MB&quot;), or gigabytes (with the abbreviation &quot;GB&quot;), depending on the value. For example, the integer value of 1000 is displayed as &quot;1,000 B&quot;.
    olFormatIntegerPlain                   =	1		# Displays integer values using the number format specified in your computer's regional settings.

    #values from outlook.olformatkeywords
    #****************************
    olFormatKeywordsText                   =	1		# Displays values as text.

    #values from outlook.olformatnumber
    #****************************
    olFormatNumber1Decimal                 =	3		# Displays formatted number values, including one fixed decimal place, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.2&quot;.
    olFormatNumber2Decimal                 =	4		# Displays formatted number values, including two fixed decimal places, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.16&quot;.
    olFormatNumberAllDigits                =	1		# Displays formatted number values, including any decimal places specified in the value, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.155&quot;.
    olFormatNumberComputer1                =	6		# Displays formatted number values, representing bytes, as kilobytes (with the abbreviation &quot;K&quot;) depending on the value. For example, the integer value of 1048576 is displayed as &quot;1,024 K&quot;.
    olFormatNumberComputer2                =	7		# Displays formatted number values, representing bytes, as either kilobytes (with the abbreviation &quot;K&quot;), megabytes (with the abbreviation &quot;M&quot;), or gigabytes (with the abbreviation &quot;G&quot;), depending on the value. For example, the integer value of 2048 is displayed as &quot;2 K&quot;.
    olFormatNumberComputer3                =	8		# Displays formatted number values, representing bytes, as either bytes (with the abbreviation &quot;B&quot;), kilobytes (with the abbreviation &quot;KB&quot;), megabytes (with the abbreviation &quot;MB&quot;), or gigabytes (with the abbreviation &quot;GB&quot;), depending on the value. For example, the integer value of 1000 is displayed as &quot;1,000 B&quot;.
    olFormatNumberRaw                      =	9		# Displays unformatted number values, including any decimal places specified in the value. For example, the value 1048576 is displayed as &quot;1048576&quot;.
    olFormatNumberScientific               =	5		# Displays formatted number values, using scientific notation. For example, the value 1048576 is displayed as &quot;1.049E+06&quot;.
    olFormatNumberTruncated                =	2		# Displays formatted number values as integers, rounding all decimal values, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.1 is displayed as &quot;4,010&quot;.

    #values from outlook.olformatpercent
    #****************************
    olFormatPercent1Decimal                =	2		# Displays formatted number values, including one fixed decimal place, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.2%&quot;.
    olFormatPercent2Decimal                =	3		# Displays formatted number values, including two fixed decimal places, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.16%&quot;
    olFormatPercentAllDigits               =	4		# Displays formatted number values, including any decimal places specified in the value, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010.155%&quot;
    olFormatPercentRounded                 =	1		# Displays formatted number values as integers, rounding all decimal values, using the group and decimal delimiters specified in the system's regional settings. For example, the value 4010.155 is displayed as &quot;4,010%&quot;.

    #values from outlook.olformatsmartfrom
    #****************************
    olFormatSmartFromFromOnly              =	2		# Display the value of the  From Outlook item property. If no value is available, display an empty string.
    olFormatSmartFromFromTo                =	1		# Display the value of the  From Outlook item property. If no value is available, display instead the value of the To Outlook item property.

    #values from outlook.olformattext
    #****************************
    olFormatTextText                       =	1		# Display values as text.

    #values from outlook.olformatyesno
    #****************************
    olFormatYesNoIcon                      =	4		# Displays a check box icon.
    olFormatYesNoOnOff                     =	2		# Displays &quot;On&quot; or &quot;Off&quot;.
    olFormatYesNoTrueFalse                 =	3		# Displays &quot;True&quot; or &quot;False&quot;.
    olFormatYesNoYesNo                     =	1		# Displays &quot;Yes&quot; or &quot;No&quot;.

    #values from outlook.olformregionicon
    #****************************
    olFormRegionIconDefault                =	1		# The default icon for an item.
    olFormRegionIconEncrypted              =	9		# The icon to display when an item has been encrypted.
    olFormRegionIconForwarded              =	5		# The icon to display when an item has been forwarded.
    olFormRegionIconPage                   =	11		# The icon to display in the ribbon for the replacement form region.
    olFormRegionIconRead                   =	3		# The icon to display when an item has been read.
    olFormRegionIconRecurring              =	12		# The icon to display when an item is a recurring appointment, meeting, or task.
    olFormRegionIconReplied                =	4		# The icon to display when an item has been replied to.
    olFormRegionIconSigned                 =	8		# The icon to display when an item has been signed with a digital signature.
    olFormRegionIconSubmitted              =	7		# The icon to display when an item has been sent.
    olFormRegionIconUnread                 =	2		# The icon to display when an item has not been read.
    olFormRegionIconUnsent                 =	6		# The icon to display when an item has not been sent.
    olFormRegionIconWindow                 =	10		# The icon to display on the Inspector when an item has been opened.

    #values from outlook.olformregionmode
    #****************************
    olFormRegionCompose                    =	1		# The form region is in a compose page of a message or any unsendable item such as a contact item.
    olFormRegionPreview                    =	2		# The form region is in the Reading Pane.
    olFormRegionRead                       =	0		# The form region is in the read page of a received message or a post note.

    #values from outlook.olformregionsize
    #****************************
    olFormRegionTypeAdjoining              =	1		# The form region is an adjoining form region.
    olFormRegionTypeSeparate               =	0		# The form region is a separate form region. This includes a separate form region that is added to a standard form as a custom page, a separate form region that replaces the default page of a standard form, or a separate form region that replaces all pages in a standard form.

    #values from outlook.olformregistry
    #****************************
    olDefaultRegistry                      =	0		# The Form is registered in the user's default form registry.
    olFolderRegistry                       =	3		# The Form is registered in a form registry specific to a particular folder, and can only be accessed from that folder.
    olOrganizationRegistry                 =	4		# The Form is registered in the organizational form registry. The form is available to all users.
    olPersonalRegistry                     =	2		# The Form is registered in the user's personal registry and is only accessible to that user.

    #values from outlook.olgender
    #****************************
    olFemale                               =	1		# Female
    olMale                                 =	2		# Male
    olUnspecified                          =	0		# Unspecified gender

    #values from outlook.olgridlinestyle
    #****************************
    olGridLineDashes                       =	3		# Dashed lines are used to draw the grid.
    olGridLineLargeDots                    =	2		# Lines using large dots are used to draw the grid.
    olGridLineNone                         =	0		# No lines are displayed.
    olGridLineSmallDots                    =	1		# Lines using small dots are used to draw the grid.
    olGridLineSolid                        =	4		# Solid lines are used to draw the grid.

    #values from outlook.olgrouptype
    #****************************
    olCustomFoldersGroup                   =	0		# Identifies a user-defined navigation group, added using either the Outlook user interface or an add-in.
    olFavoriteFoldersGroup                 =	4		# Identifies the  Favorite Folders navigation group. This navigation group exists only within the NavigationGroups collection of a MailModule object and cannot be created in or accessed from other modules.
    olMyFoldersGroup                       =	1		# Identifies a navigation group that, by default, contains any folders that are part of the local store.
    olOtherFoldersGroup                    =	3		# Identifies a navigation group that, by default, contains shared folders from sources other than that of other persons.
    olPeopleFoldersGroup                   =	2		# Identifies a navigation group that, by default, contains shared folders from other persons.
    olReadOnlyGroup                        =	6		# Identifies a navigation group that is, by default, read-only and no folders can be added or removed from that navigation group. This does not imply the folders themselves are read-only, and write access to the folders depends on how the folders are set up.
    olRoomsGroup                           =	5		# Identifies the  Rooms navigation group in the Calendar navigation module.

    #values from outlook.olhorizontallayout
    #****************************
    olHorizontalLayoutAlignCenter          =	1		# Align controls horizontally by the center of each control.
    olHorizontalLayoutGrow                 =	3		# Resize control horizontally with the form.
    olHorizontalLayoutAlignLeft            =	0		# Align controls horizontally by the left edge of each control.
    olHorizontalLayoutAlignRight           =	2		# Align controls horizontally by the right edge of each control.

    #values from outlook.oliconviewplacement
    #****************************
    olIconAutoArrange                      =	2		# Icons are automatically lined up and arranged to prevent gaps or overlaps, but are not sorted.
    olIconDoNotArrange                     =	0		# Icons are not automatically sorted, lined up, or arranged.
    olIconLineUp                           =	1		# Icons are automatically lined up, but are not sorted or arranged.
    olIconSortAndAutoArrange               =	3		# Icons are automatically sorted, lined up, and arranged to prevent gaps or overlaps.

    #values from outlook.oliconviewtype
    #****************************
    olIconViewLarge                        =	0		# Displays Outlook items as large icons, with the description for the Outlook item below the icon.
    olIconViewList                         =	2		# Displays Outlook items as a list of small icons, with the description for the Outlook item next to the icon.
    olIconViewSmall                        =	1		# Displays Outlook items as a collection of small icons, with the description for the Outlook item next to the icon.

    #values from outlook.olimportance
    #****************************
    olImportanceHigh                       =	2		# Item is marked as high importance.
    olImportanceLow                        =	0		# Item is marked as low importance.
    olImportanceNormal                     =	1		# Item is marked as medium importance.

    #values from outlook.olinspectorclose
    #****************************
    olDiscard                              =	1		# Changes to the document are discarded.
    olPromptForSave                        =	2		# User is prompted to save documents.
    olSave                                 =	0		# Documents are saved.

    #values from outlook.olitemtype
    #****************************
    olAppointmentItem                      =	1		# An  AppointmentItem object.
    olContactItem                          =	2		# A  ContactItem object.
    olDistributionListItem                 =	7		# A  DistListItem object.
    olJournalItem                          =	4		# A  JournalItem object.
    olMailItem                             =	0		# A  MailItem object.
    olNoteItem                             =	5		# A  NoteItem object.
    olPostItem                             =	6		# A  PostItem object.
    olTaskItem                             =	3		# A  TaskItem object.

    #values from outlook.oljournalrecipienttype
    #****************************
    olAssociatedContact                    =	1		# The Contact associated with the Journal item.

    #values from outlook.olmailingaddress
    #****************************
    olBusiness                             =	2		# Business mailing address
    olHome                                 =	1		# Home mailing address
    olNone                                 =	0		# No mailing address defined
    olOther                                =	3		# Other mailing address

    #values from outlook.olmailrecipienttype
    #****************************
    olBCC                                  =	3		# The recipient is specified in the  BCC property of the Item.
    olCC                                   =	2		# The recipient is specified in the  CC property of the Item.
    olOriginator                           =	0		# Originator (sender) of the  Item.
    olTo                                   =	1		# The recipient is specified in the  To property of the Item.

    #values from outlook.olmarkinterval
    #****************************
    olMarkComplete                         =	5		# Mark the task as complete.
    olMarkNextWeek                         =	3		# Mark the task due next week.
    olMarkNoDate                           =	4		# Mark the task due with no date.
    olMarkThisWeek                         =	2		# Mark the task due this week.
    olMarkToday                            =	0		# Mark the task due today.
    olMarkTomorrow                         =	1		# Mark the task due tomorrow.

    #values from outlook.olmatchentry
    #****************************
    olMatchEntryComplete                   =	1		# Extended matching. As each character is typed, the control searches for an entry matching all characters entered.
    olMatchEntryFirstLetter                =	0		# Basic matching: The control searches for the next entry that starts with the character entered. Repeatedly typing the same letter cycles through all entries beginning with that letter.
    olMatchEntryNone                       =	2		# No matching is performed.

    #values from outlook.olmeetingrecipienttype
    #****************************
    olOptional                             =	2		# Optional attendee
    olOrganizer                            =	0		# Meeting organizer
    olRequired                             =	1		# Required attendee
    olResource                             =	3		# A resource such as a conference room

    #values from outlook.olmeetingresponse
    #****************************
    olMeetingAccepted                      =	3		# The meeting was accepted.
    olMeetingDeclined                      =	4		# The meeting was declined.
    olMeetingTentative                     =	2		# The meeting was tentatively accepted.

    #values from outlook.olmeetingstatus
    #****************************
    olMeeting                              =	1		# The meeting has been scheduled.
    olMeetingCanceled                      =	5		# The scheduled meeting has been cancelled.
    olMeetingReceived                      =	3		# The meeting request has been received.
    olMeetingReceivedAndCanceled           =	7		# The scheduled meeting has been cancelled but still appears on the user's calendar.
    olNonMeeting                           =	0		# An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.

    #values from outlook.olmousebutton
    #****************************
    olMouseButtonLeft                      =	1		# Indicates that the primary (left) button on a mouse is pressed during the event.
    olMouseButtonRight                     =	2		# Indicates that the secondary (right) button on a mouse is pressed during the event.
    olMouseButtonMiddle                    =	4		# Indicates that the middle button on a mouse is pressed during the event.

    #values from outlook.olmousepointer
    #****************************
    olMousePointerAppStarting              =	13		# The cursor is the arrow and the hourglass, representing that the application is starting.
    olMousePointerArrow                    =	1		# The cursor is the standard arrow.
    olMousePointerCross                    =	2		# The cursor is the cross.
    olMousePointerCustom                   =	99		# The cursor is represented by the custom cursor bitmap.
    olMousePointerDefault                  =	0		# The default cursor.
    olMousePointerHelp                     =	14		# The cursor is the What's This help cursor.
    olMousePointerHourGlass                =	11		# The cursor is the hourglass, representing that the application is busy.
    olMousePointerIBeam                    =	3		# The cursor is the I-beam text entry cursor.
    olMousePointerNoDrop                   =	12		# The cursor is the no-drop cursor.
    olMousePointerSizeAll                  =	15		# Resize cursor in all directions.
    olMousePointerSizeNESW                 =	6		# Resize cursor from Northeast to Southwest.
    olMousePointerSizeNS                   =	7		# Resize cursor from North to South.
    olMousePointerSizeNWSE                 =	8		# Resize cursor from Northwest to Southeast.
    olMousePointerSizeWE                   =	9		# Resize cursor from West to East.
    olMousePointerUpArrow                  =	10		# The cursor is the vertical arrow.

    #values from outlook.olmultiline
    #****************************
    olAlwaysMultiLine                      =	2		# Multiple lines are always displayed.
    olAlwaysSingleLine                     =	1		# Single lines are always displayed.
    olWidthMultiLine                       =	0		# Single lines are displayed unless the length of the line, in characters, is greater than the value for the  MultiLineWidth property of the TableView object, at which point multiple lines are displayed.

    #values from outlook.olmultiselect
    #****************************
    olMultiSelectExtended                  =	2		# Supports  SHIFT and CTRL to select one or more items at a time. Pressing SHIFT and clicking the mouse, or pressing SHIFT and one of the arrow keys, extends the selection from the previously selected item to the current item. Pressing CTRL and clicking the mouse toggles the selection of an item.
    olMultiSelectMulti                     =	1		# Supports selection of one or more items at a time. Pressing  SPACEBAR or clicking the mouse toggles the selection of an item in the list.
    olMultiSelectSingle                    =	0		# Supports selection of only one item at a time.

    #values from outlook.olnavigationmoduletype
    #****************************
    olModuleCalendar                       =	1		# A  CalendarModule object that represents the Calendar navigation module.
    olModuleContacts                       =	2		# A  ContactsModule object that represents the Contacts navigation module.
    olModuleFolderList                     =	6		# A  NavigationModule object that represents the Folders List navigation module.
    olModuleJournal                        =	4		# A  JournalModule object that represents the Journal navigation module.
    olModuleMail                           =	0		# A  MailModule object that represents the Mail navigation module.
    olModuleNotes                          =	5		# A  NotesModule object that represents the Notes navigation module.
    olModuleShortcuts                      =	7		# A  NavigationModule object that represents the Shortcuts navigation module.
    olModuleSolutions                      =	8		# A  SolutionsModule object that represents the Solutions navigation module.
    olModuleTasks                          =	3		# A  TasksModule object that represents the Tasks navigation module.

    #values from outlook.olobjectclass
    #****************************
    olAccount                              =	105		# An  Account object.
    olAccountRuleCondition                 =	135		# An  AccountRuleCondition object.
    olAccounts                             =	106		# An  Accounts object.
    olAction                               =	32		# An  Action object.
    olActions                              =	33		# An  Actions object.
    olAddressEntries                       =	21		# An  AddressEntries object.
    olAddressEntry                         =	8		# An  AddressEntry object.
    olAddressList                          =	7		# An  AddressList object.
    olAddressLists                         =	20		# An  AddressLists object.
    olAddressRuleCondition                 =	170		# An  AddressRuleCondition object.
    olApplication                          =	0		# An  Application object.
    olAppointment                          =	26		# An  AppointmentItem object.
    olAssignToCategoryRuleAction           =	122		# An  AssignToCategoryRuleAction object.
    olAttachment                           =	5		# An  Attachment object.
    olAttachments                          =	18		# An  Attachments object.
    olAttachmentSelection                  =	169		# An  AttachmentSelection object.
    olAutoFormatRule                       =	147		# An  AutoFormatRule object.
    olAutoFormatRules                      =	148		# An  AutoFormatRules object.
    olCalendarModule                       =	159		# A  CalendarModule object.
    olCalendarSharing                      =	151		# A  CalendarSharing object.
    olCategories                           =	153		# A  Categories object.
    olCategory                             =	152		# A  Category object.
    olCategoryRuleCondition                =	130		# A  CategoryRuleCondition object.
    olClassBusinessCardView                =	168		# A  BusinessCardView object.
    olClassCalendarView                    =	139		# A  CalendarView object.
    olClassCardView                        =	138		# A  CardView object.
    olClassIconView                        =	137		# An  IconView object.
    olClassNavigationPane                  =	155		# A  NavigationPane object.
    olClassPeopleView                      =	183		# A  PeopleView object.
    olClassTableView                       =	136		# A  TableView object.
    olClassTimeLineView                    =	140		# A  TimelineView object.
    olClassTimeZone                        =	174		# A  TimeZone object.
    olClassTimeZones                       =	175		# A  TimeZones object.
    olColumn                               =	154		# A  Column object.
    olColumnFormat                         =	149		# A  ColumnFormat object.
    olColumns                              =	150		# A  Columns object.
    olConflict                             =	102		# A  Conflict object.
    olConflicts                            =	103		# A  Conflicts object.
    olContact                              =	40		# A  ContactItem object.
    olContactsModule                       =	160		# A  ContactsModule object.
    olConversation                         =	178		# A  Conversation object.
    olConversationHeader                   =	182		# A  ConversationHeader object.
    olDistributionList                     =	69		# An  ExchangeDistributionList object.
    olDocument                             =	41		# A  DocumentItem object.
    olException                            =	30		# An  Exception object.
    olExceptions                           =	29		# An  Exceptions object.
    olExchangeDistributionList             =	111		# An  ExchangeDistributionList object.
    olExchangeUser                         =	110		# An  ExchangeUser object.
    olExplorer                             =	34		# An  Explorer object.
    olExplorers                            =	60		# An  Explorers object.
    olFolder                               =	2		# A  Folder object.
    olFolders                              =	15		# A  Folders object.
    olFolderUserProperties                 =	172		# A  UserDefinedProperties object.
    olFolderUserProperty                   =	171		# A  UserDefinedProperty object.
    olFormDescription                      =	37		# A  FormDescription object.
    olFormNameRuleCondition                =	131		# A  FormNameRuleCondition object.
    olFormRegion                           =	129		# A  FormRegion object.
    olFromRssFeedRuleCondition             =	173		# A  FromRssFeedRuleCondition object.
    olFromRuleCondition                    =	132		# A  ToOrFromRuleCondition object.
    olImportanceRuleCondition              =	128		# An  ImportanceRuleCondition object.
    olInspector                            =	35		# An  Inspector object.
    olInspectors                           =	61		# An  Inspectors object.
    olItemProperties                       =	98		# An  ItemProperties object.
    olItemProperty                         =	99		# An  ItemProperty object.
    olItems                                =	16		# An  Items object.
    olJournal                              =	42		# A  JournalItem object.
    olJournalModule                        =	162		# A  JournalModule object.
    olMail                                 =	43		# A  MailItem object.
    olMailModule                           =	158		# A  MailModule object.
    olMarkAsTaskRuleAction                 =	124		# A  MarkAsTaskRuleAction object.
    olMeetingCancellation                  =	54		# A  MeetingItem object that is a meeting cancellation notice.
    olMeetingForwardNotification           =	181		# A  MeetingItem object that is a notice about forwarding the meeting request.
    olMeetingRequest                       =	53		# A  MeetingItem object that is a meeting request.
    olMeetingResponseNegative              =	55		# A  MeetingItem object that is a refusal of a meeting request.
    olMeetingResponsePositive              =	56		# A  MeetingItem object that is an acceptance of a meeting request.
    olMeetingResponseTentative             =	57		# A  MeetingItem object that is a tentative acceptance of a meeting request.
    olMoveOrCopyRuleAction                 =	118		# A  MoveOrCopyRuleAction object.
    olNamespace                            =	1		# A  NameSpace object.
    olNavigationFolder                     =	167		# A  NavigationFolder object.
    olNavigationFolders                    =	166		# A  NavigationFolders object.
    olNavigationGroup                      =	165		# A  NavigationGroup object.
    olNavigationGroups                     =	164		# A  NavigationGroups object.
    olNavigationModule                     =	157		# A  NavigationModule object.
    olNavigationModules                    =	156		# A  NavigationModules object.
    olNewItemAlertRuleAction               =	125		# A  NewItemAlertRuleAction object.
    olNote                                 =	44		# A  NoteItem object.
    olNotesModule                          =	163		# A  NotesModule object.
    olOrderField                           =	144		# An  OrderField object.
    olOrderFields                          =	145		# An  OrderFields object.
    olOutlookBarGroup                      =	66		# An  OutlookBarGroup object.
    olOutlookBarGroups                     =	65		# An  OutlookBarGroups object.
    olOutlookBarPane                       =	63		# An  OutlookBarPane object.
    olOutlookBarShortcut                   =	68		# An  OutlookBarShortcut object.
    olOutlookBarShortcuts                  =	67		# An  OutlookBarShortcuts object.
    olOutlookBarStorage                    =	64		# An  OutlookBarStorage object.
    olOutspace                             =	180		# An  AccountSelector object.
    olPages                                =	36		# A  Pages object.
    olPanes                                =	62		# A  Panes object.
    olPlaySoundRuleAction                  =	123		# A  PlaySoundRuleAction object.
    olPost                                 =	45		# A  PostItem object.
    olPropertyAccessor                     =	112		# A  PropertyAccessor object.
    olPropertyPages                        =	71		# A  PropertyPages object.
    olPropertyPageSite                     =	70		# A  PropertyPageSite object.
    olRecipient                            =	4		# A  Recipient object.
    olRecipients                           =	17		# A  Recipients object.
    olRecurrencePattern                    =	28		# A  RecurrencePattern object.
    olReminder                             =	101		# A  Reminder object.
    olReminders                            =	100		# A  Reminders object.
    olRemote                               =	47		# A  RemoteItem object.
    olReport                               =	46		# A  ReportItem object.
    olResults                              =	78		# A  Results object.
    olRow                                  =	121		# A  Row object.
    olRule                                 =	115		# A  Rule object.
    olRuleAction                           =	117		# A  RuleAction object.
    olRuleActions                          =	116		# A  RuleActions object.
    olRuleCondition                        =	127		# A  RuleCondition object.
    olRuleConditions                       =	126		# A  RuleConditions object.
    olRules                                =	114		# A  Rules object.
    olSearch                               =	77		# A  Search object.
    olSelection                            =	74		# A  Selection object.
    olSelectNamesDialog                    =	109		# A  SelectNamesDialog object.
    olSenderInAddressListRuleCondition     =	133		# A  SenderInAddressListRuleCondition object.
    olSendRuleAction                       =	119		# A  SendRuleAction object.
    olSharing                              =	104		# A  SharingItem object.
    olSimpleItems                          =	179		# A  SimpleItems object.
    olSolutionsModule                      =	177		# A  SolutionsModule object.
    olStorageItem                          =	113		# A  StorageItem object.
    olStore                                =	107		# A  Store object.
    olStores                               =	108		# A  Stores object.
    olSyncObject                           =	72		# A  SyncObject object.
    olSyncObjects                          =	73		# A  SyncObjects object.
    olTable                                =	120		# A  Table object.
    olTask                                 =	48		# A  TaskItem object.
    olTaskRequest                          =	49		# A  TaskRequestItem object.
    olTaskRequestAccept                    =	51		# A  TaskRequestAcceptItem object.
    olTaskRequestDecline                   =	52		# A  TaskRequestDeclineItem object.
    olTaskRequestUpdate                    =	50		# A  TaskRequestUpdateItem object.
    olTasksModule                          =	161		# A  TasksModule object.
    olTextRuleCondition                    =	134		# A  TextRuleCondition object.
    olUserDefinedProperties                =	172		# A  UserDefinedProperties object.
    olUserDefinedProperty                  =	171		# A  UserDefinedProperty object.
    olUserProperties                       =	38		# A  UserProperties object.
    olUserProperty                         =	39		# A  UserProperty object.
    olView                                 =	80		# A  View object.
    olViewField                            =	142		# A  ViewField object.
    olViewFields                           =	141		# A  ViewFields object.
    olViewFont                             =	146		# A  ViewFont object.
    olViews                                =	79		# A  Views object.

    #values from outlook.oloutlookbarviewtype
    #****************************
    olLargeIcon                            =	0		# The  Outlook Bar group displays large icons.
    olSmallIcon                            =	1		# The  Outlook Bar group displays small icons.

    #values from outlook.olpagetype
    #****************************
    olPageTypePlanner                      =	0		# The free/busy scheduling grid from a meeting request form.
    olPageTypeTracker                      =	1		# The tracking grid from a meeting request that has been sent.

    #values from outlook.olpane
    #****************************
    olFolderList                           =	2		# The folder list pane
    olNavigationPane                       =	4		# The navigation pane
    olOutlookBar                           =	1		# The Outlook bar pane
    olPreview                              =	3		# The reading pane

    #values from outlook.olpermission
    #****************************
    olDoNotForward                         =	1		# Item cannot be forwarded.
    olPermissionTemplate                   =	2		# Outlook will use an Information Rights Management (IRM) template to determine the access and usage permissions for the item. See  MailItem.PermissionService and SharingItem.PermissionService properties.
    olUnrestricted                         =	0		# Item has no permission restrictions.

    #values from outlook.olpermissionservice
    #****************************
    olPassport                             =	2		# Microsoft Passport Network permissions will be used.
    olUnknown                              =	0		# Permission service is unknown.
    olWindows                              =	1		# Microsoft Windows permissions will be used.

    #values from outlook.olpicturealignment
    #****************************
    olPictureAlignmentLeft                 =	0		# The image is aligned to the left of the text and vertically centered on the button.
    olPictureAlignmentTop                  =	1		# The image is aligned to the right of the text and horizontally centered on the button.

    #values from outlook.olrecipientselectors
    #****************************
    olShowNone                             =	0		# No edit box will be displayed.
    olShowTo                               =	1		# Only an edit box for  To recipients will be displayed.
    olShowToCc                             =	2		# Only edit boxes for  To and Cc recipients will be displayed.
    olShowToCcBcc                          =	3		# Edit boxes for  To,  Cc, and  Bcc recipients will be displayed.

    #values from outlook.olrecurrencestate
    #****************************
    olApptException                        =	3		# The appointment is an exception to a recurrence pattern defined by a master appointment.
    olApptMaster                           =	1		# The appointment is a master appointment.
    olApptNotRecurring                     =	0		# The appointment is not a recurring appointment.
    olApptOccurrence                       =	2		# The appointment is an occurrence of a recurring appointment defined by a master appointment.

    #values from outlook.olrecurrencetype
    #****************************
    olRecursDaily                          =	0		# Represents a daily recurrence pattern.
    olRecursMonthly                        =	2		# Represents a monthly recurrence pattern.
    olRecursMonthNth                       =	3		# Represents a MonthNth recurrence pattern. See  RecurrencePattern.Instance property.
    olRecursWeekly                         =	1		# Represents a weekly recurrence pattern.
    olRecursYearly                         =	5		# Represents a yearly recurrence pattern.
    olRecursYearNth                        =	6		# Represents a YearNth recurrence pattern. See  RecurrencePattern.Instance property.

    #values from outlook.olreferencetype
    #****************************
    olStrong                               =	1		# Strong reference
    olWeak                                 =	0		# Weak reference

    #values from outlook.olremotestatus
    #****************************
    olMarkedForCopy                        =	3		# Item is marked to be copied.
    olMarkedForDelete                      =	4		# Item is marked for deletion.
    olMarkedForDownload                    =	2		# Item is marked for download.
    olRemoteStatusNone                     =	0		# No remote status has been set.
    olUnMarked                             =	1		# Item is not marked.

    #values from outlook.olresponsestatus
    #****************************
    olResponseAccepted                     =	3		# Meeting accepted.
    olResponseDeclined                     =	4		# Meeting declined.
    olResponseNone                         =	0		# The appointment is a simple appointment and does not require a response.
    olResponseNotResponded                 =	5		# Recipient has not responded.
    olResponseOrganized                    =	1		# The  AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting.
    olResponseTentative                    =	2		# Meeting tentatively accepted.

    #values from outlook.olruleactiontype
    #****************************
    olRuleActionAssignToCategory           =	2		# Rule action is to assign categories to the message.
    olRuleActionCcMessage                  =	27		# Rule action is to cc the message to specified recipients.
    olRuleActionClearCategories            =	30		# Rule action is to clear all the categories assigned to the message.
    olRuleActionCopyToFolder               =	5		# Rule action is to copy the message to a specified folder.
    olRuleActionCustomAction               =	22		# Rule action is to perform a custom action.
    olRuleActionDefer                      =	28		# Rule action is to defer delivery of the message by the specified number of minutes.
    olRuleActionDelete                     =	3		# Rule action is to delete the message.
    olRuleActionDeletePermanently          =	4		# Rule action is to permanently delete the message.
    olRuleActionDesktopAlert               =	24		# Rule action is to display a desktop alert.
    olRuleActionFlagClear                  =	13		# Rule action is to clear the message flag.
    olRuleActionFlagColor                  =	12		# Rule action is to flag the message with a specified colored flag.
    olRuleActionFlagForActionInDays        =	11		# Rule action is to flag the message for action in the specified number of days.
    olRuleActionForward                    =	6		# Rule action is to forward the message to the specified recipients.
    olRuleActionForwardAsAttachment        =	7		# Rule action is to forward the message as an attachment to the specified recipients.
    olRuleActionImportance                 =	14		# Rule action is to mark the message with the specified level of importance.
    olRuleActionMarkAsTask                 =	41		# Rule action is to mark the message as a task.
    olRuleActionMarkRead                   =	19		# Rule action is to mark the message as read.
    olRuleActionMoveToFolder               =	1		# Rule action is to move the message to the specified folder.
    olRuleActionNewItemAlert               =	23		# Rule action is to display the specified text in the  New Item Alert dialog box.
    olRuleActionNotifyDelivery             =	26		# Rule action is to request delivery notification for the message being sent.
    olRuleActionNotifyRead                 =	25		# Rule action is to request read notification for the message being sent.
    olRuleActionPlaySound                  =	17		# Rule action is to play a sound file.
    olRuleActionPrint                      =	16		# Rule action is to print the message on the default printer.
    olRuleActionRedirect                   =	8		# Rule action is to redirect the message to the specified recipients.
    olRuleActionRunScript                  =	20		# Rule action is to run a script.
    olRuleActionSensitivity                =	15		# Rule action is to mark the message with the specified level of sensitivity.
    olRuleActionServerReply                =	9		# Rule action is to request the server to reply with the specified mail item.
    olRuleActionStartApplication           =	18		# Rule action is to run an .exe file.
    olRuleActionStop                       =	21		# Rule action is to stop processing more rules.
    olRuleActionTemplate                   =	10		# Rule action is to use the specified template (.oft) file as a form template.
    olRuleActionUnknown                    =	0		# Unrecognized rule action.

    #values from outlook.olruleconditiontype
    #****************************
    olConditionAccount                     =	3		# Account is the account specified in AccountRuleCondition.Account.
    olConditionAnyCategory                 =	29		# Message is assigned to any category.
    olConditionBody                        =	13		# Body contains words specified in  TextRuleCondition.Text.
    olConditionBodyOrSubject               =	14		# Body or subject contains words specified by  TextRuleCondition.Text.
    olConditionCategory                    =	18		# Category is the category specified in CategoryRuleCondition.Categories.
    olConditionCc                          =	9		# Message has my name in the  Cc box.
    olConditionDateRange                   =	22		# Message was received between x and y, where x and y are  Date values.
    olConditionFlaggedForAction            =	8		# Message is flagged for the specified action.
    olConditionFormName                    =	23		# Message uses the form specified in  FormNameRuleCondition.FormName.
    olConditionFrom                        =	1		# Sender is in the recipient list specified in  ToOrFromRuleCondition.Recipients.
    olConditionFromAnyRssFeed              =	31		# Message is generated from any RSS subscription.
    olConditionFromRssFeed                 =	30		# Message is generated from a specific RSS subscription.
    olConditionHasAttachment               =	20		# Message has one or more attachments.
    olConditionImportance                  =	6		# Message is marked with the specified level of importance.
    olConditionLocalMachineOnly            =	27		# Rule can run only on the local machine.
    olConditionMeetingInviteOrUpdate       =	26		# Message is a meeting invitation or update.
    olConditionMessageHeader               =	15		# Message header contains words specified in  TextRuleCondition.Text.
    olConditionNotTo                       =	11		# Message does not have my name in the  To box.
    olConditionOnlyToMe                    =	4		# Message is sent only to me.
    olConditionOOF                         =	19		# Message is an out-of-office message.
    olConditionOtherMachine                =	28		# Rule can run only on a specific machine that is not the current machine.
    olConditionProperty                    =	24		# Document property is exactly, contains, or does not contain specified properties.
    olConditionRecipientAddress            =	16		# Recipient address contains words specified in  TextRuleCondition.Text.
    olConditionSenderAddress               =	17		# Sender address contains words specified in  TextRuleCondition.Text.
    olConditionSenderInAddressBook         =	25		# Sender is in the address list specified in  AddressRuleCondition.Address.
    olConditionSensitivity                 =	7		# Message is marked with the specified level of sensitivity.
    olConditionSentTo                      =	12		# Sent to recipients (To,  Cc) are in the recipient list specified in  ToOrFromRuleCondition.Recipients.
    olConditionSizeRange                   =	21		# Message size is between x and y in units of KB, where x and y are  Integer values.
    olConditionSubject                     =	2		# Subject contains words specified in  TextRuleCondition.Text.
    olConditionTo                          =	5		# My name is in the  To box.
    olConditionToOrCc                      =	10		# Message has my name in the  To or Cc box.
    olConditionUnknown                     =	0		# Unrecognized condition.

    #values from outlook.olruleexecuteoption
    #****************************
    olRuleExecuteAllMessages               =	0		# Executes a rule against all messages in the specified folder or folders.
    olRuleExecuteReadMessages              =	1		# Executes a rule against messages that have been read in the specified folder or folders.
    olRuleExecuteUnreadMessages            =	2		# Executes a rule against messages that have not been read in the specified folder or folders.

    #values from outlook.olruletype
    #****************************
    olRuleReceive                          =	0		# Indicates that the rule is applied to messages that are being received.
    olRuleSend                             =	1		# Indicates that the rule is being applied to messages being sent.

    #values from outlook.olsaveastype
    #****************************
    olDoc                                  =	4		# Microsoft Office Word format (.doc)
    olHTML                                 =	5		# HTML format (.html)
    olICal                                 =	8		# iCal format (.ics)
    olMHTML                                =	10		# MIME HTML format (.mht)
    olMSG                                  =	3		# Outlook message format (.msg)
    olMSGUnicode                           =	9		# Outlook Unicode message format (.msg)
    olRTF                                  =	1		# Rich Text format (.rtf)
    olTemplate                             =	2		# Microsoft Outlook template (.oft)
    olTXT                                  =	0		# Text format (.txt)
    olVCal                                 =	7		# VCal format (.vcs)
    olVCard                                =	6		# VCard format (.vcf)

    #values from outlook.olscrollbars
    #****************************
    olScrollBarsBoth                       =	3		# Display both the horizontal and vertical scroll bars as necessary.
    olScrollBarsHorizontal                 =	1		# Display a horizontal scroll bar only.
    olScrollBarsNone                       =	0		# Display no scroll bars.
    olScrollBarsVertical                   =	2		# Display a vertical scroll bar only.

    #values from outlook.olsearchscope
    #****************************
    olSearchScopeAllFolders                =	1		# The search scope is across all folders that have the same folder type as the current folder (Folder.DefaultItemType), and all stores that have been selected for search.
    olSearchScopeAllOutlookItems           =	2		# The search scope is all Outlook items in all folders in stores that have been selected for search.
    olSearchScopeCurrentFolder             =	0		# The search scope is the folder represented by  Explorer.CurrentFolder, and only that folder.
    olSearchScopeCurrentStore              =	4		# The search scope is the store for the current folder, which contains the item displayed in the active explorer.
    olSearchScopeSubfolders                =	3		# The search scope is the folder represented by  Explorer.CurrentFolder and its subfolders.

    #values from outlook.olselectioncontents
    #****************************
    olConversationHeaders                  =	1		# Conversation header or headers in a selection.

    #values from outlook.olselectionlocation
    #****************************
    olAttachmentWell                       =	4		# The selection is an attachment of an item in the  Reading Pane or inspector.
    olDailyTaskList                        =	3		# The selection is in the daily  Tasks list in the calendar view.
    olToDoBarAppointmentList               =	2		# The selection is in the list of appointments in the  To-Do Bar.
    olToDoBarTaskList                      =	1		# The selection is in the list of tasks in the  To-Do Bar.
    olViewList                             =	0		# The selection is in a list of items in an explorer.

    #values from outlook.olsensitivity
    #****************************
    olConfidential                         =	3		# Confidential
    olNormal                               =	0		# Normal sensitivity
    olPersonal                             =	1		# Personal
    olPrivate                              =	2		# Private

    #values from outlook.olsharingmsgtype
    #****************************
    olSharingMsgTypeInvite                 =	2		# Represents a sharing invitation.
    olSharingMsgTypeInviteAndRequest       =	3		# Represents both a sharing invitation and a sharing request.
    olSharingMsgTypeRequest                =	1		# Represents a sharing request.
    olSharingMsgTypeResponseAllow          =	4		# Represents a sharing response, which indicates that a sharing request or sharing invitation has been allowed.
    olSharingMsgTypeResponseDeny           =	5		# Represents a sharing response, which indicates that a sharing request or sharing invitation has been denied.
    olSharingMsgTypeUnknown                =	0		# Represents an unknown type of sharing message.

    #values from outlook.olsharingprovider
    #****************************
    olProviderExchange                     =	1		# Represents the Exchange sharing provider.
    olProviderFederate                     =	7		# Represents a federated sharing provider. A  SharingItem object with this type of provider is used for sharing relationships across organizational boundares (for example, between two organizations using Microsoft Exchange Server 2010).
    olProviderICal                         =	4		# Represents the iCalendar sharing provider.
    olProviderPubCal                       =	3		# Represents the PubCal sharing provider.
    olProviderRSS                          =	6		# Represents the Really Simple Syndication (RSS) sharing provider.
    olProviderSharePoint                   =	5		# Represents the SharePoint sharing provider.
    olProviderUnknown                      =	0		# Represents an unknown sharing provider. This value is used if the sharing provider GUID in the sharing message does not match the GUID of any of the sharing providers represented in this enumeration.
    olProviderWebCal                       =	2		# Represents the WebCal sharing provider.

    #values from outlook.olshiftstate
    #****************************
    olShiftStateShiftMask                  =	1		# Indicates that the SHIFT key is pressed during the event.
    olShiftStateCtrlMask                   =	2		# Indicates that the CTRL key is pressed during the event.
    olShiftStateAltMask                    =	4		# Indicates that the ALT key is pressed during the event.

    #values from outlook.olshowitemcount
    #****************************
    olNoItemCount                          =	0		# No item count displayed.
    olShowTotalItemCount                   =	2		# Shows count of total number of items.
    olShowUnreadItemCount                  =	1		# Shows count of unread items.

    #values from outlook.olsolutionscope
    #****************************
    olHideInDefaultModules                 =	0		# The solution root and its subfolders are displayed in the  Solutions module and the Folder List.
    olShowInDefaultModules                 =	1		# The solution root and its subfolders are displayed in the  Solutions module and the Folder List, as well as in their respective default modules.

    #values from outlook.olsortorder
    #****************************
    olAscending                            =	1		# Ascending order
    olDescending                           =	2		# Descending order
    olSortNone                             =	0		# Not sorted

    #values from outlook.olspecialfolders
    #****************************
    olSpecialFolderAllTasks                =	0		# Specifies the  All Tasks search folder for a store.
    olSpecialFolderReminders               =	1		# Specifies the  Reminders search folder for a store.

    #values from outlook.olstorageidentifiertype
    #****************************
    olIdentifyByEntryID                    =	1		# Identifies a  StorageItem by EntryID.
    olIdentifyByMessageClass               =	2		# Identifies a  StorageItem by message class.
    olIdentifyBySubject                    =	0		# Identifies a  StorageItem by Subject.

    #values from outlook.olstoretype
    #****************************
    olStoreANSI                            =	3		# ANSI format personal folders file (.pst) compatible with all previous versions of Microsoft Outlook format.
    olStoreDefault                         =	1		# Default format compatible with the mailbox mode in which Outlook runs on the Microsoft Exchange Server.
    olStoreUnicode                         =	2		# Unicode format personal folders file (.pst) compatible with Microsoft Office Outlook 2003 and later.

    #values from outlook.olsyncstate
    #****************************
    olSyncStarted                          =	1		# Synchronization started
    olSyncStopped                          =	0		# Synchronization stopped

    #values from outlook.oltablecontents
    #****************************
    olHiddenItems                          =	1		# Only the hidden items in the folder.
    olUserItems                            =	0		# Only the non-hidden user items in the folder.

    #values from outlook.oltaskdelegationstate
    #****************************
    olTaskDelegationAccepted               =	2		# The delegate accepted the task.
    olTaskDelegationDeclined               =	3		# The delegate declined the task.
    olTaskDelegationUnknown                =	1		# The delegate response to the task is unknown.
    olTaskNotDelegated                     =	0		# The task has not been delegated.

    #values from outlook.oltaskownership
    #****************************
    olDelegatedTask                        =	1		# Task has been delegated to another user.
    olNewTask                              =	0		# Task has not yet been assigned to a user.
    olOwnTask                              =	2		# Task is assigned to the current Outlook user.

    #values from outlook.oltaskrecipienttype
    #****************************
    olFinalStatus                          =	3		# Indicates that the recipient will receive completion reports for the task.
    olUpdate                               =	2		# Indicates that the recipient will receive status updates for the task.

    #values from outlook.oltaskresponse
    #****************************
    olTaskAccept                           =	2		# Task accepted.
    olTaskAssign                           =	1		# Task reassigned.
    olTaskDecline                          =	3		# Task declined.
    olTaskSimple                           =	0		# Task is a simple task and cannot be accepted, declined, or assigned. This constant is not a valid parameter to the  TaskItem.Respond method.

    #values from outlook.oltaskstatus
    #****************************
    olTaskComplete                         =	2		# The task is complete.
    olTaskDeferred                         =	4		# The task is deferred.
    olTaskInProgress                       =	1		# The task is in progress.
    olTaskNotStarted                       =	0		# The task has not yet started.
    olTaskWaiting                          =	3		# The task is waiting on someone else.

    #values from outlook.oltextalign
    #****************************
    olTextAlignCenter                      =	2		# Text rendered in the control will be aligned to the center of the control.
    olTextAlignLeft                        =	1		# Text rendered in the control will be aligned to the left edge of the control.
    olTextAlignRight                       =	3		# Text rendered in the control will be aligned to the right edge of the control.

    #values from outlook.oltimelineviewmode
    #****************************
    olTimelineViewDay                      =	0		# Displays a timeline in which the upper scale represents days and the lower scale represents hours.
    olTimelineViewMonth                    =	2		# Displays a timeline in which the upper scale represents months and the lower scale represents days.
    olTimelineViewWeek                     =	1		# Displays a timeline in which the upper scale represents weeks and the lower scale represents days.

    #values from outlook.oltimestyle
    #****************************
    olTimeStyleShortDuration               =	4		# The drop-down portion of the time control displays only duration values with the interval set by the  OlkTimeControl.IntervalTime property.
    olTimeStyleTimeDuration                =	1		# The drop-down portion of the time control displays time values starting from the  ReferenceTime and uses the OlkTimeControl.IntervalTime property as the increment. The edit box of the time control displays the duration from the ReferenceTime to the selected time.
    olTimeStyleTimeOnly                    =	0		# The drop-down portion of the time control displays only time values with the interval set by the  OlkTimeControl.IntervalTime property.

    #values from outlook.oltrackingstatus
    #****************************
    olTrackingDelivered                    =	1		# The item has been delivered to the recipient.
    olTrackingNone                         =	0		# No tracking information is available for the recipient.
    olTrackingNotDelivered                 =	2		# The item has not been delivered to the recipient.
    olTrackingNotRead                      =	3		# The item has not been read by the recipient.
    olTrackingRead                         =	6		# The item has been read by the recipient.
    olTrackingRecallFailure                =	4		# The sender of the item attempted to recall the item but was unsuccessful.
    olTrackingRecallSuccess                =	5		# The sender of the item recalled the item.
    olTrackingReplied                      =	7		# The recipient replied to the item.

    #values from outlook.olunifiedgroupfoldertype
    #****************************


    #values from outlook.olunifiedgrouptype
    #****************************


    #values from outlook.oluserpropertytype
    #****************************
    olCombination                          =	19		# The property type is a combination of other types. It corresponds to the MAPI type  PT_STRING8.
    olCurrency                             =	14		# Represents a  Currency property type. It corresponds to the MAPI type PT_CURRENCY.
    olDateTime                             =	5		# Represents a  DateTime property type. It corresponds to the MAPI type PT_SYSTIME.
    olDuration                             =	7		# Represents a time duration property type. It corresponds to the MAPI type  PT_LONG.
    olEnumeration                          =	21		# Represents an enumeration property type. It corresponds to the MAPI type  PT_LONG.
    olFormula                              =	18		# Represents a formula property type. It corresponds to the MAPI type  PT_STRING8. See  UserDefinedProperty.Formula property.
    olInteger                              =	20		# Represents an  Integer number property type. It corresponds to the MAPI type PT_LONG.
    olKeywords                             =	11		# Represents a  String array property type used to store keywords. It corresponds to the MAPI type PT_MV_STRING8.
    olNumber                               =	3		# Represents a  Double number property type. It corresponds to the MAPI type PT_DOUBLE.
    olOutlookInternal                      =	0		# Represents an Outlook internal property type.
    olPercent                              =	12		# Represents a  Double number property type used to store a percentage. It corresponds to the MAPI type PT_LONG.
    olSmartFrom                            =	22		# Represents a smart from property type. This property indicates that if the  From property of an Outlook item is empty, then the To property should be used instead.
    olText                                 =	1		# Represents a  String property type. It corresponds to the MAPI type PT_STRING8.
    olYesNo                                =	6		# Represents a yes/no (Boolean) property type. It corresponds to the MAPI type  PT_BOOLEAN.

    #values from outlook.olverticallayout
    #****************************
    olVerticalLayoutAlignBottom            =	2		# Align controls vertically by the bottom edge of each control.
    olVerticalLayoutAlignGrow              =	3		# Resize control vertically with the form.
    olVerticalLayoutAlignMiddle            =	1		# Align controls vertically by the center of each control.
    olVerticalLayoutAlignTop               =	0		# Align controls vertically by the top edge of each control.

    #values from outlook.olviewsaveoption
    #****************************
    olViewSaveOptionAllFoldersOfType       =	2		# Indicates that the view is available in all folders of the same type.
    olViewSaveOptionThisFolderEveryone     =	0		# Indicates that the view is only available in the current folder and is available to all users.
    olViewSaveOptionThisFolderOnlyMe       =	1		# Indicates that the view is only available in the current folder and is only available to the current Outlook user.

    #values from outlook.olviewtype
    #****************************
    olBusinessCardView                     =	5		# Represents a  BusinessCardView object.
    olCalendarView                         =	2		# Represents a  CalendarView object.
    olCardView                             =	1		# Represents a  CardView object.
    olDailyTaskListView                    =	1		# Represents the  TableView object that contains the daily task list in a calendar view.
    olIconView                             =	3		# Represents an IconView object.
    olPeopleView                           =	7		# Represents a PeopleView object.
    olTableView                            =	0		# Represents a  TableView object.
    olTimelineView                         =	4		# Represents a  TimelineView object.

    #values from outlook.olwindowstate
    #****************************
    olMaximized                            =	0		# The window is maximized.
    olMinimized                            =	1		# The window is minimized.
    olNormalWindow                         =	2		# The window is in the normal state (not minimized or maximized).
}
#End Enum

$ol = new-object PSCustomObject -Property $ol

