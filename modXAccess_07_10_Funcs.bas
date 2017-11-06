Attribute VB_Name = "modXAccess_07_10_Funcs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modXAccess_07_10_Funcs"

'VGC 11/23/2016: CHANGES!

' **************************************
' ** VGC 01/13/2013:
' ** NOTE REQUIRED ZZ_'S IN THIS MOD:
' **   zz_qry_System_19_01
' **   zz_qry_System_19_02
' **   zz_qry_System_19_03
' **   zz_qry_System_70a
' **   zz_qry_System_70b
' **   zz_qry_System_72
' **************************************

' ** Issues for Vista and Access 2007 installations:
' ** 1. Myriad warnings messages! STOP! STOP!!
' **    Find a way to turn them off even in Access 2003 Runtime.
' ** 2. Don't show Ribbon, but DO show our custom Print CommandBar.
' ** 3. Don't show object NavigationPane.

' ** See IsWinVista() in modOperSysInfoFuncs1.

' ****************************************
' ** 1a. Remove warnings in Access 2007:
' ****************************************
' ** Setting the Trusted Location in Access 2007:
' ** 1. In Access 2007, click the Microsoft Office Button, and then click Access Options.
' ** 2. Click Trust Center.
' ** 3. Click Trust Center Settings.
' ** 4. Click Trusted Locations.
' ** 5. Click Add new location.
' ** 6. In the Microsoft Office Trusted Location dialog box, click Browse.
' ** 7. In the Browse dialog box, locate and select the folder that contains the Access database, and then click OK.
' ** 8. OR get a Trusted Publisher Certificate, also called 'Code Signing'.
' **    http://msdn.microsoft.com/en-us/library/ms995347.aspx

' ****************************************
' ** 1a. Remove warnings in Access 2003:
' ****************************************
' ** 1. In Access 2003, choose Tools -> Macro -> Security menu item.
' ** 2. Under the Security Level tab, choose Low.
' ** 3. OR get a Trusted Publisher Certificate, also called 'Code Signing'.
' **    http://msdn.microsoft.com/en-us/library/ms995347.aspx

'Access 2007 keywords for searching:
'NavigationPane
'TaskPane
'CustomUI : Ribbon Extensibility (or RibbonX)
'AutomationSecurity
'CommandBars
'MenuBar
'GetHiddenAttribute : Hidden check box on Properties window.
'SetHiddenAttribute : Hidden check box on Properties window.

' *****************************************
' ** DoCmd.LockNavigationPane Method:
' *****************************************
' ** You can use the LockNavigationPane action to prevent users from
' ** deleting database objects that are displayed in the Navigation Pane.
' ** Version Information
' **   Version Added:  Access 2007
' ** Syntax:
' **   expression.LockNavigationPane(Lock)
' **   expression   A variable that represents a DoCmd object.
' ** Parameters:
' **   Name    Required/Optional  Data Type  Description
' **   ======  =================  =========  ==========================================
' **   Lock    Required           Variant    Set to True to lock the Navigation Pane.
' ** Remarks:
' **   Locking the Navigation Pane prevents the user from deleting
' **   database objects or cutting database objects to the clipboard.
' **   It does not prevent the user from performing any of the following operations:
' **     1. Copying database objects to the clipboard
' **     2. Pasting database objects from the clipboard
' **     3. Displaying or hiding the Navigation Pane
' **     4. Selecting different Navigation Pane organization schemes
' **     5. Showing or hiding sections of the Navigation Pane
' *****************************************
'DoCmd.LockNavigationPane True  ' ** Lock the Navigation Pane.

' *****************************************
' ** DoCmd.SetDisplayedCategories Method:
' *****************************************
' ** Specifies which categories are displayed under Navigate
' ** to Category in the title bar of the Navigation Pane.
' ** Version Information
' **   Version Added:  Access 2007
' ** Syntax:
' **   expression.SetDisplayedCategories(Show, Category)
' **   expression   A variable that represents a DoCmd object.
' ** Parameters:
' **   Name      Required/Optional  Data Type  Description
' **   ========  =================  =========  ====================================================
' **   Show      Required           Variant    Set to Yes to show the category or categories.
' **                                           Set to No to hide them.
' **   Category  Optional           Variant    The name of the category you want to show or hide.
' **                                           Leave blank to show or hide all categories.
' ** Remarks:
' **   For example, if you want to prevent users from switching the Navigation Pane
' **   so that it displays objects sorted by Created Date, you can use this method
' **   to hide that option in the title bar's drop-down list.
' **   The caption in the title bar of the Navigation Pane indicates which filter, if any,
' **   is currently active. Click anywhere in the bar to display the drop-down list.
' **   The items that this method controls are listed under Navigate to Category.
' **   This method only enables or disables the display of the specified category
' **   or categories; it does not perform any switching of the Navigation Pane display.
' **   For example, if you are displaying objects sorted by Creation Date and you use the
' **   SetDisplayedCategories method to disable the Creation Date option, Access does not
' **   switch the Navigation Pane to another category.
' *****************************************
'DoCmd.SetDisplayedCategories No  ' ** Hide all categories.

' *****************************************
' ** DoCmd.ShowToolbar Method:
' *****************************************
' ** The ShowToolbar method carries out the ShowToolbar action in Visual Basic.
' ** Syntax:
' ** expression.ShowToolbar(ToolbarName, Show)
' ** expression   A variable that represents a DoCmd object.
' ** Parameters:
' **   Name         Required/Optional  Data Type      Description
' **   ===========  =================  =============  ===============================================================
' **   ToolbarName  Required           Variant        A string expression that's the valid name of a custom toolbar
' **                                                  you've created. If you run Visual Basic code containing the
' **                                                  ShowToolbar method in a library database, Microsoft Access
' **                                                  looks for the toolbar with this name first in the library
' **                                                  database, then in the current database.
' **   Show         Optional           AcShowToolbar  A AcShowToolbar constant that specifies whether to display
' **                                                  or hide the toolbar and in which views to display or hide it.
' **                                                  The default value is acToolbarYes.
' ** Remarks:
' ** You can use the ShowToolbar method to display or hide a custom toolbar.
' ** If you want to show a particular toolbar on just one form or report,
' ** you can set the OnActivate property of the form or report to the name
' ** of a macro that contains a ShowToolbar action to show the toolbar.
' ** Then set the OnDeactivate property of the form or report to the name
' ** of a macro that contains a ShowToolbar action to hide the toolbar.
' **
' ** AcShowToolbar Enumeration:
' **   0  acToolbarYes          Display the toolbar.
' **   1  acToolbarWhereApprop  Display the toolbar while in the appropriate view.
' **   2  acToolbarNo           Hide the toolbar.
' *****************************************

' ** How to: Hide the Ribbon When Access Starts:
' ** ===========================================
' ** To load the customized ribbon when Access starts, you should store its
' ** settings in a table named USysRibbons. The USysRibbons table must be
' ** created using specific column names in order for the Ribbon customizations
' ** to be implemented. The following table describes the settings to use when
' ** creating the USysRibbons table.
' **
' **   Column Name  Data Type  Description
' **   ===========  =========  ========================================================================================
' **   RibbonName   Text       Contains the name of the custom ribbon to be associated with this customization.
' **   RibbonXML    Memo       Contains the Ribbon Extensibility XML (RibbonX) that defines the Ribbon customization.
' **
' ** The following table describes the Ribbon customization settings to store in the USysRibbons table.
' **   Column Name  Value
' **   ===========  =============================================================================================================================
' **   RibbonName   HideTheRibbon
' **   RibbonXML    <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon startFromScratch="true"></ribbon></customUI>
' **
' ** http://msdn.microsoft.com/en-us/library/bb258192.aspx
' ** The USysRibbons table is considered a System table, and as such, is not normally visible.
' ** The Access Options must be set to show System Objects in order to see and edit it.
' ** Setting to show Hidden Objects will not reveal it.

' ** Customizing the 2007 Office Fluent Ribbon for Developers (Part 1 of 3)
' ** ======================================================================
' **
' ** UI Customization in Access 2007
' ** ===============================
' **
' ** RibbonX customizations in Access 2007 share some of the same options that the other Office
' ** applications have, but with some important differences. Just as with the other applications,
' ** you customize the Fluent UI in Access by using XML markup. And like the other applications,
' ** you can use external files that contain XML markup or COM add-ins to integrate Ribbon
' ** customizations into your application. However, unlike the other Office applications, because
' ** Access database files are binary and cannot be opened as Office Open XML Formats files, you
' ** cannot customize the Access Ribbon by adding parts to the database file.
' **
' ** Access does provide flexibility in customizing the Fluent UI. For example, customization
' ** markup can be stored in a table, embedded in a VBA procedure, stored in another Access
' ** database, or linked to from an Excel worksheet. You can also specify a custom UI for the
' ** application as a whole or for specific forms and reports.
' **
' ** The following scenarios can give you an idea of how to customize the Access UI.
' **
' **   Note: Because these walkthroughs involve changes to the database, you might want to perform
' **   these steps in a non-production database, perhaps by using a backup copy of a sample database.
' **
' ** Customizing the Fluent UI in Access:
' **
' ** When customizing the Access UI, you have two choices: You can store your customizations in a
' ** special table and have Access automatically load the markup for you, or you can store the
' ** customizations in a location of your choosing and load the markup manually by calling the
' ** Application.LoadCustomUI method.
' **
' ** If you choose to have Access load the customizations for you, you need to store them in a table
' ** named USysRibbons. The table should have at least two columns: a 255-character Text column named
' ** RibbonName, and a Memo column named RibbonXML. You place the Ribbon name in the RibbonName
' ** column, and the Ribbon markup in the RibbonXML column. After you close and re-open the database,
' ** you can select the default Ribbon to use in the Access Properties dialog box. You can select a
' ** Ribbon to appear when any form or report is selected as a property of the form or report.
' **
' ** If you decide to use a more dynamic technique, you can call the LoadCustomUI method, which loads
' ** Ribbon customizations whether the XML content is stored in a table or not. After you have loaded
' ** the customizations by calling LoadCustomUI, you can programmatically assign the named
' ** customization at run time.
' **
' **   Note: Customizations that you load by using the LoadCustomUI method are available only while
' **   the database is open. You need to call LoadCustomUI each time you open the database. This
' **   technique is useful for applications that need to assign custom UI programmatically at run time.
' **
' ** The signature for the LoadCustomUI method is as follows.
' **   expression .LoadCustomUI(CustomUIName As String, CustomUIXML As String)
' **   expression returns an Application object.
' **   CustomUIName is the name of the Custom Ribbon ID to be associated with this markup.
' **   CustomUIXML contains the XML customization markup.
' **
' ** There is an example of using the LoadCustomUI method later in this section.
' **
' ** The following procedure describes, in a generalized manner, how to add application-level
' ** customizations in Access. A later section includes a complete walkthrough.
' **
' ** To apply a customized application-level Ribbon at design time:
' **   1. Create a table named USysRibbons with columns as described earlier. Add rows for each
' **      different Ribbon you want to make available.
' **   2. Close and then reload the database.
' **   3. Click the Microsoft Office Button, and then click Access Options.
' **   4. Click the Current Database option and then, in the Ribbon and Toolbar Options section, click
' **      the Ribbon Name drop-down list and click one of the Ribbons.
' **
' ** To remove an existing customization (so that your database uses the default Fluent UI), delete the
' ** existing Ribbon name from the Ribbon Name list, and leave an empty value for the name of the Ribbon.
' **
' ** The following procedure describes, in a generalized manner, how to add form-level customizations
' ** or report-level customizations in Access.
' **
' ** To assign a specific custom Ribbon to a form or report:
' **   1. Follow the process described previously to make the customized Ribbon available to the
' **      application, by adding the USysRibbons table.
' **   2. Open the form or report in Design view.
' **   3. On the Design tab, click Property Sheet.
' **   4. In the Property window, on the Other tab, click the Ribbon Name drop-down list, and then
' **      click one of the Ribbons.
' **   5. Save, close, and then re-open the form or report. The UI you selected is displayed.
' **
' ** Note: The tabs displayed in the Fluent UI are additive. That is, unless you specifically hide
' ** the tabs or set the Start from Scratch attribute to True, the tabs displayed by a form or
' ** report's UI appear in addition to the existing tabs.
' **
' ** To explore this process further, work through the following examples.
' **
' ** The first part of the example sets an option that reports any errors that exist when you load
' ** custom UI (although you are performing these steps in Access, you can perform similar steps in
' ** other applications).
' **
' ** Creating an Access Application-Level Custom Ribbon
' ** ==================================================
' **
' ** To create an Access application-level custom ribbon:
' ** 1.  Start Access. Open an existing database, or create a new database.
' ** 2.  Click the Microsoft Office Button, click Access Options, and then click the Advanced tab.
' ** 3.  In the General section, select the option Show add-in user interface errors (this option
' **     might be in a different location, in different applications).
' ** 4.  Click OK to close the Access Options dialog box.
' **     Next, create a table that contains your customization XML markup.
' ** 5.  With a database open in Access, right-click the Navigation pane. Point to Navigation Options,
' **     and then click Show System Objects. (You cannot view the USysRibbons table in the Navigation
' **     pane unless this option is set.) Click OK to dismiss the dialog box.
' **     The Access system tables appear in the Navigation pane.
' ** 6.  On the Create tab, click Table Design.
' ** 7.  Add the following fields to the table.
' **       USysRibbons table field definitions:
' **       Field Name    Type
' **       ============  ============
' **       ID            AutoNumber
' **       RibbonName    Text
' **       RibbonXml     Memo
' ** 8.  Select the ID field. On the Design tab, select Primary Key.
' ** 9.  Click the Microsoft Office Button, and then click Save. Name the new table USysRibbons.
' ** 10. Right-click the USysRibbons tab, and then click Datasheet View.
' ** 11. Add the following data to the fields you created.
' **       USysRibbons table data:
' **       Field Name    Value
' **       ============  ==========================================================================
' **       ID            (AutoNumber)
' **       RibbonName    HideData
' **       RibbonXml     <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
' **                       <ribbon startFromScratch="false">
' **                         <tabs>
' **                           <tab idMso="TabCreate" visible="false"/>
' **                           <tab id="dbCustomTab" label="A Custom Tab" visible="true">
' **                             <group id="dbCustomGroup" label="A Custom Group">
' **                               <control idMso="Paste" label="Built-in Paste" enabled="true"/>
' **                             </group>
' **                           </tab>
' **                         </tabs>
' **                       </ribbon>
' **                     </customUI>
' **     This markup sets the startfromScratch attribute to False, and then hides the built-in
' **     Create tab. Next, it creates a custom tab and a custom group, and adds the built-in
' **     Paste control to the group.
' ** 12. Close the table.
' ** 13. Close and then re-open the database.
' ** 14. Click the Microsoft Office Button, and then click Access Options.
' ** 15. Click the Current Database tab, and scroll down until you find the Ribbon and Toolbar
' **     Options section.
' ** 16. In the Ribbon Name drop-down list, select HideData. Click OK to dismiss the dialog box.
' ** 17. Close and re-open the database.
' **     The Edit group is no longer displayed, and the Fluent UI includes the A Custom Tab tab,
' **     which contains the A Custom Group group with the Built-in Paste button.
' ** 18. To clean up, repeat the previous few steps to display the Access Options dialog box.
' **     Delete the contents of the Ribbon Name option, so that Access displays its default
' **     Fluent UI after you close and re-open the database.
' **       Note: You can also use a Ribbon from the USysRibbons table to supply the UI
' **       for a specific form or report. To do this, open the form or report in Design
' **       or Layout mode, and set the form's RibbonName property to the name of the
' **       Ribbon you want to use. You must select the form itself, rather than any
' **       control or section on the form, before you can set this property.
' **
' ** Trust Accountant custom Access 2007 Ribbons for USysRibbons.
' ** ============================================================
' **
' **   ID      RibbonName      RibbonXML
' **   ======= =============== ==========
' **   1       TAReports       <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
' **                             <ribbon startFromScratch="false">
' **                               <tabs>
' **                                 <tab idMso="TabPrintPreviewAccess" visible="false" />
' **                                 <tab idMso="TabAddIns" visible="false" />
' **                                 <tab id="dbCustomTab" label="Preview Commands" visible="true">
' **                                   <group id="dbCustomGroup" label="Print Preview">
' **                                     <control idMso="PrintDialogAccess" label="Print..." enabled="true"/>
' **                                     <control idMso="PrintPreviewZoomMenu" label="Zoom" enabled="true"/>
' **                                     <control idMso="PrintPreviewClose" label="Close" enabled="true"/>
' **                                   </group>
' **                                 </tab>
' **                               </tabs>
' **                             </ribbon>
' **                           </customUI>
' **   2       TAReportsCA     <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
' **                             <ribbon startFromScratch="false">
' **                               <tabs>
' **                                 <tab idMso="TabPrintPreviewAccess" visible="false" />
' **                                 <tab idMso="TabAddIns" visible="false" />
' **                                 <tab id="dbCustomTab" label="Preview Commands" visible="true">
' **                                   <group id="dbCustomGroup" label="Print Preview">
' **                                     <button id="TARptsCAPrint" label="Print..." onAction="mcrTAReports_CAPrint" imageMso="PrintDialogAccess"/>
' **                                     <control idMso="PrintPreviewZoomMenu" label="Zoom" enabled="true"/>
' **                                     <control idMso="PrintPreviewClose" label="Close" enabled="true"/>
' **                                   </group>
' **                                 </tab>
' **                               </tabs>
' **                             </ribbon>
' **                           </customUI>
' **
' ** Loading Customizations at Run Time
' ** ==================================
' **
' ** If you want to load static customizations at run time, you can store those customizations
' ** in the USysRibbons table, and set a form or report's RibbonName property as necessary.
' ** But, if you need to create dynamic customizations, call the Application.LoadCustomUI method.
' ** The following example creates a Ribbon customization that displays a button for each form
' ** in the application, and handles the onAction callback for each button to load the requested form.
' ** ...
' **
' ** http://msdn.microsoft.com/en-us/library/aa338202.aspx

' ** Set Startup Properties from Visual Basic:
' ** =========================================
' ** The names of the startup properties differ from the text that appears in the Startup dialog box.
' ** The following table provides the name of each startup property as it's used in Visual Basic code.
' **
' **   Text In Startup Dialog Box      Property Name
' **   ==============================  ========================
' **   Application Title               AppTitle
' **   Application Icon                AppIcon
' **   Display Form/Page               StartupForm
' **   Display Database Window         StartupShowDBWindow
' **   Display Status Bar              StartupShowStatusBar
' **   Menu Bar                        StartupMenuBar
' **   Shortcut Menu Bar               StartupShortcutMenuBar
' **   Allow Full Menus                AllowFullMenus
' **   Allow Default Shortcut Menus    AllowShortcutMenus
' **   Allow Built-In Toolbars         AllowBuiltInToolbars
' **   Allow Toolbar/Menu Changes      AllowToolbarChanges
' **   Allow Viewing Code After Error  AllowBreakIntoCode
' **   Use Access Special Keys         AllowSpecialKeys
' **
' ** http://msdn.microsoft.com/en-us/library/bb256564.aspx

' ** Set Options from Visual Basic:
' ** ==============================
' ** The value that you pass to the SetOption method as the setting argument depends on which type
' ** of option you are setting. The following table establishes some guidelines for setting options.
' **
' **   If The Option Is                            Then The Setting Argument Is
' **   ==========================================  ======================================================
' **   A text box                                  A string
' **   A check box                                 A Boolean value — True (–1) or False (0)
' **   An option button in an option group, or     An integer corresponding to the option's position
' **     an option in a combo box or a list box    in the option group or list (starting with zero [0])
' **
' ** The following tables list the names of all options that can be set or returned from code
' ** and the tabs on which they can be found in the Access Options dialog box, followed by the
' ** corresponding string argument that you must pass to the SetOption or GetOption method.
' **
' **   Popular Tab:
' **     Creating Databases Section:
' **       Option Text                String Argument
' **       =========================  ============================
' **       New database sort order    New Database Sort Order
' **       Default database folder    Default Database Directory
' **       Default file format        Default File Format
' **
' **   Current Database Tab:
' **     Application Options Section:
' **       Option Text                                                 String Argument
' **       ==========================================================  =================================
' **       Compact on Close                                            Auto Compact
' **       Remove personal information from file properties on save    Remove Personal Information
' **       Use Windows-themed Controls on Forms                        Themed Form Controls
' **       Enable Layout View for this database                        DesignWithData
' **       Check for truncated number fields                           CheckTruncatedNumFields
' **       Picture Property Storage Format                             Picture Property Storage Format
' **
' **     Name AutoCorrect Options Section:
' **       Option Text                     String Argument
' **       ==============================  ==============================
' **       Track name AutoCorrect info     Track Name AutoCorrect Info
' **       Perform name AutoCorrect        Perform Name AutoCorrect
' **       Log name AutoCorrect changes    Log Name AutoCorrect Changes
' **
' **     Filter Lookup Options For <Database Name> Database Section:
' **       Option Text                                                        String Argument
' **       =================================================================  ============================
' **       Show list of values in, Local indexed fields                       Show Values in Indexed
' **       Show list of values in, Local nonindexed fields                    Show Values in Non-Indexed
' **       Show list of values in, ODBC fields                                Show Values in Remote
' **       Show list of values in, Records in local snapshot                  Show Values in Snapshot
' **       Show list of values in, Records at server                          Show Values in Server
' **       Don't display lists where more than this number of records read    Show Values Limit
' **
'Navigation Section:

'Ribbon and Toolbar Options Section:
'  Ribbon Name:
'  Shortcut Menu Bar:
'  Allow Full Menus:
'  Allow Default Shortcut Menus:

' **   Datasheet Tab:
' **     Default Colors Section:
' **       Option Text                   String Argument
' **       ============================  ==========================
' **       Font color                    Default Font Color
' **       Background color              Default Background Color
' **       Alternate background color    _64
' **       Gridlines color               Default Gridlines Color
' **
' **     Gridlines And Cell Effects Section:
' **       Option Text                              String Argument
' **       =======================================  ==============================
' **       Default gridlines showing, Horizontal    Default Gridlines Horizontal
' **       Default gridlines showing, Vertical      Default Gridlines Vertical
' **       Default cell effect                      Default Cell Effect
' **       Default column width                     Default Column Width
' **
' **     Default Font Section:
' **       Option Text    String Argument
' **       =============  ========================
' **       Font           Default Font Name
' **       Size           Default Font Size
' **       Weight         Default Font Weight
' **       Underline      Default Font Underline
' **       Italic         Default Font Italic
' **
' **   Object Designers Tab:
' **     Table Design Section:
' **       Option Text                            String Argument
' **       =====================================  ======================================
' **       Default text field size                Default Text Field Size
' **       Default number field size              Default Number Field Size
' **       Default field type                     Default Field Type
' **       AutoIndex on Import/Create             AutoIndex on Import/Create
' **       Show Property Update Option Buttons    Show Property Update Options Buttons
' **
' **     Query Design Section:
' **       Option Text                                String Argument
' **       =========================================  ============================
' **       Show table names                           Show Table Names
' **       Output all fields                          Output All Fields
' **       Enable AutoJoin                            Enable AutoJoin
' **       SQL Server Compatible Syntax (ANSI 92),    ANSI Query Mode
' **         This database
' **       SQL Server Compatible Syntax (ANSI 92),    ANSI Query Mode Default
' **         Default for new databases
' **       Query design font, Font                    Query Design Font Name
' **       Query design font, Size                    Query Design Font Size
' **
' **     Forms/Reports Section:
' **       Option Text                    String Argument
' **       =============================  ============================
' **       Selection behavior             Selection Behavior
' **       Form template                  Form Template
' **       Report template                Report Template
' **       Always use event procedures    Always Use Event Procedures
' **
' **     Error Checking Section:
' **       Option Text                                 String Argument
' **       ==========================================  ============================
' **       Enable error checking                       Enable Error Checking
' **       Error indicator color                       Error Checking Indicator Color
' **       Check for unassociated label and control    Unassociated Label and Control Error Checking
' **       Check for new unassociated labels           New Unassociated Labels Error Checking
' **       Check for keyboard shortcut errors          Keyboard Shortcut Errors Error Checking
' **       Check for invalid control properties        Invalid Control Properties Error Checking
' **       Check for common report errors              Common Report Errors Error Checking
' **
' **   Proofing Tab:
' **     When Correcting Spelling In Microsoft Office Programs Section:
' **       Option Text                           String Argument
' **       ====================================  ============================
' **       Ignore words in UPPERCASE             Spelling ignore words in UPPERCASE
' **       Ignore words that contain numbers     Spelling ignore words with number
' **       Ignore Internet and file addresses    Spelling ignore Internet and file addresses
' **       Suggest from main dictionary only     Spelling suggest from main dictionary only
' **       Dictionary Language                   Spelling dictionary language
' **
' **   Advanced Tab:
' **     Editing Section:
' **       Option Text                         String Argument
' **       ==================================  ============================
' **       Move after enter                    Move After Enter
' **       Behavior entering field             Behavior Entering Field
' **       Arrow key behavior                  Arrow Key Behavior
' **       Cursor stops at first/last field    Cursor Stops at First/Last Field
' **       Default find/replace behavior       Default Find/Replace Behavior
' **       Confirm, Record changes             Confirm Record Changes
' **       Confirm, Document deletions         Confirm Document Deletions
' **       Confirm, Action queries             Confirm Action Queries
' **       Default direction                   Default Direction
' **       General alignment                   General Alignment
' **       Cursor movement                     Cursor Movement
' **       Datasheet IME control               Datasheet Ime Control
' **       Use Hijri Calendar                  Use Hijri Calendar
' **
' **     Display Section:
' **       Option Text                                String Argument
' **       =========================================  ============================
' **       Show this number of Recent Documents       Size of MRU File List
' **       Status bar                                 Show Status Bar
' **       Show animations                            Show Animations
' **       Show Smart Tags on Datasheets              Show Smart Tags on Datasheets
' **       Show Smart Tags on Forms and Reports       Show Smart Tags on Forms and Reports
' **       Show in Macro Design, Names column         Show Macro Names Column
' **       Show in Macro Design, Conditions column    Show Conditions Column
' **
' **     Printing Section:
' **       Option Text      String Argument
' **       ===============  =================
' **       Left margin      Left Margin
' **       Right margin     Right Margin
' **       Top margin       Top Margin
' **       Bottom margin    Bottom Margin
' **
' **     General Section:
' **       Option Text                                           String Argument
' **       ====================================================  ==========================================
' **       Provide feedback with sound                           Provide Feedback with Sound
' **       Use four-year digit year formatting, This database    Four-Digit Year Formatting
' **       Use four-year digit year formatting, All databases    Four-Digit Year Formatting All Databases
' **
' **     Advanced Section:
' **       Option Text                                     String Argument
' **       ==============================================  ============================================
' **       Open last used database when Access starts      Open Last Used Database When Access Starts
' **       Default open mode Default                       Open Mode for Databases
' **       Default record locking                          Default Record Locking
' **       Open databases by using record-level locking    Use Row Level Locking
' **       OLE/DDE timeout (sec)                           OLE/DDE Timeout (sec)
' **       Refresh interval (sec)                          Refresh Interval (sec)
' **       Number of update retries                        Number of Update Retries
' **       ODBC refresh interval (sec)                     ODBC Refresh Interval (sec)
' **       Update retry interval (msec)                    Update Retry Interval (msec)
' **       DDE operations, Ignore DDE requests             Ignore DDE Requests
' **       DDE operations, Enable DDE refresh              Enable DDE Refresh
' **       Command-line arguments                          Command-Line Arguments
' **
' ** http://msdn.microsoft.com/en-us/library/bb256546.aspx

' ** Set Properties of Data Access Objects in Visual Basic:
' ** ======================================================
' ** Keep in mind that when you create the property, you must correctly specify
' ** its Type property before you append it to the Properties collection.
' ** You can determine the Type property based on the information in the Settings
' ** section of the Help topic for the individual property. The following table
' ** provides some guidelines for determining the setting of the Type property.
' **
' **   If The Property Setting Is  Then The Type Property Setting Should Be
' **   ==========================  ========================================
' **   A string                    dbText
' **   True/False                  dbBoolean
' **   An integer                  dbInteger
' **
' ** The following table lists some Microsoft Access–defined properties that apply to DAO objects.
' **
' **   DAO object     Microsoft Access–defined properties
' **   =============  ================================================================================================
' **   Database       AppTitle, AppIcon, StartUpShowDBWindow, StartUpShowStatusBar, AllowShortcutMenus,
' **                  AllowFullMenus, AllowBuiltInToolbars, AllowToolbarChanges, AllowBreakIntoCode,
' **                  AllowSpecialKeys, Replicable, ReplicationConflictFunction
' **   SummaryInfo    Title, Subject, Author, Manager, Company, Category, Keywords, Comments, Hyperlink Base
' **     Container    (See the Summary tab of the DatabaseName Properties dialog box, available by clicking
' **                  Database Properties on the File menu.)
' **   UserDefined    (See the Summary tab of the DatabaseName Properties dialog box, available by clicking
' **     Container     Database Properties on the File menu.)
' **   TableDef       DatasheetBackColor, DatasheetCellsEffect, DatasheetFontHeight, DatasheetFontItalic,
' **                  DatasheetFontName, DatasheetFontUnderline, DatasheetFontWeight, DatasheetForeColor,
' **                  DatasheetGridlinesBehavior, DatasheetGridlinesColor, description, FrozenColumns,
' **                  RowHeight, ShowGrid
' **   QueryDef       DatasheetBackColor, DatasheetCellsEffect, DatasheetFontHeight, DatasheetFontItalic,
' **                  DatasheetFontName, DatasheetFontUnderline, DatasheetFontWeight, DatasheetForeColor,
' **                  DatasheetGridlinesBehavior, DatasheetGridlinesColor, description, FailOnError,
' **                  FrozenColumns, LogMessages, MaxRecords, RecordLocks, RowHeight, ShowGrid, UseTransaction
' **   Field          Caption, ColumnHidden, ColumnOrder, ColumnWidth, DecimalPlaces, description, Format, InputMask
' **
' ** http://msdn.microsoft.com/en-us/library/bb256552.aspx

' ** How to: Use Existing Custom Menus and Toolbars:
' ** ===============================================
' ** This topic explains how custom toolbars and menu bars that you created in
' ** earlier versions of Access behave when you open those older databases in
' ** Microsoft Office Access 2007. This topic also explains how to turn off
' ** the Ribbon so that you can use just your customtoolbars and menu bars.
' **
' ** http://msdn.microsoft.com/en-us/library/bb258174.aspx
' ** Strictly menu-driven suggestions; no VBA references.

' ****************************************************************************
' ** Hotfixes:
' ****************************************************************************

' ** ***************************************************************
' ** THIS ONE'S OURS!
' ** ***************************************************************
' ** KB958378
' ** You cannot open a compiled database (.mde) file or run a .Net
' ** Framework application on a Windows Vista-based computer.
' **
' **
' ** Issues That The Hotfix Package Fixes
' **
' ** When you open a compiled database (.mde) file, or you run a .Net Framework
' ** application on a Windows Vista-based computer, the operation may fail.
' **
' ** This problem occurs for one of the following reasons:
' **
' ** 1. The .mde file was created on a Windows XP-based computer by using one of
' **    the following applications, and the .mde file has a reference to ADOX 2.8:
' **      Microsoft Office Access 2007
' **      Access 2003
' **      Access 2002
' **      Access 2000
' **
' ** 2. The .Net Framework application is created on a Windows XP-based computer,
' **    and the .Net Framework application uses ADOX 2.8 as a referenced library.
' **
' ** http://support.microsoft.com/kb/958378
' **
' ** File:     361729_intl_i386_zip.exe
' ** Password: WNQro4I3
' **
' ** ***************************************************************

' ** ***************************************************************
' ** KB943249
' ** Description of the Access 2007 hotfix package: January 28, 2008.
' **
' ** Issues That The Hotfix Package Fixes
' **
' ** This hotfix fixes the following issues that were not
' ** previously documented in a Microsoft Knowledge Base article:
' **
' ** 1. You open a database in Access 2007 and then enable the Overlapping
' **    Windows setting under Access Options in Current Database.
' **
' **    In the database, you have a tabbed form that you bring to the front.
' **    You switch between the tabbed sections of the form by using the arrow keys.
' **
' **    When you switch from one tab to another, Access 2007 may take longer than
' **    expected to complete the switch. Access 2007 may take from five to 50 seconds
' **    or longer to complete the switch, depending on the speed of the computer.
' **
' ** 2. You use a macro in Access 2007 that was originally used in Microsoft Office
' **    Access 2003 or earlier versions. This macro uses either the SendObject function
' **    to send an object (table, query, form or report) to an e-mail recipient or the
' **    OutputTo function to save an object in the .xls format that is used by the
' **    following versions of Microsoft Excel:
' **        Microsoft Excel 97
' **        Microsoft Excel 2000
' **        Microsoft Excel 2002
' **        Microsoft Excel 2003
' **
' **    When you run this macro in Access 2007, you receive the following error message:
' **      The format in which you are attempting to output the current object is not available.
' **      Either you are attempting to output the current object to a format that is not valid
' **      for its object type, or the formats that enable you to output data as a Microsoft Excel,
' **      rich-text format, MS-DOS text, or HTML file are missing from the Windows Registry.
' **      Run Setup to reinstall Microsoft Office Access or, if you’re familiar with the settings
' **      in the Registry, try to correct them yourself. For more information on the Registry, click Help.
' **
' ** 3. You use the release version of Access 2007, and you try to open a database (.accde or .mde)
' **    or a project (.ade) that was compiled in Access 2007 Service Pack 1 (SP1).
' **    Then, you receive the following error message:
' **      This database is in an unrecognized format. The database may have been created
' **      with a later version of Microsoft Office Access than the one you are using.
' **      Upgrade your version of Microsoft Office Access to the current one, then open this database.
' **
' ** http://support.microsoft.com/kb/943249
' **
' **  File:     339458_intl_i386_zip.exe
' **  Password: n*jTGkad
' **
' ** ***************************************************************

' ** ***************************************************************
' ** KB949404
' ** Description of the Access 2007 hotfix package: March 6, 2008
' **
' ** Issues That This Hotfix Package Fixes
' **
' ** This hotfix fixes the following issues that are not
' ** previously documented in a Microsoft Knowledge Base article:
' **
' ** 1. When you try to compact and to repair a database in a shared folder, a new database is created.
' **    The new database is named as the default name for that version of Access. For Access 2003,
' **    the name is "Db1.mdb." For Access 2007, the name is "Database1.mdb." However, the old database
' **    is not compacted and not deleted, but the new database is compacted. The same problem occurs
' **    if the Compact on Close database option is selected.
' **
' ** 2. Access 2007 may take a long time to respond when you modify a query from within Design view.
' **    This behavior occurs if the tables that the query is based on are from an Open Database
' **    Connectivity (ODBC) database data source.
' **
' ** http://support.microsoft.com/kb/949404
' **
' **  File:     341536_intl_i386_zip.exe
' **  Password: UI#AH_$%up
' **
' ** ***************************************************************

' ** ***************************************************************
' ** KB960307
' ** Description of the Access 2007 hotfix package (Access.msp): December 16, 2008
' **
' ** Issues That This Hotfix Package Fixes
' **
' ** 1. You use source control to control your Access 2007 database. You change the code module in
' **      Visual Basic Editor, and then you close the Visual Basic Editor without saving changes.
' **      When you check the object back in and use the Get Last Version feature, the code changes
' **      are not there.
' **
' ** 2. After you set the main form's recordset in Access 2007, the subform becomes blank.
' **      You may also receive the following error message:
' **        Run-time error 2467 - The expression you entered does not exist.
' **
' ** 3. You experience slower performance than you did in earlier versions of Access when you share
' **      a database file over a network.
' **
' ** 4. The file size of an Access 2007 file spontaneously increases in increments of 4 KB after you
' **      open and close a form. The file growth can be seen even when the form is closed.
' **
' ** http://support.microsoft.com/kb/960307
' **
' **  File:     367941_intl_i386_zip.exe
' **  Password: H%P+FZW93Y
' **
' ** ***************************************************************

' ** Startup Options:
' **   All items on the Tools->Startup window.
' **   Items stored in CurrendDb.Properties collection.
' **   Note: Properties do not exist until they have been used at least once.
' **     The easiest way to set them is to use the Tools->Startup window.
' **     If VBA is used, the property must first be created using the
' **     Add method of the Application.AccessObjectProperties collection.
' **   Example:
' **     Get: strTmp = CurrentDb.Properties("AppTitle")
' **            strTmp:  "Trust Accountant™"
' **     Set: CurrentDb.Properties("StartupForm") = "frmMenu_Title"
' **
' **   Option Name               Tools->Startup Entry
' **   ========================  ================================
' **   AppTitle                  Application Title
' **   AppIcon                   Application Icon
' **   StartupForm               Display Form/Page
' **   StartupShowDBWindow       Display Database Window
' **   StartupShowStatusBar      Display Status Bar
' **   StartupMenuBar            Menu Bar
' **   StartupShortcutMenuBar    Shortcut Menu Bar
' **   AllowFullMenus            Allow Full Menus
' **   AllowShortcutMenus        Allow Default Shortcut Menus
' **   AllowBuiltInToolbars      Allow Built-In Toolbars
' **   AllowToolbarChanges       Allow Toolbar/Menu Changes
' **   AllowBreakIntoCode        Allow Viewing Code After Error
' **   AllowSpecialKeys          Use Access Special Keys

' ** Access Options:
' **   All items on the Tools->Options window tabs.
' **   Items stored in registry under:
' **     HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Access\Settings\
' **   Use GetOption() and SetOption() internal functions.
' **   See Help file for full list of options.
' **   Example:
' **     Get: blnTmp = Application.GetOption("Confirm Action Queries")
' **            blnTmp = False
' **     Set: Application.SetOption "Default Record Locking", 2
' **
' **   Notes:
' **     Option String                               Tools->Options Entry
' **     ==========================================  ===============================================
' **                   General Tab:
' **     Perform Name AutoCorrect                    Name AutoCorrect: Perform name AutoCorrect
' **                                                   This refers to names of Access objects
' **                                                   only, not spelling or data entry (below).
' **     Four-Digit Year Formatting                  Use four-digit year formatting: This database
' **     Four-Digit Year Formatting All Databases    Use four-digit year formatting: All databases
' **                   Keyboard Tab:
' **     Arrow Key Behavior                          Next Field: 0, Next Character: 1
' **

' ** AutoCorrect Options:
' **   All items on the Tools->AutoCorrect window.
' **   This refers to spelling and data entry corrections.
' **   Items stored in registry under:
' **     HKEY_CURRENT_USER\Software\Microsoft\Office\9.0\Common\AutoCorrect\
' **   Use AutoCorrect_Get() and AutoCorrect_Set() functions in modSecurityFunctions.
' **   Example:
' **     AutoCorrect_Get(optionname)               string
' **     AutoCorrect_Set(optionname, optionvalue)  string, boolean
' **
' **   Option Name (Key Name)       Tools->AutoCorrect Entry
' **   ===========================  =========================================
' **   CorrectTwoInitialCapitals    Correct two initial capitals
' **   CapitalizeSentence           Capitalize first letter of sentence
' **   CapitalizeNamesOfDays        Capitalize names of days
' **   ToggleCapsLock               Correct accidental use of CAPS LOCK key
' **   ReplaceText                  Replace text as you type
' *****************************************************************************

'Private intWindowVisible As Integer, lngWindowParent As Long
'Private Const W_ALL As Integer = 0  ' ** All, both visible and hidden.
'Private Const W_VIS As Integer = 1  ' ** Just visible.
'Private Const W_HID As Integer = 2  ' ** Just hidden.

Private Const MAX_LEN As Integer = 256
' **

Public Function IsAccess2007() As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "IsAccess2007"

        Dim blnRetVal As Boolean

110     blnRetVal = False
120     gblnIsAccess2007 = False

130     If Val(Application.SysCmd(acSysCmdAccessVer)) = 12 Then
          ' ** Also: Application.Version = 12.0 (Numeric).
140       blnRetVal = True
150       gblnIsAccess2007 = True
160     End If

EXITP:
170     IsAccess2007 = blnRetVal
180     Exit Function

ERRH:
190     blnRetVal = False
200     gblnIsAccess2007 = False
210     Select Case ERR.Number
        Case Else
220       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
230     End Select
240     Resume EXITP

End Function

Public Function IsAccess2010() As Boolean

300   On Error GoTo ERRH

        Const THIS_PROC As String = "IsAccess2010"

        Dim blnRetVal As Boolean

310     blnRetVal = False
320     gblnIsAccess2010 = False

330     If Val(Application.SysCmd(acSysCmdAccessVer)) >= 14 Then
          ' ** ? Application.SysCmd(acSysCmdAccessVer)
          ' ** 14.0
          ' ** ? Application.Version
          ' ** 14.0
340       blnRetVal = True
350       gblnIsAccess2010 = True
360     End If

EXITP:
370     IsAccess2010 = blnRetVal
380     Exit Function

ERRH:
390     blnRetVal = False
400     gblnIsAccess2010 = False
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Function

Public Function SetOption_Access2007(blnHide As Boolean, Optional varProc As Variant) As Boolean

500   On Error GoTo ERRH

        Const THIS_PROC As String = "SetOption_Access2007"

        Dim obj As Object
        Dim varRibbonState As Variant
        Dim intOpenVer As Integer
        Dim blnIsRuntime As Boolean, blnIsVis As Boolean
        Dim blnRetVal As Boolean

510     blnRetVal = True

520     If IsAccess2007 = True Then  ' ** Function: Above.
530       intOpenVer = 12
540     ElseIf IsAccess2010 = True Then  ' ** Function: Above.
550       intOpenVer = 14
560     Else
570       intOpenVer = 0
580     End If

590     If intOpenVer > 0 Then

          ' ** Detect whether this is running on a Runtime version, because a
          ' ** Runtime will never have a Navigation Pane or Database Window.
600       blnIsRuntime = Application.SysCmd(acSysCmdRuntime)

610       Select Case blnHide
          Case True

620         Application.Echo False

630         If blnIsRuntime = False Then
640           blnIsVis = SetOption_DatabaseWindowVisible  ' ** Function: Below.
650           If blnIsVis = True Then

                ' ** Because the NavigateTo Method is not found in Access 2000 and Access 2003,
                ' ** using a generic object type allows it to be compiled in Access 2000.
660             Set obj = Application
670             obj.DoCmd.NavigateTo "acNavigationCategoryObjectType"  ' ** Not documented properly in help.
680             DoCmd.RunCommand acCmdWindowHide

                ' *******************************************************************
                ' ** These commands aren't recognized in Access 2000. Or 2003?
                ' *******************************************************************
                ' ** Get rid of the Navigation Pane.
                'DoCmd.LockNavigationPane True       ' ** Lock the Navigation Pane.
                'DoCmd.SetDisplayedCategories False  ' ** Hide all categories.
                ' *******************************************************************

690           End If
700         End If

710         DoCmd.ShowToolbar "Ribbon", acToolbarNo  ' ** Turn off the Ribbons.

            ' *******************************************************************
            ' ** The command above works fine, and compiles in Access 2000.
            ' *******************************************************************

            ' *******************************************************************
            ' ** Check to see if our special report ribbons
            ' ** are in the hidden System table USysRibbon.
720         SetRibbon_Access2007  ' ** Function: Below.
            ' *******************************************************************

730         Application.Echo True
740         DoEvents

750       Case False

            ' *******************************************************************
            'DoCmd.LockNavigationPane True  ' ** Lock the Navigation Pane.
            'DoCmd.SetDisplayedCategories False  ' ** Hide all categories.
            ' *******************************************************************

760         If blnIsRuntime = False Then

770           DoCmd.ShowToolbar "Ribbon", acToolbarYes  ' ** Turn on the Ribbons.

780           Select Case intOpenVer
              Case 12
790             varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                  "Software\Microsoft\Office\12.0\Common\Toolbars\Access", _
                  "QuickAccessToolbarStyle")  ' ** Function: Below.
800           Case 14
810             varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                  "Software\Microsoft\Office\14.0\Common\Toolbars\Access", _
                  "QuickAccessToolbarStyle")  ' ** Function: Below.
820           End Select
830           If IsNumeric(varRibbonState) = True Then
840             If Val(varRibbonState) = acRibbonStateAutohide Then
                  ' ** Ribbon in auto hide state, show it.
850               SendKeys "^{F1}", False
860               DoEvents
870             End If
880           Else
890             If varRibbonState = RET_ERR Then
                  ' ** Not really much I can do!
900             End If
910           End If

920           DoCmd.SelectObject acForm, "", True  '"frmMenu_Title"
930           DoCmd.RunCommand acCmdWindowUnhide

940         End If

950       End Select

960     End If

EXITP:
970     Set obj = Nothing
980     SetOption_Access2007 = blnRetVal
990     Exit Function

ERRH:
1000    Application.Echo True
1010    blnRetVal = False
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Function

Public Function SetOption_Access2010(blnHide As Boolean) As Boolean

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "SetOption_Access2010"

        Dim blnRetVal As Boolean

        ' ** For now, these will be the same.
1110    blnRetVal = SetOption_Access2007(blnHide)  ' ** Function: Above.

EXITP:
1120    SetOption_Access2010 = blnRetVal
1130    Exit Function

ERRH:
1140    Application.Echo True
1150    blnRetVal = False
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Function

Public Function SetOption_DatabaseWindowVisible() As Boolean
' ** Returns True if the Access Navigation Pane or Database Window is visible, otherwise False.

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "SetOption_DatabaseWindowVisible"

        Dim lngHWindow As Long
        Dim intOpenVer As Integer
        Dim blnRetVal As Boolean

1210    blnRetVal = False

1220    If IsAccess2007 = True Then  ' ** Function: Above.
1230      intOpenVer = 12
1240    ElseIf IsAccess2010 = True Then  ' ** Function: Above.
1250      intOpenVer = 14
1260    Else
1270      intOpenVer = 0
1280    End If

1290    If intOpenVer > 0 Then
1300      blnRetVal = IsNavPaneOpen  ' ** Function: Below.
          'lngHWindow = FindWindowEx(Application.hWndAccessApp, 0, "NetUINativeHWNDHost", "Navigation Pane Host")  ' ** API Function: modWindowFunctions.
          'lngHWindow = FindWindowEx(lngHWindow, 0, "NetUIHWND", vbNullString)  ' ** API Function: modWindowFunctions.
1310    Else                                          ' ** Access 2000/2003 Database Window.
1320      lngHWindow = FindWindowEx(Application.hWndAccessApp, 0, "MDIClient", vbNullString)  ' ** API Function: modWindowFunctions.
1330      lngHWindow = FindWindowEx(lngHWindow, 0, "Odb", vbNullString)  ' ** API Function: modWindowFunctions.
1340      blnRetVal = (IsWindowVisible(lngHWindow) <> 0)  ' ** API Function: modWindowFunctions.
1350    End If

EXITP:
1360    SetOption_DatabaseWindowVisible = blnRetVal
1370    Exit Function

ERRH:
1380    blnRetVal = False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Function

Public Function IsNavPaneOpen() As Boolean
' ** For some reason, this result always comes up false during startup,
' ** even when it's clearly visible, while running this function from
' ** the Immediate Window always gives an accurate response.
' ** So, it's also called in the Form_Timer() event, which then works.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "IsNavPaneOpen"

        Dim blnRetVal As Boolean

1510    blnRetVal = Win_List_Open_Child(True)  ' ** Module Function: modWindowFunctions.

EXITP:
1520    IsNavPaneOpen = blnRetVal
1530    Exit Function

ERRH:
1540    blnRetVal = False
1550    Select Case ERR.Number
        Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Function

Public Function SetReport_Access2007(blnShow As Boolean) As Boolean
' ** QAT : QuickAccessToolbar

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "SetReport_Access2007"

        Dim varRibbonState As Variant
        Dim intOpenVer As Integer
        Dim blnRetVal As Boolean

1610    blnRetVal = True

1620    If IsAccess2007 = True Then  ' ** Function: Above.
1630      intOpenVer = 12
1640    ElseIf IsAccess2010 = True Then  ' ** Function: Above.
1650      intOpenVer = 14
1660    Else
1670      intOpenVer = 0
1680    End If

        ' ** acRibbonState enumeration:  (my own)
        ' **   0  acRibbonStateNormalAbove  QAT Normal, above ribbon.
        ' **   1  acRibbonStateNormalBelow  QAT Normal, below ribbon.
        ' **   4  acRibbonStateAutohide     QAT Autohide.

1690    If intOpenVer > 0 Then
1700      Select Case blnShow
          Case True
1710        DoCmd.ShowToolbar "Ribbon", acToolbarYes  ' ** Turn on the Ribbons.
1720        DoEvents
1730        Select Case intOpenVer
            Case 12
1740          varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                "Software\Microsoft\Office\12.0\Common\Toolbars\Access", _
                "QuickAccessToolbarStyle")  ' ** Function: Below.
1750        Case 14
1760          varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                "Software\Microsoft\Office\14.0\Common\Toolbars\Access", _
                "QuickAccessToolbarStyle")  ' ** Function: Below.
1770        End Select
1780        If IsNumeric(varRibbonState) = True Then
1790          If Val(varRibbonState) = acRibbonStateAutohide Then
                ' ** Ribbon in auto hide state, show it.
                'SendKeys "^{F1}", False
1800            DoEvents
1810          End If
1820        Else
1830          If varRibbonState = RET_ERR Then
                ' ** Not really much I can do!
1840          End If
1850        End If
1860      Case False
1870        Select Case intOpenVer
            Case 12
1880          varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                "Software\Microsoft\Office\12.0\Common\Toolbars\Access", _
                "QuickAccessToolbarStyle")  ' ** Function: Below.
1890        Case 14
1900          varRibbonState = ReturnRegKeyValue(HKEY_CURRENT_USER, _
                "Software\Microsoft\Office\14.0\Common\Toolbars\Access", _
                "QuickAccessToolbarStyle")  ' ** Function: Below.
1910        End Select
1920        If IsNumeric(varRibbonState) = True Then
1930          If Val(varRibbonState) = acRibbonStateNormalAbove Then
                ' ** Ribbon is in full view, hide it.
                'SendKeys "^{F1}", False
1940            DoEvents
1950          End If
1960        Else
1970          If varRibbonState = RET_ERR Then
                ' ** Not really much I can do!
1980          End If
1990        End If
2000        DoCmd.ShowToolbar "Ribbon", acToolbarNo  ' ** Turn off the Ribbons.
2010        DoEvents
2020      End Select
2030    End If

EXITP:
2040    SetReport_Access2007 = blnRetVal
2050    Exit Function

ERRH:
2060    blnRetVal = False
2070    Select Case ERR.Number
        Case 2585  ' ** This action can't be carried out while processing a form or report event.
          ' ** Ignore.
2080    Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Function

Public Function SetReport_Access2010(blnShow As Boolean) As Boolean

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "SetReport_Access2010"

        Dim blnRetVal As Boolean

        ' ** For now, these will be the same.
2210    blnRetVal = SetReport_Access2007(blnShow)  ' ** Function: Above.

EXITP:
2220    SetReport_Access2010 = blnRetVal
2230    Exit Function

ERRH:
2240    blnRetVal = False
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Function

Public Function SetRibbon_Access2007() As Boolean
' ** Check for the presence of our custom ribbons in the table USysRibbons.
' ** Of course, since the USysRibbons table comes with our MDE, it should always be there.
' ** But I don't trust different versions of Access to depend on it.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "SetRibbon_Access2007"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, tdf As DAO.TableDef, fld As DAO.Field
        Dim blnFound As Boolean, blnID_Fld As Boolean
        Dim blnAppend1 As Boolean, blnAppend2 As Boolean, blnEdit1 As Boolean, blnEdit2 As Boolean
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

        Const TRIBBON As String = "USysRibbons"
        Const RIBBON1 As String = "TAReports"
        Const RIBBON2 As String = "TAReportsCA"

2310    blnRetVal = True

2320    Set dbs = CurrentDb
2330    With dbs

2340      blnFound = False: blnID_Fld = False
2350      For Each tdf In .TableDefs
2360        With tdf
2370          If .Name = TRIBBON Then
2380            blnFound = True
2390            For Each fld In .Fields
2400              With fld
2410                If .Name = "ID" Then
2420                  blnID_Fld = True
2430                  Exit For
2440                End If
2450              End With
2460            Next
2470            Exit For
2480          End If
2490        End With
2500      Next

2510      If blnFound = False Then
            ' ** Data-Definition: Create table USysRibbons.
2520        Set qdf = .QueryDefs("zz_qry_System_19_01")
2530        qdf.Execute
            ' ** Data-Definition: Create index [ID] PrimaryKey on table USysRibbons.
2540        Set qdf = .QueryDefs("zz_qry_System_19_02")
2550        qdf.Execute
            ' ** Data-Definition: Create index [RibbonName] Unique on table USysRibbons.
2560        Set qdf = .QueryDefs("zz_qry_System_19_03")
2570        qdf.Execute
2580        blnID_Fld = True
2590      End If
2600      .TableDefs.Refresh
2610      .TableDefs.Refresh

2620      Set rst = .OpenRecordset(TRIBBON)
2630      With rst
2640        blnAppend1 = True: blnAppend2 = True: blnEdit1 = False: blnEdit2 = False
2650        If .BOF = True And .EOF = True Then
              ' ** Add both of our ribbons.
2660        Else
2670          .MoveLast
2680          lngRecs = .RecordCount
2690          .MoveFirst
2700          For lngX = 1& To lngRecs
2710            Select Case ![RibbonName]
                Case RIBBON1
2720              blnAppend1 = False
2730              If IsNull(![RibbonXML]) = False Then
2740                If Trim(![RibbonXML]) <> vbNullString Then
                      ' ** At this point I'm going to assume it's all there.
2750                Else
2760                  blnEdit1 = True
2770                End If
2780              Else
2790                blnEdit1 = True
2800              End If
2810            Case RIBBON2
2820              blnAppend2 = False
2830              If IsNull(![RibbonXML]) = False Then
2840                If Trim(![RibbonXML]) <> vbNullString Then
                      ' ** At this point I'm going to assume it's all there.
2850                Else
2860                  blnEdit1 = True
2870                End If
2880              Else
2890                blnEdit1 = True
2900              End If
2910            End Select
2920            If lngX < lngRecs Then .MoveNext
2930          Next
2940        End If
2950        .Close
2960      End With  ' ** rst.

2970      If blnAppend1 = True Then
2980        Select Case blnID_Fld
            Case True
              ' ** Append tblTemplate_USysRibbons, with ID, to USysRibbons, by specified [nam].
2990          Set qdf = .QueryDefs("zz_qry_System_70a")
3000        Case False
              ' ** Append tblTemplate_USysRibbons, without ID, to USysRibbons, by specified [nam].
3010          Set qdf = .QueryDefs("zz_qry_System_70b")
3020        End Select
3030        With qdf.Parameters
3040          ![nam] = RIBBON1
3050        End With
3060        qdf.Execute
3070      End If

3080      If blnAppend2 = True Then
3090        Select Case blnID_Fld
            Case True
              ' ** Append tblTemplate_USysRibbons, with ID, to USysRibbons, by specified [nam].
3100          Set qdf = .QueryDefs("zz_qry_System_70a")
3110        Case False
              ' ** Append tblTemplate_USysRibbons, without ID, to USysRibbons, by specified [nam].
3120          Set qdf = .QueryDefs("zz_qry_System_70b")
3130        End Select
3140        With qdf.Parameters
3150          ![nam] = RIBBON2
3160        End With
3170        qdf.Execute
3180      End If

3190      If blnEdit1 = True Then
            ' ** Update zz_qry_System_71 (USysRibbons, with RibbonXML_new, via DLookups() to tblTemplate_USysRibbons, by specified [nam]).
3200        Set qdf = .QueryDefs("zz_qry_System_72")
3210        With qdf.Parameters
3220          ![nam] = RIBBON1
3230        End With
3240        qdf.Execute dbFailOnError
3250      End If

3260      If blnEdit2 = True Then
            ' ** Update zz_qry_System_71 (USysRibbons, with RibbonXML_new, via DLookups() to tblTemplate_USysRibbons, by specified [nam]).
3270        Set qdf = .QueryDefs("zz_qry_System_72")
3280        With qdf.Parameters
3290          ![nam] = RIBBON2
3300        End With
3310        qdf.Execute dbFailOnError
3320      End If

3330    End With  ' ** dbs.

EXITP:
3340    Set fld = Nothing
3350    Set tdf = Nothing
3360    Set rst = Nothing
3370    Set qdf = Nothing
3380    Set dbs = Nothing
3390    SetRibbon_Access2007 = blnRetVal
3400    Exit Function

ERRH:
3410    Application.Echo True
3420    blnRetVal = False
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Function

Public Function SetRibbon_Access2010() As Boolean

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "SetRibbon_Access2010"

        Dim blnRetVal As Boolean

        ' ** For now, these will be the same.
3510    blnRetVal = SetRibbon_Access2010  ' ** Function: Above.

EXITP:
3520    SetRibbon_Access2010 = blnRetVal
3530    Exit Function

ERRH:
3540    Application.Echo True
3550    blnRetVal = False
3560    Select Case ERR.Number
        Case Else
3570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3580    End Select
3590    Resume EXITP

End Function

Public Sub SetNav_Access2007(frm As Access.Form)
' ** CLR_AC07 = 13603685 : Medium blue, standard Access 2007 border color (101 147 207).

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "SetNav_Access2007"

        Dim ctl As Access.Control
        Dim intOpenVer As Integer

3610    If IsAccess2007 = True Then  ' ** Function: Above.
3620      intOpenVer = 12
3630    ElseIf IsAccess2010 = True Then  ' ** Function: Above.
3640      intOpenVer = 14
3650    Else
3660      intOpenVer = 0
3670    End If

3680    If intOpenVer > 0 Then
3690      With frm
3700        For Each ctl In .Detail.Controls
3710          With ctl
3720            If Left(.Name, 4) = "Nav_" Then
3730              If .ControlType = acLine Then
                    ' ** Turn off Access 2000 lines.
3740                If .Name = "Nav_hline03" Then
                      ' ** Turn on Access 2007 line.
                      ' ** Blue color: 101, 147, 207.
3750                  If intOpenVer = 14 Then
3760                    .BorderColor = 12038060  ' ** 172, 175, 183.
3770                  End If
3780                  .Visible = True
3790                ElseIf Mid(.Name, 5, 5) = "hline" Or Mid(.Name, 5, 5) = "vline" Then
                      ' ** Nav_hline01, Nav_hline02
                      ' ** Nav_vline01, Nav_vline02, Nav_vline03, Nav_vline04
3800                  .Visible = False
3810                End If
3820              End If
3830            End If
3840          End With
3850        Next
3860      End With
3870    End If

EXITP:
3880    Set ctl = Nothing
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub

Public Sub SetNav_Access2010(frm As Access.Form)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "SetNav_Access2010"

        ' ** For now, these will be the same.
4010    SetNav_Access2007 frm  ' ** Function: Above

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Function AC07_Ribbon_Check() As Boolean

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "AC07_Ribbon_Check"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, idx As DAO.index, rst As DAO.Recordset
        Dim strRibbon_XML As String
        Dim lngRecs As Long
        Dim blnRetVal As Boolean

4110    blnRetVal = True

4120    Set dbs = CurrentDb
4130    With dbs

          ' ** USysRibbons table field definitions:
          ' **   Field Name    Data Type    Description
          ' **   ============  ============ ========================================================================================
          ' **   ID            AutoNumber   Primary Key field.
          ' **   RibbonName    Text         Contains the name of the custom ribbon to be associated with this customization.
          ' **   RibbonXml     Memo         Contains the Ribbon Extensibility XML (RibbonX) that defines the Ribbon customization.

4140      If TableExists("USysRibbons") = False Then

4150        Set tdf = .CreateTableDef("USysRibbons")
4160        With tdf
4170          .Attributes = dbSystemObject ' ** + dbHiddenObject  ' ** If I set Hidden here, I don't seem to be able to ever
4180          Set fld = .CreateField("ID")                        ' ** see it, even with both Hidden and System showing.
4190          With fld                                            ' ** Without setting it, it automatically gets set Hidden,
4200            .Type = dbLong                                    ' ** but is visible when System showing.
4210            .Required = True                                  ' ** In both cases, however, a query can access it.
4220            .Attributes = dbAutoIncrField
4230          End With
4240          .Fields.Append fld
4250          Set fld = .CreateField("RibbonName")
4260          With fld
4270            .Type = dbText
4280            .Size = 255
4290            .Required = True
4300            .AllowZeroLength = False
4310          End With
4320          .Fields.Append fld
4330          Set fld = .CreateField("RibbonXml")
4340          With fld
4350            .Type = dbMemo
4360            .AllowZeroLength = False
4370          End With
4380          .Fields.Append fld
4390          .Fields.Refresh
4400          Set idx = .CreateIndex("PrimaryKey")
4410          With idx
4420            Set fld = .CreateField("ID", dbLong)
4430            .Fields.Append fld
4440            .Unique = True
4450            .Primary = True
4460          End With
4470          .Indexes.Append idx
4480          Set idx = .CreateIndex("RibbonName")
4490          With idx
4500            Set fld = .CreateField("RibbonName", dbText)
4510            .Fields.Append fld
4520            .Unique = True
4530          End With
4540          .Indexes.Append idx
              '.CreateProperty "Description", dbText, "Vista only: Access Application-Level Custom Ribbon.", True  ' ** DOESN'T TAKE!
4550        End With
4560        .TableDefs.Append tdf
4570        .TableDefs.Refresh

4580      End If

          ' ** USysRibbons table data:
          ' **   Column Name  Value
          ' **   ===========  =========================================================================
          ' **   RibbonName   HideTheRibbon
          ' **   RibbonXML    <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
          ' **                  <ribbon startFromScratch="true">
          ' **                  </ribbon>
          ' **                </customUI>

          ' ** The Microsoft URL refers to the schema standard, in the same way that most HTML
          ' ** documents have a reference to some international standards organization. It does
          ' ** not mean we're connecting to Microsoft, nor that an internet connection is required.
          ' ** Also, making this all on one line is for simplicity. A useful, complex Ribbon
          ' ** definition would have tags on different lines, with appropriate indents.
4590      strRibbon_XML = "<customUI xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34) & ">" & _
            "<ribbon startFromScratch=" & Chr(34) & "True" & Chr(34) & "></ribbon></customUI>"

          ' ** Note: Since this is our own frontend table, there should
          ' ** never be anything else in it but what we put there.
          ' ** Nonetheless, I'm going to check anyway, since I really
          ' ** don't know what might happen in Access 2007 on a Vista
          ' ** machine over the long run.
4600      Set rst = .OpenRecordset("USysRibbons", dbOpenDynaset)
4610      With rst
4620        If .BOF = True And .EOF = True Then
4630          .AddNew
4640          ![RibbonName] = "HideTheRibbon"
4650          ![RibbonXML] = strRibbon_XML
4660          .Update
4670        Else
4680          .MoveLast
4690          lngRecs = .RecordCount
4700          .MoveFirst
4710          If lngRecs = 1& Then
4720            If ![RibbonName] <> "HideTheRibbon" Then
4730              .AddNew
4740              ![RibbonName] = "HideTheRibbon"
4750              ![RibbonXML] = strRibbon_XML
4760              .Update
4770            Else
4780              If ![RibbonXML] <> strRibbon_XML Then
4790                .Edit
4800                ![RibbonXML] = strRibbon_XML
4810                .Update
4820              Else
                    ' ** All's well.
4830              End If
4840            End If
4850          Else
4860            .FindFirst "[RibbonName] = 'HideTheRibbon'"
4870            If .NoMatch = True Then
4880              .AddNew
4890              ![RibbonName] = "HideTheRibbon"
4900              ![RibbonXML] = strRibbon_XML
4910              .Update
4920            Else
4930              If ![RibbonXML] <> strRibbon_XML Then
4940                .Edit
4950                ![RibbonXML] = strRibbon_XML
4960                .Update
4970              Else
                    ' ** All's well.
4980              End If
4990            End If
5000          End If
5010        End If
5020        .Close
5030      End With

5040      .Close
5050    End With

        'DoCmd.ShowToolbar "Ribbon", acToolbarNo completely hides the ribbon.

        'But if this is for a custom application, you probably want some kind of
        'menus and/or toolbars. Of course, you can specify toolbars, ribbons, and
        'menus in your form and report properties.

        'You can also make Access run with your custom Access 2003-style menus and
        'toolbars and no ribbon, by setting the startup properties: File (Alt-F),
        'Access Options, Current Database, Ribbon and Toolbar Options, then turn off
        'Allow Full Menus and Allow Built-in Toolbars. You can set these options
        'programmatically the same as Access 2003 (AllowFullMenus and
        'AllowBuiltInToolbars), but this does not take effect until you reopen the

        ' ** GetHiddenAttribute : Hidden check box on Properties window.
        ' **   ? GetHiddenAttribute(acTable, "USysRibbons") = True
        ' ** SetHiddenAttribute : Hidden check box on Properties window.
        ' **   SetHiddenAttribute acTable, "USysRibbons", False  'Error: 2016 - You can't modify the attributes of System Tables.
        ' ** DoCmd.DeleteObject acTable, "USysRibbons"  'Successful!

        ' ** TableDef Attributes enumeration:
        ' **   -2147483646  dbSystemObject     The table is a system table provided by the Microsoft Jet database engine.
        ' **                                   You can set this constant on an appended TableDef object.
        ' **             1  dbHiddenObject     The table is a hidden table provided by the Microsoft Jet database engine.
        ' **                                   You can set this constant on an appended TableDef object.
        ' **         65536  dbAttachExclusive  For databases that use the Microsoft Jet database engine, the table is a
        ' **                                   linked table opened for exclusive use. You can set this constant on an
        ' **                                   appended TableDef object for a local table, but not on a remote table.
        ' **        131072  dbAttachSavePWD    For databases that use the Microsoft Jet database engine, the user ID and
        ' **                                   password for the remotely linked table are saved with the connection information.
        ' **                                   You can set this constant on an appended TableDef object for a remote table,
        ' **                                   but not on a local table.
        ' **     536870912  dbAttachedODBC     The table is a linked table from an ODBC data source, such as
        ' **                                   Microsoft SQL Server (read-only).
        ' **    1073741824  dbAttachedTable    The table is a linked table from a non-ODBC data source such as a Microsoft Jet
        ' **                                   or Paradox database (read-only).

        ' ** Field Attributes enumeration:
        ' **       1  dbFixedField      The field size is fixed (default for Numeric fields).
        ' **       1  dbDescending      The field is sorted in descending (Z to A or 100 to 0) order; this option applies
        ' **                            only to a Field object in a Fields collection of an Index object. If you omit this
        ' **                            constant, the field is sorted in ascending (A to Z or 0 to 100) order. This is the
        ' **                            default value for Index and TableDef fields (Microsoft Jet workspaces only).
        ' **       2  dbVariableField   The field size is variable (Text fields only).
        ' **      16  dbAutoIncrField   The field value for new records is automatically incremented to a unique Long
        ' **                            integer that can't be changed (in a Microsoft Jet workspace, supported only for
        ' **                            Microsoft Jet database(.mdb) tables).
        ' **      32  dbUpdatableField  The field value can be changed.
        ' **    8192  dbSystemField     The field stores replication information for replicas; you can't delete this type
        ' **                            of field (Microsoft Jet workspaces only).
        ' **   32768  dbHyperlinkField  The field contains hyperlink information (Memo fields only).
        'Access 97 Access 2000 Access 2002 Access 2003 Access 2007

EXITP:
5060    Set idx = Nothing
5070    Set fld = Nothing
5080    Set tdf = Nothing
5090    Set rst = Nothing
5100    Set dbs = Nothing
5110    AC07_Ribbon_Check = blnRetVal
5120    Exit Function

ERRH:
5130    blnRetVal = False
5140    Select Case ERR.Number
        Case Else
5150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5160    End Select
5170    Resume EXITP

End Function

Private Function AC10_Ribbon_Check() As Boolean

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "AC10_Ribbon_Check"

        Dim blnRetVal As Boolean

5210    blnRetVal = AC07_Ribbon_Check  ' ** Function: Above.

EXITP:
5220    AC10_Ribbon_Check = blnRetVal
5230    Exit Function

ERRH:
5240    blnRetVal = False
5250    Select Case ERR.Number
        Case Else
5260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5270    End Select
5280    Resume EXITP

End Function

Private Function ReturnRegKeyValue(ByVal lngKeyToGet As Long, ByVal strKeyName As String, ByVal strValueName As String) As String
' ********Code Start**************
' ** This code was originally written by Terry Kreft
' **  and Dev Ashish.
' ** It is not to be altered or distributed,
' ** except as part of an application.
' ** You are free to use it in any application,
' ** provided the copyright notice is left unchanged.
' **
' ** Code Courtesy of
' **   Dev Ashish & Terry Kreft

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "ReturnRegKeyValue"

        Dim lngHKey As Long, strClassName As String, lngClassLen As Long, lngSecurity As Long
        Dim lngReserved As Long, lngSubKeys As Long, lngMaxSubKeyLen As Long, lngMaxClassLen As Long
        Dim lngValues As Long, lngMaxValueNameLen As Long, lngMaxValueLen As Long
        Dim ftLastWrite As FILETIME
        Dim lngType As Long, lngData As Long
        Dim lngTmp01 As Long
        Dim strRetVal As String, lngRetVal As Long, varRetVal As Variant

        ' ** Open the key first.
5310    lngTmp01 = RegOpenKeyEx(lngKeyToGet, strKeyName, 0&, KEY_READ, lngHKey)  ' ** API Function: modSecurityFunctions.

        ' ** Are we ok?
5320    If lngTmp01 <> ERROR_SUCCESS Then
          'ERR.Raise lngTmp01 + vbObjectError
5330      varRetVal = RET_ERR
5340    Else

5350      lngReserved = 0&
5360      strClassName = String$(MAX_LEN, 0):  lngClassLen = MAX_LEN

          ' ** Get boundary values.
5370      lngTmp01 = RegQueryInfoKey(lngHKey, strClassName, lngClassLen, lngReserved, lngSubKeys, lngMaxSubKeyLen, _
            lngMaxClassLen, lngValues, lngMaxValueNameLen, lngMaxValueLen, lngSecurity, ftLastWrite)  ' ** API Function: Above.

          ' ** How we doin?
5380      If Not (lngTmp01 = ERROR_SUCCESS) Then ERR.Raise lngTmp01 + vbObjectError

          ' ** Now grab the value for the key.
5390      strRetVal = String$(MAX_LEN - 1, 0)
5400      lngTmp01 = RegQueryValueEx(lngHKey, strValueName, lngReserved, lngType, ByVal strRetVal, lngData)  ' ** API Function: modSecurityFunctions.
5410      Select Case lngType
          Case REG_SZ
5420  On Error Resume Next
5430        lngTmp01 = RegQueryValueEx(lngHKey, strValueName, lngReserved, lngType, ByVal strRetVal, lngData)  ' ** API Function: modSecurityFunctions.
5440        If ERR.Number = 0 Then
5450  On Error GoTo ERRH
5460          varRetVal = Left(strRetVal, lngData - 1)
5470        Else
5480  On Error GoTo ERRH
5490          varRetVal = -9
5500        End If
5510      Case REG_DWORD
5520  On Error Resume Next
5530        lngTmp01 = RegQueryValueEx(lngHKey, strValueName, lngReserved, lngType, lngRetVal, lngData)  ' ** API Function: modSecurityFunctions.
5540        If ERR.Number = 0 Then
5550  On Error GoTo ERRH
5560          varRetVal = lngRetVal
5570        Else
5580  On Error GoTo ERRH
5590          varRetVal = -9
5600        End If
5610      Case REG_BINARY
5620  On Error Resume Next
5630        lngTmp01 = RegQueryValueEx(lngHKey, strValueName, lngReserved, lngType, ByVal strRetVal, lngData)  ' ** API Function: modSecurityFunctions.
5640        If ERR.Number = 0 Then
5650  On Error GoTo ERRH
5660          varRetVal = Left(strRetVal, lngData)
5670        Else
5680  On Error GoTo ERRH
5690          varRetVal = -9
5700        End If
5710      End Select

          ' ** All quiet on the western front?
5720      If Not (lngTmp01 = ERROR_SUCCESS) Then
            'ERR.Raise lngTmp01 + vbObjectError
5730        varRetVal = RET_ERR
5740      End If

5750    End If

EXITP:
5760    ReturnRegKeyValue = varRetVal
5770    lngTmp01 = RegCloseKey(lngHKey)
5780    Exit Function

ERRH:
5790    varRetVal = RET_ERR
5800    Select Case ERR.Number
        Case Else
5810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5820    End Select
5830    Resume EXITP
        ' ********Code End**************

End Function
