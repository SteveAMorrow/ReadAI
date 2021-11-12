/*--------------------------------------------------------------------------------------+
*
*      $Source: pid/PidApp.cs $
*
*   $Copyright: (c) 2021 Bentley Systems, Incorporated. All rights reserved. $
*
* +--------------------------------------------------------------------------------------*/

/*--------------------------------------------------------------------------------------+
|
|   Usings
|
+--------------------------------------------------------------------------------------*/
using System;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using EcObj = Bentley.ECObjects;
using Bentley.ECObjects.Instance;
using Bentley.ECObjects.Schema;
using MsDgn = Bentley.Interop.MicroStationDGN;
using BsiLog = Bentley.Logging;
using Bentley.PlantBIMECPlugin;
using CTLog = Bentley.Plant.CommonTools.Logging;
using Bentley.Plant.ConfigVariable;
using BmfAi = Bentley.Plant.MF.AppInterface;
using BmfCad = Bentley.Plant.CadSys;
using BmfMsCad = Bentley.Plant.MSCadSys;
using MsUtils = Bentley.Plant.MSCadSys.MSUtilities;
using BmfGs = Bentley.Plant.MF.GraphicsSystem;
using BmfPer = Bentley.Plant.MFPersist;
using BmfUt = Bentley.Plant.MF.Utilities;
using BmfCat = Bentley.Plant.MF.Catalog;
using BmfECPer = Bentley.EC.Persistence;
using Bom = Bentley.Plant.ObjectModel;
using SchInt = Bentley.Plant.AppFW.Schematics.Interfaces;
using SchApp = Bentley.Plant.AppFW.Schematics;
using SchCel = Bentley.Plant.AppFW.Schematics.Cells;
using SchCat = Bentley.Plant.AppFW.Schematics.Catalog;
using BmfVal = Bentley.Plant.MF.Validation;
using PidInt = Bentley.Plant.App.Pid.Interfaces;
using BmfBKC = Bentley.Plant.MF.BusinessKeyControls;
using BmfPS = Bentley.Plant.MF.PipingSpecification;
using BmfOpenPlant = Bentley.Plant.MF.OpenPlant;
using BmfArb = Bentley.Plant.MFPersist.AltRepositoryBase;
using BCT = Bentley.Plant.CommonTools;
using DPNet = Bentley.DgnPlatformNET;
using System.Xml.Linq;
using System.Linq;
using Bentley.MstnPlatformNET;
using Bentley.Connect.Client.API.V2;
using BmfPSpec = Bentley.Plant.MF.PipingSpecification;

namespace Bentley.Plant.App.Pid
{
/*====================================================================================**/
/// <summary>
/// PID Application class
/// </summary>
/*==============+===============+===============+===============+===============+======*/
[Bentley.MstnPlatformNET.AddInAttribute (MdlTaskID = "PidApp")]
public sealed class PidApplication : Bentley.MstnPlatformNET.AddIn
{

#region DllImports

/*------------------------------------------------------------------------------------**/
/// <summary>Exits microstation without any prompts.</summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
[DllImport ("ustation.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.Cdecl)]
private static extern void mdlSystem_exitMicroStation (int status, int error, int arg);

#endregion

#region Attributes
private const string CMD_DRAWINGFORMERGE = "drawing";
private const string MS_ECREPOSITORY_PROPERTYENABLER_SKIPPED_ECSCHEMAS  = "MS_ECREPOSITORY_PROPERTYENABLER_SKIPPED_ECSCHEMAS";
private const string BMF_PROPERTY_ENABLER_SKIP_LIST                     = "BMF_PROPERTY_ENABLER_SKIP_LIST";
private static PidApplication               s_instance;
private static BmfAi.AppInterface           s_appInterface;
private static SchApp.SchematicsApplication s_schematicsApp;
private static SchCel.Utility               s_cellUtilities;
private static string                       s_appName = "Pid";
private static int                          s_appId = 1786;     // application id is provided by Bentley
private static bool                         s_initialized = false;
private static string                       s_behavioralSchemaName;
private static string                       s_documentClassName;
private static string                       s_currentWorkSetName;
private static IList<PidInt.IPowerPIDAddin> s_powerPIDAddins;
private static string                       s_insertUnparsedParameter;
private static MsDgn.Application            s_msApp;

private static bool                         s_MicroStationMode = false;
private static OPPIDSymbolProvider          s_OPPIDSymbolProvider;
private static ConnectClientAPI             s_connectApi;
private static bool                         s_closeOccurring = false;

/*------------------------------------------------------------------------------------**/
/// <summary>Local member StandardPreferencesDialog</summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static StandardPreferencesDialog m_standardPreferencesDialog = null;


#endregion


/*------------------------------------------------------------------------------------**/
/// <summary>
/// PidApplication constructor
/// Sets singleton s_instance
/// </summary>
/// <author>Dustin.Parkman</author>                             <date>7/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private                         PidApplication
(
IntPtr mdldesc
) : base(mdldesc)
    {
    s_instance = this;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
///Loads custom 3rd Party Addins
/// </summary>
/// <author>Dustin.Parkman</author>                             <date>07/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private  void                   loadPowerPIDAddins
(
)
    {
    string[] powerPIDApps;
    PidInt.IPowerPIDAddin powerPIDAddin = null;
    Assembly ecAssembly = null;

    if (!ConfigurationVariableUtility.GetValueList (out powerPIDApps, "POWERPID_APPS"))
        return;

    for (int i = 0; i < powerPIDApps.Length; i++)
        {
        ecAssembly = Assembly.LoadFrom (powerPIDApps[i]);
        Type[] Types = ecAssembly.GetTypes ();
        foreach (Type clrType in Types)
            {
            if (!isPowerPIDAddin (clrType))
                continue;

            powerPIDAddin = Activator.CreateInstance (clrType, null) as PidInt.IPowerPIDAddin;
            if (powerPIDAddin == null)
                continue;

            if (s_powerPIDAddins == null)
                s_powerPIDAddins = new List<PidInt.IPowerPIDAddin> ();

            s_powerPIDAddins.Add (powerPIDAddin);
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
///Determines if a clrtype implements a PowerPIDAddin interface
/// </summary>
/// <remarks>
/// There is probably a cooler way of doing this with reflection but I could not
/// figure it out.
/// </remarks>
/// <author>Dustin.Parkman</author>                             <date>07/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    isPowerPIDAddin
(
Type clrType
)
    {
    foreach(Type interfaceType in clrType.GetInterfaces())
        {
        if (interfaceType.Name.Contains( "IPowerPIDAddin"))
            return true;
        }
    return false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Force application to exit
/// </summary>
/// <author>Diana.Fisher</author>                              <date>05/2017</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             forceExit
(
)
    {
    mdlSystem_exitMicroStation (0, 0, 0);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Method Run when AddIn is loaded
/// </summary>
/// <param name="commandLine">string from command line</param>
/// <returns>0</returns>
/*--------------+---------------+---------------+---------------+---------------+------*/
protected override int          Run
(
string[] commandLine
)
    {
    if (!isWorkSetSelected ())
        {
        CTLog.FrameworkLogger.Instance.SubmitResponse
                (Bentley.Logging.SEVERITY.LOG_FATAL,
                 CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingPIDCaption"), 
                 CTLog.FrameworkLogger.Instance.GetLocalizedString ("NoWorkspaceProjectSelected"),
                 Bentley.Plant.CommonTools.Logging.MessageBoxButtons.OK,
                 Bentley.Plant.CommonTools.Logging.DialogResponse.OK);

        forceExit ();
        return -1;
        }

    ////// File->Print organizer functionality in projectwise mode spawns a new OPPID process
    ////// for every file added to print set. This is identifiable by presence of MSPrintServer
    ////// environment variable (only spawned process has this variable).
    ////// The issue is that before spawning, all environment variables are converted to lower
    ////// case. This causes issues with say logger initialization which is case sensitive with
    ////// respect to environment variables.
    ////// As a workaround, we are detecting if OPPID has been launched from print organizer
    ////// and if so, resetting env variable APPDATA to upper case (by deleting and re adding).
    ////// This has to be done for every env variable used in framework configuration xml file.
    ////// So far there is just one i.e. APPDATA
    ////// Todo:AliA: Remove this workaround when proper fix is available from Andrew Edge.
    ////// This is with respect to Defect 123673:Print Organizer not responding.
    ////// As per comments of Andrew Edge, this problem might go away itself when we move to
    ////// .net 4.0 or MS 21.x+
    ////if (!string.IsNullOrEmpty (System.Environment.GetEnvironmentVariable ("MSPrintServer")))
    ////    {
    ////    if (!string.IsNullOrEmpty (System.Environment.GetEnvironmentVariable("APPDATA")))
    ////        {
    ////        string appDataValue = System.Environment.GetEnvironmentVariable ("APPDATA");
    ////        System.Environment.SetEnvironmentVariable ("APPDATA", null);
    ////        System.Environment.SetEnvironmentVariable ("APPDATA", appDataValue);
    ////        }
    ////    }

    // The commented code above is to be uncommented if print organizer is unable to show some graphics
    // during print that require oppid element handlers for correct output. For now, we shall skip loading
    // of OPPID altogether in case it is launched from print organizer and evaluate.
    if (!string.IsNullOrEmpty (System.Environment.GetEnvironmentVariable ("MSPrintServer")))
        return 0;

    // AppUnloadedEvent is not working
    //this.UnloadedEvent += (m_unloadedDelegate = new Bentley.MicroStation.AddIn.UnloadedEventHandler (AppUnloadedEvent));
    ModelChangedEvent += new Bentley.MstnPlatformNET.AddIn.ModelChangedEventHandler (AppModelChangedEventHandler);
    
    MstnPlatformNET.Session.Instance.OnMasterFileStart += new MstnPlatformNET.Session.DgnFileEventHandler(OnMasterFileStart);
            MstnPlatformNET.Session.Instance.OnMasterFileChanging += OnMasterFileChanging;
    if (s_initialized)
        return 0;

    // P&ID Application initialization
    if (InitializeApp () &&
        InitializeSchemas ())
        {
        BmfAi.AppInterface.InitializeModel ();
        InitializeForFenceStretchPostProcessing ();

        s_initialized = true;

        //register the property enabler
        Pid.PIDUtilities.RegisterPropertyEnablers (s_behavioralSchemaName);

        InitializeSpecManager ();

        s_msApp = BmfCad.CadApplication.Instance.MSDGNApplication;
        // Register oppid symbol provider to support named expressions
        if (null == s_OPPIDSymbolProvider)
            {
            s_OPPIDSymbolProvider = new OPPIDSymbolProvider();
            s_OPPIDSymbolProvider.RegisterAsSymbolProvider();
            }
        }
    else
        CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.ERROR,
            CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializationFailed"),
            CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializeFailedMessage"),
            CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);

    return 0;
    }
    /*------------------------------------------------------------------------------------**/
/// <summary>
/// Disables calls to DgnFile.ProcessChanges in platform as they cause problem in dwg mode
/// The calls need to remain disabled until file is fully initialized as indicated by call
/// to OnMasterFileStart. Currently ImportSchema and CommitChangeset are methods that 
/// have to be told to suppress call to ProcessChanges via extended data. 
/// </summary>
/// <param name="dgnfile">dgnFile being loaded</param>
/// <author>Ali.Aslam</author>                             <date>5/2017</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    OnMasterFileChanging
(
DPNet.DgnFile dgnfile
)
    {
    BmfPer.PersistenceManager.AvoidCallingProcessChangesOnDgnFile = true;
    }
    /*------------------------------------------------------------------------------------**/
/// <summary>
/// Enables calls to DgnFile.ProcessChanges which are disabled on app start and file open/new
/// as they cause problem in dwg mode
/// The calls need to remain disabled until file is fully initialized as indicated by call
/// to OnMasterFileStart
/// Resuming calls to DgnFile.ProcessChanges here as method indicates load is done. 
/// </summary>
/// <param name="dgnfile">dgnFile being loaded</param>
/// <author>Ali.Aslam</author>                             <date>5/2017</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public void                     OnMasterFileStart
(
Bentley.DgnPlatformNET.DgnFile dgnfile
)
    {
    BmfPer.PersistenceManager.AvoidCallingProcessChangesOnDgnFile = false;
    // Since all calls to ProcessChanges were suppressed until now, call it manually to save changes. 
    dgnfile.ProcessChanges(DgnPlatformNET.DgnSaveReason.ApplicationInitiated);
    ECSystem.Extensibility.ExtensionDirectory.LoadExtension("Bentley.DgnDisplayNET.dll");
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Checks whether a valid WorkSet has been selected.
/// Sets private variable to name of selected.
/// </summary>
/// <author>Ali.Aslam</author>                             <date>5/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    isWorkSetSelected
(
)
    {
    string workSetName = string.Empty;

    ConfigurationVariableUtility.GetValue (out workSetName, "_USTN_WORKSETNAME");
    if (string.IsNullOrEmpty (workSetName))
        return false;

    s_currentWorkSetName = workSetName;
    return true;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Checks if WorkSet changed, if so resets current value
/// </summary>
/// <remarks>
/// PowerPlatform does not allow switching WorkSet within an session
/// </remarks>
/// <returns>true or false</returns>
/// <author>Diana.Fisher</author>                               <date>4/2018</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    didWorkSetChange
(
)
    {
    string workSetName = string.Empty;

    ConfigurationVariableUtility.GetValue (out workSetName, "_USTN_WORKSETNAME");
    // this should never happen, since the first thing Run() does is verify it has value
    if (string.IsNullOrEmpty (workSetName))
        return true;

    if (s_currentWorkSetName.Equals (workSetName))
        return false;

    // update current workset name
    s_currentWorkSetName = workSetName;
    return true;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is application Initialized
/// If application is not Initialized, display message
/// The application may not be initialized because of file types:
///     dgnlib, cel, i-model, overlay,...
/// </summary>
/// <author>Steve.Morrow</author>                             <date>5/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             isApplicationInitialized
    {
    get
        {
        if (!s_initialized || !isActiveModelSet)
            {
            CTLog.Logger.StaticSeverityLevelOverrideEnabled = true;
            CTLog.Logger.StaticLoggingEnabledOverride = CTLog.GlobalPropertyOverride.PropertyEnabled;
            CTLog.Logger.StaticSeverityLevelOverride = Bentley.Logging.SEVERITY.WARNING;
            string message = CTLog.FrameworkLogger.Instance.GetLocalizedString ("OPPIDCommandsEdit");
            CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.WARNING, message);
            if (Bom.Workspace.InDwg && Bom.Workspace.IsActiveModelSheet)
                {
                message = CTLog.FrameworkLogger.Instance.GetLocalizedString ("DWGSheetModelInvalid");
                CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.WARNING, message);
                }

            CTLog.Logger.StaticLoggingEnabledOverride = CTLog.GlobalPropertyOverride.PropertyDisabled;
            CTLog.Logger.StaticSeverityLevelOverrideEnabled = false;
            CTLog.Logger.StaticSeverityLevelOverride = Bentley.Logging.SEVERITY.FATAL;

            // Standard OpenPlant PID creation tools cannot be used to edit this file
            string showDialog = "";
            ConfigurationVariableUtility.GetValue (out showDialog, "BMF_SHOW_INVALIDDESIGNFILE_MESSAGEBOX");
            if (showDialog != null && showDialog.Equals ("1"))
                {
                string caption = CTLog.FrameworkLogger.Instance.GetLocalizedString ("OPPIDCommands");
                CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.INFO, caption,
                    message, CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
                }

            return false;
            }

        if (BmfPer.PersistenceManager.IsWorkingDocumentReadOnly)
            return false;

        return true;
        }
    }

#region PersistenceManager
// This isn't working
//private void AppUnloadedEvent
//(
//Bentley.MicroStation.AddIn sender,
//Bentley.MicroStation.AddIn.UnloadedEventArgs eventArgs
//)
//    {
//    this.NewDesignFileEvent -= m_NewDesignFileDelegate;
//    this.ModelChangedEvent -= m_ModelChangedDelegate;
//    if (eventArgs.UnloadKind == Bentley.MicroStation.AddIn.UnloadReasons.Shutdown)
//        BmfPer.PersistenceManager.CloseModel ();
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Called by MicroStation when a change to a Model is about to occur or has occurred.
/// We're only interested in when the Change is set to Active, which means a Model
/// has been opened and made active.
/// </summary>
/// <param name="sender">The AddIn originator</param>
/// <param name="eventArgs">The ModelChangedEventArgs</param>
///
/// <author>Ed.Becnel</author>                                  <date>09/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    AppModelChangedEventHandler
(
Bentley.MstnPlatformNET.AddIn sender,
Bentley.MstnPlatformNET.AddIn.ModelChangedEventArgs eventArgs
)
    {
    switch (eventArgs.Change)
        {
        case ModelChangedEventArgs.ChangeType.BeforeActive:
            // TODO: If the type of the model has been changed, then can we deduce that the project type has been changed?
            if (didWorkSetChange ())
                {
                // Initialize the application and catalog server
                InitializeApp ();

                // Re-initialize the schemas
                InitializeSchemas ();

                // Refresh the schema analyzer
                BmfPer.PersistenceManager.RefreshSchemaAnalyzer ();
                }
            break;
        case ModelChangedEventArgs.ChangeType.Active:
            //TODO shouldn't this be moved to Model.Open?
            if (isFileValidForPersistenceManagerOpen (eventArgs))
                BmfPer.PersistenceManager.OpenModel ();

            PIDUtilities.InitializeReports();
            // Reinitialize the logging provider to handle any new loggers created 
            CTLog.FrameworkLogger.Instance.ReInitializeProvider();
            break;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is File Valid For PersistenceManager Open
/// If file type is CEL (cell Library) or DGNLIB or DgnModelType is not Normal(referenced in dgn),
///    do not open with Persistence Manager
/// <param name="eventArgs">ModelChangedEventArgs </param>
/// </summary>
/// <returns>true if file is ok to open with PersistenceManager; otherwise false</returns>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    isFileValidForPersistenceManagerOpen
(
ModelChangedEventArgs eventArgs
)
    {
    //IntPtr activeDgnFileIntPtr = eventArgs.ModelReference.DgnFileIntPtr;
    bool isValid = true;
    try
        {
        string fileName = eventArgs.DgnModelRef.GetDgnFile ().GetFileName ();
        if(string.IsNullOrEmpty (fileName))
            return false;

        string ext = Path.GetExtension (fileName);
        if (ext.Equals (".CEL", StringComparison.OrdinalIgnoreCase) ||
            ext.Equals (".DGNLIB", StringComparison.OrdinalIgnoreCase) ||
            Bentley.DgnPlatformNET.DgnModelType.Normal != eventArgs.DgnModelRef.ModelType)
            isValid = false;
        }
    catch (Exception ex)
        {
        isValid = false;
        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.WARNING, ex.Message);
        }
    return isValid;
    }

#endregion

#region Startup dialog

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Workspace model has been upgraded
/// </summary>
/// <param name="modelToUpgrade">model potentially required upgrade</param>
/// <param name="majorNumber">model major number before upgrade</param>
/// <param name="minorNumber">model minor number before upgrade</param>
/// <param name="revisionNumber">model revision number before upgrade</param>
/// <param name="buildNumber">model build number before upgrade</param>
/// <returns>true if upgraded; otherwise false</returns>
/// <author>Diana.Fisher</author>                              <date>02/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    onWorkspace_ModelUpgraded
(
Bom.IModel modelToUpgrade,
int majorNumber,
int minorNumber,
int revisionNumber,
int buildNumber
)
    {

    int numCheck = 7;

//#if DEBUG
//    numCheck = 9;
//#endif
     if ((10 == majorNumber) && (minorNumber >= numCheck))
        return;

    //set packet to be passed for upgrading
    DrawingUpgrade.UpgradePacket upgradePacket = new DrawingUpgrade.UpgradePacket
                        (majorNumber, minorNumber, revisionNumber, buildNumber, modelToUpgrade);
    upgradePacket.ApplicationObjectRequired = true; //oppid specific

    PidUpgradeHelper upgrader = new PidUpgradeHelper (upgradePacket);
    if (upgrader.Valid)
        upgrader.UpgradeModel ();
    else
        forceExit ();
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Workspace model has been opened
/// </summary>
/// <param name="sender">model</param>
/// <param name="args">null</param>
/// <author>Steve.Morrow</author>                                  <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    onWorkspace_ModelOpened
(
object sender,
EventArgs args
)
    {
    Bom.IModel model = (Bom.IModel)sender;
    if (null == model)
        return;

    // if this event has been triggered on a saved model, do nothing
    if (Bom.Workspace.DocumentSaved)
        {
        Bom.Workspace.DocumentSaved = false;
        return;
        }

    // Check for valid project
    if (!isAValidDocumentInProject (model))
        {
        string currentFileName = BmfAi.AppInterface.CurrentDocumentFileName;

        string messageBody1 = String.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString ("Theopeneddocumentbelongstoaprojectwhichisnotcurrentlyset"));
        string messageBody2 = String.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString ("Thedocumentshouldbeclosedandreopenedwithavalidmatchingproject"));
        string messageBody3 = String.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString ("Doyouwanttocontinue"));
        string messageBody = messageBody1 + "\n" + messageBody2 + "\n" + messageBody3;
        string messageCaption = String.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString ("InvalidProjectForDocument"));
        CTLog.DialogResponse result = CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.WARNING,
            messageCaption, messageBody, CTLog.MessageBoxButtons.YesNo, CTLog.DialogResponse.No);
        if (result == CTLog.DialogResponse.No)
            {
            BmfCad.ICadDocument activeCadDoc = BmfGs.GFXSystem.ActiveHost.CadDocumentServer.GetActiveCadDocument ();
            activeCadDoc.Close ();
            }
        else
            BmfAi.AppInterface.Workspace.CloseModel (currentFileName);

        return;
        }

    showStartUp ((Bom.IModel)sender, false, false, false, 0);

    // Refresh properties
    refreshPropertiesAndRedrawOnStartup (false);

    // register post database sync event
    Bom.Workspace.ActiveModel.DatabaseSynchronized += new Bom.DatabaseSynchronizedEventHandler (onDatabaseSynchronized);

    // Notify Addins that file was opened
    if (null != s_powerPIDAddins)
        {
        foreach (PidInt.IPowerPIDAddin addin in s_powerPIDAddins)
            addin.OnModelOpen ((Bom.IModel)sender);
        }

    bool shouldForceExit = false;
    SchApp.SchematicsApplication.HandleModelOpened (out shouldForceExit);
    if (shouldForceExit)
        forceExit ();

    // ensure spec manager is set for current specificiation of ActiveModel
    string activeSpec = PIDUtilities.GetActiveSpecification (null);
    BmfPSpec.SpecSystem.ActiveSpecManager.SetCurrentSpec (activeSpec);

    // determine if active associated items dialog should be opened
    if (StateShowStandardPreferencesDialog)
        ShowStandardPreferencesDialog ();

    s_closeOccurring = false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Workspace model has been closed
/// </summary>
/// <param name="sender">model</param>
/// <param name="args">null</param>
/// <author>Dustin.Parkman</author>                             <date>7/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    onWorkspace_ModelClosed
(
object sender,
EventArgs args
)
    {
    s_closeOccurring = true;
    CloseStandardPreferencesDialog (); 

    //Notify Addins that file was closed
    if (null != s_powerPIDAddins)
        {
        foreach (PidInt.IPowerPIDAddin addin in s_powerPIDAddins)
            addin.OnModelClosed ((Bom.IModel)sender);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Model Conflicts event, user canceled moved, renamed dialog
/// </summary>
/// <remarks>
/// Workspace handles Copied
/// PidApp only deals with NotHandled conflicts
/// </remarks>
/// <param name="model">model with conflicts</param>
/// <param name="conflict">conflict</param>
/// <param name="text">
/// error message if NotHandled
/// new cad document name if MovedOrRenamed
/// </param>
/// <author>Diana.Fisher</author>                              <date>06/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    onWorkspace_ModelConflicts
(
Bom.IModel model,
Bom.FileConflict conflict,
string text
)
    {
    if (Bom.FileConflict.NotHandled == conflict)
        {
        string applicationExitMessage = BCT.Logging.FrameworkLogger.Instance.GetLocalizedString("ModelConflict_ApplicationExitMessageWithoutReason");
        if (!string.IsNullOrEmpty(text))
            applicationExitMessage = string.Format(BCT.Logging.FrameworkLogger.Instance.GetLocalizedString("ModelConflict_ApplicationExitMessageWithReason"), text);

        string caption = BCT.Logging.FrameworkLogger.Instance.GetLocalizedString ("ModelConflict_ApplicationExitMessage_Caption");
        MessageBox.Show (applicationExitMessage, caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        forceExit ();
        }

    SchApp.SchematicsApplication.HandleModelConflicts (model, conflict, text);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refresh Properties on model open if cfg variable set
/// </summary>
/// <param name="forceRefresh">force refresh</param>
/// <author>Steve.Morrow</author>                                  <date>08/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             refreshPropertiesAndRedrawOnStartup
(
bool forceRefresh
)
    {
    if (!forceRefresh)
        {
        string cfgVar = "";
        ConfigurationVariableUtility.GetValue (out cfgVar, "BMF_REFRESH_PROPERTIES_ON_FILEOPEN");
        if (null == cfgVar || cfgVar.Length == 0)
            return;

        //if variable is set to 0 exit
        if (cfgVar.Equals ("0"))
            return;
        }

    Bom.Workspace.ActiveModel.RedrawWithRefreshedProperties ();
    refreshActiveModelAndUpdateTitleSheets ();
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Check to see if currently opened document belongs to the current set project
/// </summary>
/// <param name="model">model</param>
/// <author>Steve.Morrow</author>                                  <date>11/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    isAValidDocumentInProject
(
Bom.IModel model
)
    {
    bool status = true;

    //check to see if the model is Initialized
    //if the model is Initialized, this denotes that the model belongs to  the currently set project
    //if the model is not Initialized, this denotes that the model is either new or belongs to a project that is not currently set
    if (!model.IsInitialized)
        {

        //See if there are any components in the activemodel, if not then the drawing is new
        Bom.IComponentCollection components = Bom.Workspace.ActiveModel.FindAllComponentsInActiveModel ();
        if (null == components || components.Count == 0)
            return true;

        //check to see if there is an ecinstance that matches the document class name which is stored in the pcf file
        //if the class name does not exist, then the document belongs to a different project.
        components = Bom.Workspace.ActiveModel.FindAllComponentsInActiveModelByClassName (s_documentClassName);
        if (null == components || components.Count == 0)
            return false;

        return true;
        }

    return status;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Check too see if active model is set
/// </summary>
/// <author>Steve.Morrow</author>                                  <date>02/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              isActiveModelSet
    {
    get
        {
        if (null == Bom.Workspace.ActiveModel)
            return false;

        return true;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Force Startup dialog to show
/// </summary>
/// <param name="allowBorderCellChange">allow the border to be changed</param>
/// <param name="tabPageIndex">set the tab page index</param>
/// <author>Steve.Morrow</author>                                  <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             forceShowStartUp
(
bool allowBorderCellChange,
int tabPageIndex
)
    {
    // still have to determine whether it's a new drawing because the user may have pressed
    // cancel on initial startup dialog and later opened it manually. That would still qualify
    // as a new drawing until OK is pressed and ECInstance of PID_DOCUMENT written to file
    showStartUp (Bom.Workspace.ActiveModel, true, false, allowBorderCellChange, tabPageIndex);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Private Startup member
/// Determine if the dgn is new or existing
/// </summary>
/// <param name="model">BOM.Model</param>
/// <param name="forceShow"> force dialog to show</param>
/// <param name="forceHide">force the hide of the startup dialog</param>
/// <param name="allowBorderCellChange">allow the border to be changed</param>
/// <param name="tabPageIndex">set the tab page index</param>
/// <author>Steve.Morrow</author>                                  <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             showStartUp
(
Bom.IModel model,
bool forceShow,
bool forceHide,
bool allowBorderCellChange,
int tabPageIndex
)
    {
    if (null == model)
        return;

    if (!BmfUt.MiscellaneousUtilities.IsValidDesignFile (model.ModelName))
        return;

    //Force the hide of the startup dialog
    if (forceHide)
        return;

    //Check to see if this is a NEW drawing. If NEW then force showing of settings.
    bool showStartUp = !model.IsInitialized;

    //if the drawing is not new, check to see if the user wants to hide the startup dialog
    //the value is stored in the registry
    if (!showStartUp)
        {
        int val = (int)BmfUt.StartupRegistryUtilities.GetRegistryValue ("checkBoxShowDialogOnStartup", 1);
        if (val == 0)
            showStartUp = false;
        else
            showStartUp = true;
        }

    //force the startup dialog to show
    if (forceShow)
        showStartUp = true;

    //If multiple values are set (suppress startup, model not initialized and not components in dgn), do
    //not show startup dialog. This denotes conversion batch mode
    if (suppressStartupDialog (model))
        return;

     //if _USTN_MSMODE == true, do not show startup dialog
     if(s_MicroStationMode)
        showStartUp = false;

     if (showStartUp)
        showBusinessKeyDialog (model, "PIDDoc");
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Determine if startup dialog is to be suppressed
/// This is qualified by a variable set in cfg, model IsInitialized and number of
/// components in dgn
///
/// pid.cfg
/// # Variable to suppress startup dialog (for conversion)
/// PID_SUPPRESS_STARTUP_DIALOG=1
/// <param name="model">Model</param>
/// </summary>
/// <returns>true if in conversion batch mode</returns>
/// <author>Steve.Morrow</author>                                  <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             suppressStartupDialog
(
Bom.IModel model
)
    {
    //1st check for batch mode variable
    string batchMode = null;
    ConfigurationVariableUtility.GetValue (out batchMode, "PID_SUPPRESS_STARTUP_DIALOG");
    if (batchMode == null || batchMode.Length == 0 || batchMode.Equals ("0"))
        return false;

    //if model Is Initialized return false
    if (model.IsInitialized)
        return false;

    //if model does not have components then return false
    BmfCad.ISelectionSetServices selectionSetServices = BmfGs.GFXSystem.ActiveHost.CadDocumentServer.GetActiveCadDocument ().SelectionSetServices;
    bool status = selectionSetServices.DoesActiveModelReferenceHaveElements ();
    if (!status)
        return false;

    //set IsInitialized
    model.IsInitialized = true;
    BmfPer.PersistenceManager.Update ((EcObj.Instance.IECInstance)model);

    //all conditions set return true for converion batch mode
    return true;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show BusinessKey Dialog
/// <param name="model">Model</param>
/// <param name="helpTopic">help topic string</param>
/// </summary>
/// <author>Diana.Fisher</author>                                  <date>07/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             showBusinessKeyDialog
(
Bom.IModel model,
string helpTopic
)
    {
    IList<BmfBKC.IECBusinessKeyControlAddIn> addins = BmfBKC.AddInManager.GetComponentAddins (s_documentClassName, true);

    // create and show business key dialog
    BmfBKC.ECBusinessKeyControl ecBusinessKeyControl = new BmfBKC.ECBusinessKeyControl (model);

    // existing document or new...
    ((EcObj.Instance.IECInstance)model).ExtendedData.Clear ();

    string previousBusinessKey = BusinessKeyUtility.GetValue (model);

    // TODO Vancouver call ReserveCodeValue for pid doc, if newly created?
    IECInstance modelInstance = ecBusinessKeyControl.ShowECBusinessKeyForInstance (model, addins, 
                                        false, true, helpTopic);
    if (null != modelInstance)
        {
        // user pressed OK, check if business key changed
        string businessKey = BusinessKeyUtility.GetValue (modelInstance);
        if (!previousBusinessKey.Equals (businessKey))
            SchApp.SchematicsApplication.HandleModelConflicts (model, Bom.FileConflict.Renamed, businessKey);

        // should active associated items dialog be shown or not
        if (StateShowStandardPreferencesDialog)
            ShowStandardPreferencesDialog ();
        else
            CloseStandardPreferencesDialog ();
        }

    // Check to see if model ecinstance is set or model is not initialized
    // On a new drawing if the cancel is selected on settings dialog, the IsInitialized is false
    // This could cause errors, so on a new drawing the IsInitialized is set to true, even on a cancel
    if (null != modelInstance || !model.IsInitialized)
        {
        model.IsInitialized = true;
        BmfPer.PersistenceManager.Update ((EcObj.Instance.IECInstance)model);
        }
    }


#endregion


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Post process database synchronization
/// </summary>
/// <param name="syncDirection">direction of database synchronization</param>
/// <author>Diana.Fisher</author>                               <date>10/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    onDatabaseSynchronized
(
int syncDirection
)
    {
    BmfArb.Sync_Direction syncDir = (BmfArb.Sync_Direction) syncDirection;

    if (BmfArb.Sync_Direction.ALTREPO_2_DGN == syncDir)
        refreshActiveModelAndUpdateTitleSheets ();
    }

#region Initialization methods
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Display time information as noun plus verb plus current time and milleseconds
/// </summary>
/// <remarks>since this is strictly for internal use, translation is not required</remarks>
/// <param name="methodName">name of method</param>
/// <param name="useStartedVerb">true for started, false for ended</param>
/// <author>Diana.Fisher</author>                              <date>10/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    displayTime
(
string methodName,
bool useStartedVerb
)
    {
    string displayVerb;
    if (useStartedVerb)  // translation NOT required, strictly internal use
        displayVerb = "started";
    else
        displayVerb = "ended";

    BmfCat.CatalogServer.DisplayTime (methodName, displayVerb);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Initialize schema paths and load schemas
/// </summary>
/// <returns>true if successfully initialized; otherwise false</returns>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    InitializeSchemas
(
)
    {
    string[] pidCatalogs;
    int i;

    try
        {
        // TODO Feature Tracking 2.0 load time?
        displayTime ("InitializeSchemas", true);
        if(!s_MicroStationMode)
            Bentley.UI.Controls.WinForms.WaitForm.ShowForm (CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingPIDCaption"),
            CTLog.FrameworkLogger.Instance.GetLocalizedString ("LoadingSchemas"));

        // load libraries
        BmfCat.Catalog catalog = BmfCat.CatalogServer.LoadCatalog (PidCatalogName, 1, 0, EcObj.Schema.SchemaMatchType.Latest);
        if (null == catalog)
            return false;

        // behavioral pid schema should be read only
        catalog.IsReadOnly = true;

        // read other schemas
        if (ConfigurationVariableUtility.GetValueList (out pidCatalogs, Bom.BmfConstant.ProjectSchemaNames))
            {
            for (i = 0; i < pidCatalogs.Length; i++)
                {
                catalog = BmfAi.AppInterface.LoadCatalog (pidCatalogs[i], 1, 0, EcObj.Schema.SchemaMatchType.Latest);
                if (null == catalog)
                    return false;
                }
            }

        if (ConfigurationVariableUtility.GetValueList (out pidCatalogs, Bom.BmfConstant.BehavioralSchemaNames))
            {
            s_behavioralSchemaName = pidCatalogs[0]; // DN 031910: Main behavioral schema must be first in list
            for (i = 0; i < pidCatalogs.Length; i++)
                {
                catalog = BmfAi.AppInterface.LoadCatalog (pidCatalogs[i], 1, 0, EcObj.Schema.SchemaMatchType.Latest);
                if (null == catalog)
                    return false;

                //Support for external reports
                if(OPEF.EC.ECHelper.Instance.Schema == null)
                    {
                    OPEF.Reporting.Controllers.SelectorDlgStateStorage.EnableReferencesFlagInUI = false;

                    //added this to stop OPEF.EC.ECHelper.Instance getting set to null in OPSE Add.cs in Addin_NewDesignFileEvent()
                    OPEF.EC.ECHelper.Instance.AllowReset = false;
                    OPEF.EC.ECHelper.Instance.Schema = catalog.CatalogSchema;
                    OPEF.EC.ECHelper.Instance.OnGetQueryResults -= OnGetQueryResults;
                    OPEF.EC.ECHelper.Instance.OnGetQueryResults += OnGetQueryResults;
                    }
                }
            }

        initializeAssociatedItemsUtility ();

        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.INFO, CTLog.FrameworkLogger.Instance.GetLocalizedString ("CatalogsLoaded"));
        }
    catch (System.Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    finally
        {
        Bentley.UI.Controls.WinForms.WaitForm.CloseForm ();
        displayTime ("InitializeSchemas", false);
        }

    return true;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// On Get QueryResults event
/// This event is used to return a query result for Reports
/// </summary>
/// <param name="ecQuery">ecquery</param>
/// <param name="queryResults">query results</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    OnGetQueryResults
(
BmfECPer.Query.ECQuery ecQuery, 
out BmfECPer.QueryResults queryResults
)
    {
    queryResults = PIDUtilities.GetReportQueryResults (ecQuery);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Initialization related to handling post fence stretch updates.
/// </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    InitializeForFenceStretchPostProcessing
(
)
    {
    IECClass[] runClasses = BmfCat.CatalogServer.GetDerivedClasses(SchApp.SchematicsConstant.SchematicsRun,
        SchApp.SchematicsConstant.SchematicsBehavioralCatalogName);
    if (null != runClasses)
        foreach (IECClass runClass in runClasses)
            BmfCad.CadApplication.Instance.AddStretchableClassName (runClass.Name);
    BmfCad.CadApplication.Instance.StretchableVerticesPropertyName = SchApp.SchematicsConstantProperty.Vertices;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Function intiailizes Bom.Workspace creates a Bom.Model for
/// the current drawing and
/// creates a PlantCad Host for the current application.
/// User can access that Host through K2GfxSystem.oHost property.
/// This function should be called from the application and the
/// application should contain the BmfAi.AppInterface reference
/// for component creation and performing other operations on a
/// Bom.Model.
/// </summary>
/// <returns>true if successfully initialized; otherwise false</returns>
/*--------------+---------------+---------------+---------------+---------------+------*/
private bool                    InitializeApp
(
)
    {
    bool initialized = false;
    string msMode = string.Empty;

    //Check for MicroStation mode. This value is passed in the command line: OpenPlantPID.exe -ws_USTN_MSMODE=True
    //this variable is used tp suppress loading dialogs
    ConfigurationVariableUtility.GetValue (out msMode, "_USTN_MSMODE");
    if (!string.IsNullOrEmpty (msMode) && msMode.Equals ("true", StringComparison.InvariantCultureIgnoreCase))
        s_MicroStationMode = true;

    try
        {
        displayTime ("InitializeApp", true);

        // initialize framework logger
        InitializeLogger ();

        s_connectApi = new ConnectClientAPI();

        ///-------------------------------------------------------------------------------------------------------------------------
        Bentley.UI.Controls.WinForms.WaitForm.ShowForm (CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingPIDCaption"),
        CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingApplication"));
        if(s_MicroStationMode)
            UI.Controls.WinForms.WaitForm.CloseForm ();
        ///-------------------------------------------------------------------------------------------------------------------------

        // initialize registry path
        BmfUt.BMFRegistryPathConstants.ApplicationRegistryPath = @"Software\Bentley\OpenPlant PID CONNECT Edition\" + s_currentWorkSetName;

        // initialize application interface
        s_appInterface = new BmfAi.AppInterface (s_appName, BmfCad.HostName.MicroStation);

        // now that the CadSystem has been created, its console logger can be set
        CTLog.FrameworkLogger.Instance.Logger.ConsoleLogAppender = BmfGs.GFXSystem.ActiveHost.ConsoleLogAppender;

        // register for MicroStation cad element events
        // can't be done in Bmf, because its cad specific
        BmfMsCad.Registration.RegisterCadSystemElementHandler ();

        // set the host's match lifetime addin
        BmfMsCad.IBmfMsHost iHost = BmfGs.GFXSystem.ActiveHost as BmfMsCad.IBmfMsHost;
        iHost.MatchLifetime = this;

        //Setup help values
        BmfUt.Helper.SetupHelp ("OpenPlantPIDExt", "10.00.00.00","OpenPlantPID", s_appId);

        //Check for valid pid applcation project type.
        //Variable  :   PID_APPLICATION_PROJECT_TYPE=OpenPlantPID
        //This variable is defined in a pcf file (ansi.pcf, iso.pcf)
        //If this variable is not found or is not of a known specified type, exit initialize applcation.
        string applicationType = "";
        bool validApplicationType = false;

        ConfigurationVariableUtility.GetValue (out applicationType, PIDConstants.PidApplicationProjectType);
        if (string.IsNullOrEmpty (applicationType))
            validApplicationType = false;
        else
            {
            switch ( applicationType.ToUpper() )
                {
                case "OPENPLANTPID":
                    validApplicationType = true;
                    break;
                default:
                    validApplicationType = false;
                    break;
                }
            }

        if (!validApplicationType)
            {
            Bentley.UI.Controls.WinForms.WaitForm.CloseForm();
            string msgCaption = String.Format(CTLog.FrameworkLogger.Instance.GetLocalizedString("PidApplicationProjectTypeNotDefinedCaption"));
            string msgBody1 = String.Format(CTLog.FrameworkLogger.Instance.GetLocalizedString("PidApplicationProjectTypeNotDefinedBody1"), s_currentWorkSetName);
            string msgBody2 = String.Format(CTLog.FrameworkLogger.Instance.GetLocalizedString("PidApplicationProjectTypeNotDefinedBody2"));
            string msgBody = msgBody1 + "\n" + msgBody2;
            CTLog.FrameworkLogger.Instance.SubmitResponse(BsiLog.SEVERITY.INFO, msgCaption,
                msgBody, CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
                    
            forceExit();
            return false;
            }

        // determine the document type to create
        s_documentClassName = BmfPer.PersistenceManager.DocumentClassName;
        if (s_documentClassName == null || s_documentClassName.Length == 0)
            throw new ApplicationException (String.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString ("CfgVariableNotDefined"),
                Bom.BmfConstant.DocumentTypeName));

        if (BmfAi.AppInterface.InitWorkSpace (s_appId, s_currentWorkSetName, s_documentClassName))
            {
            // set the model scope type model name
            initializeBusinessKeyUtility ();
            DrawingUpgrade.VersionUtility.ApplicationID = s_appId;

            // Schematics Application utilities
            s_schematicsApp = new SchApp.SchematicsApplication ();

            // initialize schematics
            if (SchApp.SchematicsApplication.Initialize ())
                {
                // Schematics Cell utilities. The class has static variables so dispose any existing instance
                // before calling new otherwise calling new has no effect
                if (null != s_cellUtilities)
                    s_cellUtilities.Dispose ();
                s_cellUtilities = null;
                s_cellUtilities = new SchCel.Utility ();

                // determine the angle units of measure preference
                // set this value in Bmf Draw - used by DefineShape methods
                BmfGs.Draw.AngleUnits = BmfGs.ANGLE_UNITS.Degrees;
                string angleUnitsOfMeasure;
                if (ConfigurationVariableUtility.GetValue (out angleUnitsOfMeasure, PIDConstants.PIDAngleUnits))
                    {
                    if (angleUnitsOfMeasure.Equals (PIDConstants.Radians, StringComparison.OrdinalIgnoreCase))
                        BmfGs.Draw.AngleUnits = BmfGs.ANGLE_UNITS.Radians;
                    }

                //Check to if Addins exist and load them into memory
                loadPowerPIDAddins ();

                // register model events
                BmfAi.AppInterface.Workspace.ModelOpened += new Bom.ModelOpenedEventHandler (onWorkspace_ModelOpened);
                BmfAi.AppInterface.Workspace.ModelClosed += new Bom.ModelClosedEventHandler (onWorkspace_ModelClosed);
                BmfAi.AppInterface.Workspace.ModelUpgraded += new Bom.ModelUpgradedFromPreviousRevisionEventHandler (onWorkspace_ModelUpgraded);
                BmfAi.AppInterface.Workspace.ModelConflicts += new Bom.ModelConflictsEventHandler (onWorkspace_ModelConflicts);

                if (s_powerPIDAddins != null)
                    BmfPer.PersistenceManager.ComponentChanged += new BmfPer.ComponentChangedEventHandler (componentChanged);
                initialized = true;
                }
            }
        }
    catch (System.Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    finally
        {
        Bentley.UI.Controls.WinForms.WaitForm.CloseForm ();
        displayTime ("InitializeApp", false);
        }

    return initialized;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// initialize business key utility
/// </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    initializeBusinessKeyUtility
(
)
    {
    // TODO what to pass for app name?
    BusinessKeyUtility.InitializeClientData ( "OpenPlant_PID", true);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Initialize associated items list using application specific backwards compatibility
/// </summary>
/// <remarks>event triggered from AssociatedItemsUtility</remarks>
/// <param name="schema">schema</param>
/// <returns>dictionary with 
///    key = associated item class
///    value = relationship class for key associated item class
/// </returns>
/// <author>Diana.Fisher</author>                              <date>08/2018</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private Dictionary <IECClass, IECRelationshipClass> onAssociatedItems_initializeBackwardCompatibilty
(
IECSchema schema
)
    {
    string associatedItemDefaultListClassName= "ASSOCIATED_ITEMS_LIST";
    string cfgVarValue = null;
    if (ConfigurationVariableUtility.GetValue (out cfgVarValue, "ASSOCIATED_ITEM_DEFAULT_LIST_CLASSNAME"))
        associatedItemDefaultListClassName = cfgVarValue;
    //else
    //    associatedItemDefaultListClassName = "ASSOCIATED_ITEMS_LIST";

    BmfUt.DisplayListSchemaUtilities listUtilities = new BmfUt.DisplayListSchemaUtilities ();
    ArrayList list = listUtilities.GetArrayListDisplay (associatedItemDefaultListClassName, listUtilities.StringListName);
    if (null == list)
        return null;

    Dictionary <IECClass, IECRelationshipClass> classes = new Dictionary <IECClass, IECRelationshipClass> ();
    for (int index = 0; index < list.Count; index++)
        {
        string className = list[index].ToString ();
        className = BmfCat.CatalogServer.StripClassName (className);
        IECClass currentClass = schema.GetClass (className);
        if (null != currentClass)
            {
            IECRelationshipClass currentRelClass = null;
            IECInstance custAttrInstance = currentClass.GetCustomAttributes (Bom.BmfConstantClass.BmfPropertyCustomAttributes);
            if (null != custAttrInstance)
                {
                string relClassName = BCT.ECPropertyValueUtility.GetString (custAttrInstance, "DefaultValue", String.Empty);
                relClassName.Trim ();
                relClassName = BmfCat.CatalogServer.StripClassName (relClassName);
                currentRelClass = schema.GetClass (relClassName) as IECRelationshipClass;
                }

            classes.Add (currentClass, currentRelClass);
            }
        }

    return classes;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Associated Item ClassName List
/// This class name is stored in the ASSOCIATED_ITEM_LIST_POINTER class in the ASSOCIATED_ITEM_CLASS custom attribute.
/// </summary>
/// <param name="ecInstance">ecInstance</param>
/// <returns>Class name to get Associate list from</returns>
///<author>Steve.Morrow</author>                                <date>08/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
//private static string           getAssociatedItemClassNameList
//(
//IECInstance ecInstance
//)
//    {
            // TODO: AssociatedItemsUtility below is backwards compatibility for associate item classes for a component
    ////Get the property from the current ecinstance
    //IECInstance ecInstanceCustomAttrib = ecInstance.ClassDefinition.GetCustomAttributes ("ASSOCIATED_ITEM_LIST_POINTER");
    //if (ecInstanceCustomAttrib == null)
    //    return null;

    ////Get the value from above property
    //return BCT.ECPropertyValueUtility.GetString (ecInstanceCustomAttrib, "ASSOCIATED_ITEM_CLASS", null);
   // }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get associated item class names for a component using application specific backwards compatibility
/// Event will be raised if GetClassNamesForComponent does not find any associated items classes
/// </summary>
/// <param name="componentClass">component class</param>
/// <param name="stripPrefix">true to strip schema prefix, if exists; otherwise false</param>
/// <returns>list of Associated Item class names or null if none</returns>
/// <author>Diana.Fisher</author>                              <date>03/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static StringCollection onAssociatedItems_getComponentClassNamesBackwardCompatibilty
(
IECClass componentClass,
bool stripPrefix
)
    {
    StringCollection assocItemClassNames = new StringCollection ();

    // get the custom attribute from given ecinstance
    IECInstance ecInstanceCustomAttrib = componentClass.GetCustomAttributes ("ASSOCIATED_ITEM_LIST_POINTER");
    if (null == ecInstanceCustomAttrib)
        return null;

    string associatedClassListName = BCT.ECPropertyValueUtility.GetString (ecInstanceCustomAttrib, "ASSOCIATED_ITEM_CLASS", null);
    if (string.IsNullOrEmpty (associatedClassListName))
        return null;

    BmfUt.DisplayListSchemaUtilities listUtilities = new BmfUt.DisplayListSchemaUtilities ();
    ArrayList itemClassNames = listUtilities.GetArrayListDisplay (associatedClassListName, listUtilities.StringListName);
    if (null == itemClassNames)
        return null;

    for (int index = 0; index < itemClassNames.Count; index++)
        {
        string className = itemClassNames[index].ToString ();
        if (stripPrefix)
            assocItemClassNames.Add (BmfCat.CatalogServer.StripClassName (className));
        else
            assocItemClassNames.Add (className);
        }

    return assocItemClassNames;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// initialize associated items utility
/// </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    initializeAssociatedItemsUtility
(
)
    {
    string primarySchemaName = string.Empty;
    BmfCat.Catalog primaryCatalog = null;
    if (ConfigurationVariableUtility.GetValue (out primarySchemaName, "OP_MODEL_PRIMARY_SCHEMA"))
        primaryCatalog = BmfCat.CatalogServer.GetCatalogByName (primarySchemaName);
    else
        {
        // fall back, although OP_MODEL_PRIMARY_SCHEMA is defined in the application level cfg and required for CE
        BmfCat.CatalogCollection projectCatalogs = BmfCat.CatalogServer.ProjectCatalogs;
        if ((null != projectCatalogs) && (projectCatalogs.Count > 0))
            primaryCatalog = projectCatalogs[0]; // TODO making assumption primary is first in list
        }
  
    if (null != primaryCatalog)
        {
        string baseClassName = string.Empty;
        ConfigurationVariableUtility.GetValue (out baseClassName, "OP_ASSOCIATED_ITEM_BASE_CLASS_NAME");
        BCT.AssociatedItemsSpec.OnInitializeBackwardCompatibilty += new BCT.InitializeBackwardCompatibilty (onAssociatedItems_initializeBackwardCompatibilty);
        BCT.AssociatedItemsSpec.Initialize (primaryCatalog.CatalogSchema, baseClassName);
        BCT.AssociatedItemsSpec.OnInitializeBackwardCompatibilty -= new BCT.InitializeBackwardCompatibilty (onAssociatedItems_initializeBackwardCompatibilty);

        BCT.AssociatedItemsSpec.OnGetComponentClassNamesBackwardCompatibilty += new BCT.GetComponentClassNamesBackwardCompatibilty (onAssociatedItems_getComponentClassNamesBackwardCompatibilty);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// initialize logging
/// </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    InitializeLogger
(
)
    {
    System.Reflection.Assembly thisAssembly = System.Reflection.Assembly.GetExecutingAssembly ();
    System.Reflection.AssemblyName assemName = thisAssembly.GetName ();

    string loggerXml = string.Empty;
    ConfigurationVariableUtility.GetValue (out loggerXml, "OP_FRAMEWORKLOGGER_CONFIG");
    new CTLog.FrameworkLogger (loggerXml, assemName.Name);

    CTLog.FrameworkLogger.Instance.Logger.InteractiveMode = true;
    CTLog.FrameworkLogger.Instance.Logger.LoggingEnabled = true;
    CTLog.FrameworkLogger.Instance.Logger.LogToConsoleEnabled = true;

    string logLevel;
    if (ConfigurationVariableUtility.GetValue (out logLevel, PIDConstants.PIDLogLevel))
        CTLog.FrameworkLogger.Instance.Logger.SeverityLevel = (BsiLog.SEVERITY)System.Convert.ToInt32(logLevel);
    else
        CTLog.FrameworkLogger.Instance.Logger.SeverityLevel = BsiLog.SEVERITY.WARNING;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// initialize logging
/// </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    InitializeSpecManager
(
)
    {
    try
        {
        displayTime ("InitializeSpecManager", true);
        if(!s_MicroStationMode)
            Bentley.UI.Controls.WinForms.WaitForm.ShowForm (CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingPIDCaption"),
            CTLog.FrameworkLogger.Instance.GetLocalizedString ("InitializingSpecManager"));

        // DN 02-28-10: Check to see if the configuration variable BMF_SPEC_MANAGER is set to "OPENPLANT"
        // If so, use the new common specs SpecManager; otherwise, use the existing SpecManager
        string specManager = null;
        ConfigurationVariableUtility.GetValue (out specManager, "BMF_SPEC_MANAGER");
        if ((specManager != null) && (String.Compare (specManager, "OPENPLANT", true) == 0))
            {
            // Set the ActiveSpecManager to be the OpenPlant SpecManager
            BmfPS.SpecSystem.ActiveSpecManager = BmfOpenPlant.SpecManager.Instance;

            // Initialize the OpenPlant Spec Manager
            BmfOpenPlant.SpecManager.Instance.Initialize (Bentley.Plant.App.Pid.OpenPlantSpecManagerPlugin.SpecManagerPlugin.Instance,
                s_behavioralSchemaName);
            }
        else
            {
            // Set the ActiveSpecManager to be the Old PID SpecManager
            BmfPS.SpecSystem.ActiveSpecManager = BmfPS.SpecManager.Instance;

            //register spec provider
            BmfPS.SpecManager.Instance.RegisterSpecProvider (s_behavioralSchemaName);
            }
        }
    catch (System.Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    finally
        {
        Bentley.UI.Controls.WinForms.WaitForm.CloseForm ();
        displayTime ("InitializeSpecManager", false);
        }

    }

#endregion

/*------------------------------------------------------------------------------------**/
/// <summary>
/// On Empty Placement Queue callback function
/// </summary>
/// <author>Diana.Fisher</author>                               <date>03/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              OnEmptyPlacementQueue
(
)
    {
    Insert (s_insertUnparsedParameter);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Initialize pid key-in command.  Stops previous command and clears Insert
/// empty placement queue event.
/// </summary>
/// <remarks>Should be called by all non-settings commands</remarks>
/// <author>Diana.Fisher</author>                               <date>10/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             initializePidKeyinCommand
(
)
    {
    Bom.Workspace.InitializeApplicationKeyinCommand ();
    }

#region Methods for Pid MicroStation commands
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Insert a pid component
/// </summary>
/// <remarks>keyin = pid insert</remarks>
/// <param name="unparsed">
/// command line string,
/// should be schema name and class name
/// </param>
/// <author>Diana.Fisher</author>                               <date></date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              Insert
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        // all other commands also call initializePidKeyinCommand (), why not here also?

        // parse the parameters
        string[] insertParams = new string[1];
        insertParams[0] = string.Empty;
        insertParams = unparsed.Split (' ');

        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.TRACE, "Key-in: pid insert " + unparsed);

        if (insertParams.Length == 2)
            {
            // schema name & class name
            Bom.IComponent comp = BmfCat.CatalogServer.CreateComponent (insertParams[0], insertParams[1]);
            if (null != comp)
                {
                s_insertUnparsedParameter = unparsed;

                // TODO - remove the following when overriden values from comp are being used
                if (comp is SchInt.IPageConnector)
                    Bom.Workspace.ActiveModel.InsertComponent (comp, new Bom.OnEmptyPlacementQueueFunc (SchApp.PageConnectorPlacement.OnEmptyPlacementQueue));
                else if (comp.IsRepeatInsertEnabled)
                    Bom.Workspace.ActiveModel.InsertComponent (comp, new Bom.OnEmptyPlacementQueueFunc (OnEmptyPlacementQueue));
                else
                    Bom.Workspace.ActiveModel.InsertComponent (comp, null);
                }
            }
        else
            CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, CTLog.FrameworkLogger.Instance.GetLocalizedString ("InsertRequiresTwoParametersSchemaNameAndClassName"));
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Run the Assembly Creation Wizard
/// </summary>
/// <remarks>keyin = pid assembly create</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateAssembly
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowAssemblyCreationDialog ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Insert an assembly as an assembly
/// </summary>
/// <remarks>keyin = pid assembly insert asassembly</remarks>
/// <param name="unparsed">command line string, full file name of assembly</param>
/// <author>Diana.Fisher</author>                               <date>11/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              InsertAssemblyAsAssembly
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.InsertAssembly (unparsed, Bom.AssemblyInsertionModes.Assembly);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Insert an assembly as individual components
/// </summary>
/// <remarks>keyin = pid assembly insert asindividual</remarks>
/// <param name="unparsed">command line string, full file name of assembly</param>
/// <author>Diana.Fisher</author>                               <date>11/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              InsertAssemblyAsIndividualComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.InsertAssembly (unparsed, Bom.AssemblyInsertionModes.IndividualComponents);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Run the Component Manager
/// </summary>
/// <remarks>keyin = pid component build manage</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <Author> Usman.Nisar </Author>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RunComponentBuilderConsole
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowComponentManagerConsole ();
        }
    catch (Exception e)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, e);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Run the Assembly Manager
/// </summary>
/// <remarks>keyin = pid assembly manage</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RunAssemblyManager
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowAssemblyManagerDialog ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Run component replacment tool
/// </summary>
/// <remarks>keyin = pid component replace</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReplaceComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowComponentReplacementTool ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Run the Component Builder
/// </summary>
/// <remarks>keyin = pid component build wizard</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RunComponentBuilderWizard
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowComponentBuilderDialog ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Redraw all components
/// </summary>
/// <remarks>Keyin = pid component redraw all</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RedrawAllComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RedrawComponents (Bom.ComponentsRedrawMode.All);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Automatic Cell Conformance
/// </summary>
/// <remarks>Converts all cells on the drawing to user components. Conversion is done
/// using conversion rules provided in conversion schema</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ConformCellsAutomatically
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.ConformCellsAutomatically ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Launches Bentley Dashboard
/// </summary>
/// <remarks>Logs a warning in message center if dashboard is not installed, otherwise
/// invokes the dashboard application</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>09/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              LaunchDashboard
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.LaunchBentleyDashboard ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Display Automatic Cell Conformance Mapping Editor
/// </summary>
/// <remarks>User can create a mapping schema using this editor to be used for automatic
/// cell conformance </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowEditorConformCellsAutomatically
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        AutomaticCellConformanceMappingEditor conformanceEditor = new AutomaticCellConformanceMappingEditor ();
        conformanceEditor.ShowDialog ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Redraw selected components
/// </summary>
/// <remarks>Keyin = pid component redraw selected</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RedrawSelectedComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RedrawComponents (Bom.ComponentsRedrawMode.Selected);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Redraw cell based components
/// </summary>
/// <remarks>Keyin = pid component redraw cell</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RedrawCellComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RedrawComponents (Bom.ComponentsRedrawMode.CellBased);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update cell based components to their place scale factor and redraw them
/// </summary>
/// <remarks>
/// Keyin = pid component redraw cells all
/// Special hidden command for Buhler
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateCellsToPlaceScaleFactorAndRedrawAll
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.UpdateToPlaceScaleFactorAndRedrawCellComponents (false);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update cell based components to their place scale factor and redraw them
/// </summary>
/// <remarks>
/// Keyin = pid component redraw cells selected
/// Special hidden command for Buhler
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateCellsToPlaceScaleFactorAndRedrawSelected
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.UpdateToPlaceScaleFactorAndRedrawCellComponents (true);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Dump the selected components
/// Keyin = pid component dump
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DumpSelectedComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        string fileName = BmfUt.MiscellaneousUtilities.GetTempPathAndFileName (PIDConstants.EcInstanceDumpFile);
        Bom.Workspace.ActiveModel.DumpSelectedComponent (fileName);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Release Code Value for Selected Components
/// Keyin = pid component releaseCodeValue
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReleaseCodeSelectedComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ReleaseCodeSelectedComponent();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Release Code Value of Current DGN
/// Keyin = pid model releaseCodeValue
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReleaseCodeCurrentDGN
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ReleaseCodeCurrentDGN();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ReSet IsReserved Flag for the Model and All the Componnets contained in the DGN
/// Keyin = pid model resetReservedFlag
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReSetReservedFlag
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ReSetReservedFlag();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Regenerate GUID for the Current Model and Set IsReserved flag to False for DGN and All 
/// compnents within the active dgn model.
/// </summary>
/// <author>Qarib.Kazmi</author>                             <date>6/2021</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RegenerateGUIDandResetReserveTag
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RegenerateGUIDandResetReserveTag();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>Oppids the adp issue message.</summary>
/// <param name="message">The message.</param>
/// <param name="ApplicationExceptions">The application exceptions.</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void oppidAdapter_IssueMessage(string message, List<ApplicationException> ApplicationExceptions)
{

}

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Ensure all objects that should have an informational relationship with the
/// document have it
/// </summary>
/// <remarks>
/// keyin = pid model fixDocumentRelationships
/// unknown how, but a customer had drawings with the Document_Is_Related_To_Object
/// relationship on the target component, but missing from the source document
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              FixDocumentRelationships
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.FixDocumentRelationships ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show selected components connect points
/// </summary>
/// <param name="unparsed">command line string, "0" to hide; otherwise shows</param>
/// <author>Diana.Fisher</author>                               <date>09/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DisplaySelectedComponentsConnectPoints
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        bool showConnectPoints = true;
        if (unparsed == "0")
            showConnectPoints = false;
        Bom.Workspace.ActiveModel.DisplaySelectedComponentsConnectPoints (showConnectPoints);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show dialog to set Clear ECProperties
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Qarib.Kazmi</author>                               <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ClearECProperties
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowClearECPropertiesDialog ();
        s_msApp.ActiveModelReference.UnselectAllElements ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show selected components connect points
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowSelectedComponentsConnectPoints
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.DisplaySelectedComponentsConnectPoints (true);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Hide selected components connect points
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              HideSelectedComponentsConnectPoints
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.DisplaySelectedComponentsConnectPoints (false);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Place a cloud around a selected component.
/// </summary>
/// <remarks>keyin = pid component cloud add</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AddCloudToComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.AddCloudToSelectedComponent ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Shows RSS Feed for Bentley OpenPLANT PowerPID
/// </summary>
///
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>07/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowRSSFeed
(
string unparsed
)
    {
    //try
    //    {
    //    if (!isApplicationInitialized)
    //        return;

    //    initializePidKeyinCommand ();

    //    if (null == s_RSSFeedForm)
    //        {
    //        string appDataPath = Environment.GetFolderPath (Environment.SpecialFolder.ApplicationData);
    //        string appDataPowerPIDPath = Path.Combine (appDataPath, "Bentley\\PowerPID\\RSSFeed\\");
    //        string feedCaption = CTLog.FrameworkLogger.Instance.GetLocalizedString ("PidApplication_RSSFeedCaption");
    //        s_RSSFeedForm = new BCT.OpenPlantRSSFeed.RSSFeedDisplayContainer
    //            ("http://feeds.rapidfeeds.com/75437/", 
    //            "OpenPlant PID", 
    //            feedCaption, 
    //            appDataPowerPIDPath);

    //        s_RSSFeedForm.FormClosed +=new FormClosedEventHandler (s_RSSFeedForm_FormClosed);
    //        }
    //    s_RSSFeedForm.Show ();

    //    }
    //catch (Exception ex)
    //    {
    //    CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
    //    }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reads reports.ecinstance.xml and creates the same hierarchy of report definitions in 
/// dgnlib.
/// </summary>
///
/// <param name="unparsed">optional name of file, uses reports.ecinstance.xml by default</param>
/// <note> Assumes that the reports.ecinstance.xml is places at path specified by 
/// configuration variable. </note>
/// <remarks>keyin = pid xmlReport toDgnLib</remarks>
/// <author>Nayab.Javed</author>                             <date>09/2017</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              XmlReportToDgnLib
(
string unparsed
)
    {
    MsDgn.CellLibrary activeLib = null;
    if (s_msApp.IsCellLibraryAttached)
        activeLib = s_msApp.AttachedCellLibrary;
    Bentley.DgnPlatformNET.DgnFile templateDgnLib = null;

    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        // Opens reports.ecinstance.xml if no file name given in command
        string report_file_name = "reports.ecinstance.xml";
        if (unparsed.Length > 0)
            report_file_name = unparsed;

        //Getting dgLib and getting write access
        string dgnLibPath = "";
        ConfigurationVariableUtility.GetValue (out dgnLibPath, "BMF_COMPONENT_MANAGER_DGNLIB");
        IntPtr dgnLib_nPtr = MsUtils.MiscellaneousUtilities.GetOpenedFileByName (dgnLibPath.Substring (0, dgnLibPath.LastIndexOf ('\\') + 1) + "PIDProjectTemplates.dgnlib");
        templateDgnLib = Bentley.DgnPlatformNET.DgnFile.GetDgnFile (dgnLib_nPtr);
        bool wasReadOnly = templateDgnLib.IsReadOnly;
        Session.Instance.MakeDgnlibWriteable (ref wasReadOnly, templateDgnLib);
        templateDgnLib.SetTransactable (true);

        //opeing reports xml file
        string reportFilePath = "";
        ConfigurationVariableUtility.GetValue (out reportFilePath, "BMF_PROJECT_REPORTS");
        string reportInstanceFile = reportFilePath + report_file_name;
        string parentCategory = "";
        string currentCategory = "";

        string report_schema_prefix = "{ECREPORT.01.00}";
        string query_schema_prefix = "{http://www.bentley.com/schemas/Bentley.ECQuery.1.0}";

        XElement root = XElement.Load (reportInstanceFile);

        IEnumerable<XElement> reportCategory = root.Elements (report_schema_prefix + "REPORT_CATEGORY");
        IEnumerable<XElement> categoryHasReport = root.Elements (report_schema_prefix + "CATEGORY_HAS_REPORT");
        IEnumerable<XElement> Report = root.Elements (report_schema_prefix + "REPORT");
        IEnumerable<XElement> DgnReportTemplate = root.Elements (report_schema_prefix + "DGN_REPORT_TEMPATE");
        IEnumerable<XElement> ReportHasTemplate = root.Elements (report_schema_prefix + "REPORT_HAS_TEMPLATE");

        //Making a scanCriteria used later for getting column names from cells defined in dgnTemplates
        MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
        scanCriteria.ExcludeNonGraphical ();
        scanCriteria.ExcludeAllTypes ();
        scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.Text);

        foreach (XElement el in reportCategory)
            {
            currentCategory = el.Element (report_schema_prefix + "DISPLAY_LABEL").Value;

            if (currentCategory.Equals ("Reports"))
                continue;

            if (!el.Element (report_schema_prefix + "BASE_CATEGORY_ID").Value.Equals (""))
                {
                string value = el.Element (report_schema_prefix + "BASE_CATEGORY_ID").Value;
                XElement els = (from ele in root.Elements ()
                                where ele.Attribute ("instanceID") != null && ele.Attribute ("instanceID").Value.Equals (value)
                                select ele).FirstOrDefault ();
                parentCategory = els.Element (report_schema_prefix + "DISPLAY_LABEL").Value;
                }
            else
                parentCategory = "";

            // Call to make a category node
            MsUtils.MiscellaneousUtilities.CreateCategoryInDgn (dgnLib_nPtr, currentCategory, parentCategory);

            string ID = el.Attribute ("instanceID").Value;
            IEnumerable<XElement> report_list = (from element in categoryHasReport
                                                    where element.Attribute ("sourceInstanceID").Value.Equals (ID)
                                                    select element);

            string includePropertiesCsv = "";

            foreach (XElement report in report_list)
                {
                string report_id = report.Attribute ("targetInstanceID").Value;
                XElement reportNode = (from element in Report
                                        where element.Attribute ("instanceID").Value.Equals (report_id)
                                        select element).FirstOrDefault ();

                string reportName = reportNode.Element (report_schema_prefix + "NAME").Value;

                string queryText = reportNode.Element (report_schema_prefix + "ECQUERY_DEFINITION").Value;
                XDocument queryDoc = new XDocument ();
                queryDoc = XDocument.Parse (queryText);
                IEnumerable<XElement> query_elements = queryDoc.Root.Elements ();
                XElement searchClassElement = (from element in query_elements
                                                where element.Name == query_schema_prefix + "SearchClass"
                                                select element).FirstOrDefault ();

                List<string> columnClassList = new List<string> ();
                List<string> primaryClassList = new List<string> ();
                List<string> propertyList = new List<string> ();
                List<string> relationshipList = new List<string> ();
                string whereCriteriaRowFilters = "";
                string ecSchema = searchClassElement.Attribute ("ecSchema").Value;
                IECSchema ecSchemaObj = Bentley.Plant.MF.Catalog.CatalogServer.GetCatalogByName (ecSchema).CatalogSchema;
                string ecClass = searchClassElement.Attribute ("ecClass").Value;
                columnClassList.Add (ecClass);
                primaryClassList.Add (ecSchema + ":" + ecClass);

                string isPolymorphic = searchClassElement.Attribute ("isPolymorphic").Value;

                string dgnTemplate_id = (from element in ReportHasTemplate
                                            where element.Attribute ("sourceInstanceID").Value.Equals (report_id) && element.Attribute ("targetClass").Value.Equals ("DGN_REPORT_TEMPATE")
                                            select element).FirstOrDefault ().Attribute ("targetInstanceID").Value;

                XElement dgnTemplateELement = (from element in DgnReportTemplate
                                                where element.Attribute ("instanceID").Value.Equals (dgnTemplate_id)
                                                select element).FirstOrDefault ();

                string use_cell_tag = dgnTemplateELement.Element (report_schema_prefix + "USE_CELL_TAG").Value;
                IList<string> templateColumnNames = new List<string> ();
                if (use_cell_tag.ToLower ().Equals ("true"))
                    {
                    string dgn_cell_library = dgnTemplateELement.Element (report_schema_prefix + "CELL_LIBRARY").Value;
                    string dgn_header_cell = dgnTemplateELement.Element (report_schema_prefix + "DETAIL_CELL").Value;

                    string fullCellLibraryName = "";
                    BmfGs.GFXSystem.ActiveHost.CellLibraryManager.AttachCellLibrary (dgn_cell_library, MsDgn.MsdConversionMode.Prompt, out fullCellLibraryName);
                    bool containsCells = BmfMsCad.ElementConversion.CellLibraryManager.CellLibraryContainsCells (fullCellLibraryName);

                    if (containsCells)
                        {
                        IntPtr cell_nPtr = MsUtils.MiscellaneousUtilities.GetOpenedFileByName (fullCellLibraryName);
                        Bentley.DgnPlatformNET.DgnFile cellDgnFile = Bentley.DgnPlatformNET.DgnFile.GetDgnFile (cell_nPtr);
                        Bentley.DgnPlatformNET.StatusInt cellFileLoaded;
                        cellDgnFile.LoadDgnFile (out cellFileLoaded);

                        if (cellFileLoaded == 0)
                            {
                            Bentley.DgnPlatformNET.ModelId cellModelId = cellDgnFile.FindModelIdByName (dgn_header_cell);
                            Bentley.DgnPlatformNET.StatusInt errorDetails;
                            Bentley.DgnPlatformNET.DgnModel cellModel = cellDgnFile.LoadRootModelById (out errorDetails, cellModelId);
                            cellModel.FillSections (DPNet.DgnModelSections.GraphicElements);
                            DPNet.DgnModelSections isFilled = cellModel.IsFilled (DPNet.DgnModelSections.GraphicElements);

                            if (isFilled == DPNet.DgnModelSections.GraphicElements)
                                {
                                Bentley.DgnPlatformNET.DgnModelRef cellModelRef = cellDgnFile.LoadRootModelById (out errorDetails, cellModelId);
                                MsDgn.ModelReference workReference = BmfUt.MiscellaneousUtilities.GetModelReferenceFromIntPtr (cellModelRef.GetNative ());
                                MsDgn.ElementEnumerator elemEnum = workReference.Scan (scanCriteria);

                                while (elemEnum.MoveNext ())
                                    {

                                    MsDgn.Element curElement = elemEnum.Current;
                                    if (curElement.Type != Bentley.Interop.MicroStationDGN.MsdElementType.Text)
                                        {
                                        continue;
                                        }
                                    MsDgn.TextElement txtElement = curElement.AsTextElement ();
                                    string text = txtElement.Text;
                                    text = text.Substring (1);
                                    templateColumnNames.Add (text.ToUpper ());
                                    }
                                }
                            }
                        }
                    includePropertiesCsv = string.Join (",", templateColumnNames);

                    }
                // if use cell tag was false or the cell file did not provide column names then get column name from ECQuery
                if (includePropertiesCsv.Length == 0 || use_cell_tag.ToLower ().Equals ("false"))
                    {
                    XElement abstractCriteria = query_elements.ElementAt (1).Element (query_schema_prefix + "AbstractSelectCriteria");

                    if (abstractCriteria.Attribute ("selectAllProperties").Value.ToLower ().Equals ("false"))
                        {
                        IEnumerable<string> includeProperties = from ecProp in abstractCriteria.Descendants (query_schema_prefix + "ECPropertyReference")
                                                                select ecProp.Attribute ("ecProperty").Value;
                        includePropertiesCsv = string.Join (",", includeProperties);
                        }
                    else
                        includePropertiesCsv = "";
                    }

                propertyList.Add (includePropertiesCsv);
                //get cloumns for related classses defined in query
                if (use_cell_tag.ToLower ().Equals ("false"))
                    {
                    XElement selectCriteria = query_elements.ElementAt (1);
                    IEnumerable<XElement> related = selectCriteria.Elements (query_schema_prefix + "RelatedInstanceSelectCriteria");

                    foreach (XElement relatedInstance in related)
                        {
                        IEnumerable<XElement> related_properties = relatedInstance.Descendants (query_schema_prefix + "ECPropertyReference");

                        IEnumerable<string> related_classes = (from rel in related_properties
                                                                select rel.Attribute ("ecClass").Value).Distinct ();

                        if (relatedInstance.Element (query_schema_prefix + "SelectCriteria").Element (query_schema_prefix + "AbstractSelectCriteria").Attribute ("selectAllProperties").Value.ToLower ().Equals ("true"))
                            {

                            foreach (string rel in related_classes)
                                {
                                propertyList.Add ("");
                                columnClassList.Add (rel);
                                }
                            }
                        else
                            {
                            foreach (string rel in related_classes)
                                {
                                IEnumerable<string> ec_properties = (from prop in related_properties
                                                                        where prop.Attribute ("ecClass").Value.Equals (rel)
                                                                        select prop.Attribute ("ecProperty").Value);
                                propertyList.Add (string.Join (",", ec_properties));
                                columnClassList.Add (rel);
                                }
                            }

                        IEnumerable<XElement> queryRelatedClassSpecifier = relatedInstance.Descendants (query_schema_prefix + "QueryRelatedClassSpecifier");

                        foreach (XElement queryRelatedClass in queryRelatedClassSpecifier)
                            {
                            string relationClassSchema = queryRelatedClass.Element (query_schema_prefix + "RelationshipClassReference").Attribute ("ecSchema").Value;
                            string relationshipName = queryRelatedClass.Element (query_schema_prefix + "RelationshipClassReference").Attribute ("ecClass").Value;
                            string relatedClassSchema = queryRelatedClass.Element (query_schema_prefix + "RelatedClassReference").Attribute ("ecSchema").Value;
                            string relatedClassName = queryRelatedClass.Element (query_schema_prefix + "RelatedClassReference").Attribute ("ecClass").Value;
                            string relatedDirection = queryRelatedClass.Element (query_schema_prefix + "RelatedInstanceDirection").Attribute ("direction").Value;

                            IECSchema primarySchema = BmfCat.CatalogServer.GetCatalogByName (relationClassSchema).CatalogSchema;

                            IECRelationshipClass relationshipClass = primarySchema.GetClass (relationshipName) as IECRelationshipClass;
                            // platform reporting mechanism has no support for n to 1 relationships, [October 2017]
                            // reversing the relationships is the temporary fix for this limitation
                            // to reverse swap the role of primaryclass and realted class, and reverse the direction

                            if (relationshipClass.Source.Cardinality.UpperLimit == 1 &&
                                relationshipClass.Target.Cardinality.IsUpperLimitUnbounded ())
                                {
                                //reverse the direction of relationship
                                if (relatedDirection.Equals ("Forward"))
                                    relatedDirection = "Backward";
                                else
                                    relatedDirection = "Forward";

                                relationshipList.Add (relationClassSchema + ":" + relationshipName + "," + ecSchema + ":" + ecClass + "," + relatedDirection);

                                // removing the original primaryClass
                                primaryClassList.Remove (ecSchema + ":" + ecClass);
                                // adding related-class into primary-classes list
                                primaryClassList.Add (relatedClassSchema + ":" + relatedClassName);
                                }
                            else if (relationshipClass.Source.Cardinality.IsUpperLimitUnbounded () &&
                                    relationshipClass.Target.Cardinality.IsUpperLimitUnbounded ())
                                {
                                // platform reporting mechanism has no support for n to n relationships, check 
                                // as of October 2017 check if current versions have support
                                continue;
                                }
                            else
                                relationshipList.Add (relationClassSchema + ":" + relationshipName + "," + relatedClassSchema + ":" + relatedClassName + "," + relatedDirection);
                            }
                        }

                    //Extracting date for where criteria, for now only works for EQ operator on String Data
                    XElement whereCriteria = query_elements.ElementAt (2);
                    IEnumerable<XElement> expressions = whereCriteria.Descendants (query_schema_prefix + "Expression");
                    List<string> expressionData = new List<string> ();
                    foreach (XElement expression in expressions)
                        {
                        if (expression.Element ("RelationalOperator").Value.Equals ("EQ") && expression.Element (query_schema_prefix + "RightSideObject").Element ("{http://www.bentley.com/schemas/Bentley.ECSerializable.1.0}" + "BuiltInType").Attribute ("typeCode").Value.Equals ("String"))
                            {
                            if (expressionData.Count > 0)
                                expressionData.Add ((expression.PreviousNode as XElement).Value);

                            string str = "System.String.CompareI(this.";
                            str = str + ecSchema
                                        + "::" + expression.Element (query_schema_prefix + "LeftSideObject").Element (query_schema_prefix + "ECPropertyReference").Attribute ("ecClass").Value
                                        + "::" + expression.Element (query_schema_prefix + "LeftSideObject").Element (query_schema_prefix + "ECPropertyReference").Attribute ("ecProperty").Value
                                        + ", \""
                                        + expression.Element (query_schema_prefix + "RightSideObject").Element ("{http://www.bentley.com/schemas/Bentley.ECSerializable.1.0}" + "BuiltInType").Value
                                        + "\")=true";
                            expressionData.Add (str);
                            }
                        }
                    whereCriteriaRowFilters = string.Join (" ", expressionData);
                    }

                // call to make a report definition node under the current category
                MsUtils.MiscellaneousUtilities.CreateReportDefinitionInDgn (dgnLib_nPtr, reportName, currentCategory, ecSchema, ecSchemaObj.VersionMajor, ecSchemaObj.VersionMinor, primaryClassList, columnClassList, isPolymorphic, propertyList, relationshipList, whereCriteriaRowFilters);
                }
            }

        // once done call process changes to persist the changes made to dgnlib
        MsUtils.MiscellaneousUtilities.CallProcessChanges (dgnLib_nPtr);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    finally
        {
        if (templateDgnLib != null)
            {
            bool wasReadOnly = templateDgnLib.IsReadOnly;
            Session.Instance.MakeDgnlibReadOnly (templateDgnLib, wasReadOnly);
            }

        if (activeLib != null)
            {
            // re attach the cell library that was active before this process started
            s_msApp.AttachCellLibrary (activeLib.FullName, MsDgn.MsdConversionMode.Prompt);
            }
        }
    }
//ends xmlReportToDgnLib


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Closes static instance of RSS Feed form.
/// </summary>
/// <param name="e"></param>
/// <param name="sender"></param>
/// <author>Ali.Aslam</author>                             <date>07/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
//static void                     s_RSSFeedForm_FormClosed
//(
//object sender, 
//FormClosedEventArgs e
//)
//    {
//    s_RSSFeedForm = null;
//    }

//static Bentley.Plant.CommonTools.OpenPlantRSSFeed.RSSFeedDisplayContainer s_RSSFeedForm;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Remove a cloud around a selected component.
/// </summary>
/// <remarks>keyin = pid component cloud remove</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RemoveCloudFromComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RemoveCloudFromComponent ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Assign a template to selected component(s) and optionally set their design state.
/// </summary>
/// <remarks>keyin = pid component assigntemplate</remarks>
/// <param name="unparsed">
/// command line string,
/// first part: template name; ex "equip\equipment"
/// optional second part: "INVALID", "VALID", or "NOTYETVALIDATED"
/// </param>
/// <author>Dustin.Parkman</author>                             <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AssignTemplateToComponent
(
string unparsed
)
    {
    char separator = ' ';
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        string[] options = unparsed.Split (separator);
        if (options.GetUpperBound (0) == 0)
            Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponent (unparsed);

        if (options.GetUpperBound (0) == 1)
            {
            string state = options[1].ToUpper ();
            switch (state)
                {
                case "INVALID":
                        {
                        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (options[0], 2);
                        break;
                        }
                case "VALID":
                        {
                        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (options[0], 1);
                        break;
                        }
                case "NOTYETVALIDATED":
                        {
                        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (options[0], 0);
                        break;
                        }
                }
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Assign a template to selected component(s), if their design state is invalid
/// </summary>
/// <remarks>keyin = pid component assigntemplate invalid</remarks>
/// <param name="unparsed">command line string, template name; ex "equip\equipment"</param>
/// <author>Diana.Fisher</author>                             <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AssignTemplateToInvalidComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (unparsed, 2);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Assign a template to selected component(s), if their design state is valid
/// </summary>
/// <remarks>keyin = pid component assigntemplate valid</remarks>
/// <param name="unparsed">command line string, template name; ex "equip\equipment"</param>
/// <author>Diana.Fisher</author>                             <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AssignTemplateToValidComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (unparsed, 1);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Assign a template to selected component(s), if their design state is not yet
/// validated.
/// </summary>
/// <remarks>keyin = pid component assigntemplate notYetValidated</remarks>
/// <param name="unparsed">command line string, template name; ex "equip\equipment"</param>
/// <author>Diana.Fisher</author>                             <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AssignTemplateToNotYetValidatedComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetElementTemplateForSelectedComponentsByValidationState (unparsed, 0);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Iterates over each component within the model and calls it's Validate method.
/// Each components Validate method will add it to the InvalidComonents collection,
/// if the ValidationState is InValid.
/// Optionally shows form
/// </summary>
/// <remarks>
/// keyin = pid model validate
/// if Model Validation form was already opened, and no parameter.  The results will
/// display in form.
/// </remarks>
/// <param name="unparsed">
/// command line string,
/// if "TRUE" or "1" also shows model validation form
/// </param>
/// <author>Dustin.Parkman</author>                             <date>10/2007</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ValidateModel
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.Validate ();
        if (!string.IsNullOrEmpty (unparsed) &&
            (unparsed.Equals ("TRUE", StringComparison.OrdinalIgnoreCase) ||
             unparsed.Equals ("1")))
            {
            BmfVal.ModelValidationForm modelValidationform = BmfVal.ModelValidationForm.GetInstance (Bom.Workspace.ActiveModel, (Bom.IWorkspace)BmfAi.AppInterface.Workspace, false);
            modelValidationform.Show ();
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Iterates over each component within the model and calls it's Validate method.
/// Each components Validate method will add it to the InvalidComonents collection,
/// if the ValidationState is InValid.
/// </summary>
/// <remarks>
/// keyin = pid model validate showResults
/// if Model Validation form was already opened.  Another form will be displayed.
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ValidateModelAndShowForm
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.Validate ();
        BmfVal.ModelValidationForm modelValidationform = BmfVal.ModelValidationForm.GetInstance (Bom.Workspace.ActiveModel, (Bom.IWorkspace)BmfAi.AppInterface.Workspace, false);
        modelValidationform.Show ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Iterates over each component within the model and validates it against the
/// alternate repository.
/// </summary>
/// <remarks>keyin = pid model validateAlternateRepository</remarks>
/// <param name="unparsed">
/// command line string,
/// if "TRUE" or "1" shows model validation form
/// </param>
/// <author>Shoaib.Ehsan</author>                               <date>09/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ValidateModelWithAlternateRepository
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        if (!string.IsNullOrEmpty (unparsed) &&
            (unparsed.Equals ("TRUE", StringComparison.OrdinalIgnoreCase) ||
             unparsed.Equals ("1")))
            {
            BmfVal.ModelValidationForm modelValidationform = BmfVal.ModelValidationForm.GetInstance (Bom.Workspace.ActiveModel, (Bom.IWorkspace)BmfAi.AppInterface.Workspace, true);
            modelValidationform.Show ();
            }
        //else
        //    Bom.Workspace.ActiveModel.ValidateWithAlternateRepository();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Creates list of logical groups which do not have a logical related component
/// </summary>
/// <remarks>keyin = pid model delete CorruptComponents</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                               <date>09/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteOrphanedLogicalGroups
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.DeleteOrphanedLogicalGroups ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Find and delete corrupt components in current drawing
/// </summary>
/// <remarks>keyin = pid model delete CorruptComponents</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>11/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteCorruptComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        int corruptCount = Bom.Workspace.ActiveModel.FindCorruptComponents ();
        if (corruptCount > 0)
            {
            BmfVal.ModelCorruptionForm corruptionForm = BmfVal.ModelCorruptionForm.GetInstance (Bom.Workspace.ActiveModel,
                (Bom.IWorkspace)BmfAi.AppInterface.Workspace);
            corruptionForm.Show ();
            }
        else if (0 == corruptCount)
            {
            CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.INFO,
                CTLog.FrameworkLogger.Instance.GetLocalizedString ("CorruptComponents"),
                CTLog.FrameworkLogger.Instance.GetLocalizedString ("NoCorruptComponents"),
                    CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ClearOrphanedPBSItems
/// </summary>
/// <remarks>keyin = 
/// pid model delete ClearOrphanedPBSItems
/// pid model delete ClearOrphanedPBSItems PLANT_AREA
/// pid model delete ClearOrphanedPBSItems PLANT_AREA,SERVICE,UNIT
/// </remarks>
/// <param name="unparsed">The PBS Class</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ClearOrphanedPBSItems
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();
        string[] pbsClasses = unparsed.Split(',');
        if (string.IsNullOrEmpty(unparsed))
            pbsClasses = null;

        if (Bom.Workspace.ActiveModel.ClearOrphanedPBSItems(pbsClasses))
            {
                RefreshStandardPreferencesDialog(null, null);
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Resets the standard preferences dialog.
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <remarks>pid model ResetStandardPreferencesDialog</remarks>
/// <param name="unparsed">The unparsed.</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ResetStandardPreferencesDialog
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();
        RefreshStandardPreferencesDialog(null, null);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Identify the components which are not unique within the plant database and
/// assign next available (unique) business key to such components
/// </summary>
/// <remarks>keyin = pid PlantProject assignUniqueKeys</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                               <date>11/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AssignNextAvailableBusinessKey
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.AssignNextAvailableBusinessKey ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refreshes local briefcase file from imodel server 
/// </summary>
/// <remarks>keyin = pid PlantProject refreshlocaliModelcopy</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                               <date>09/2017</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RefreshLocaliModelCopy
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.RefreshLocaliModelCopyFromServer ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reissues the local iTwin briefcase
/// Prompts user to reissue, if yes, the local briefcase is deleted and the oppid is restarted
/// The next time OPPID is started the local briefcase is downloaded
/// </summary>
/// <remarks>keyin = pid PlantProject ReIssueiTwinBriefcase</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReIssueiTwinBriefcase
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        if (Bom.Workspace.ActiveModel.ReIssueiTwinBriefcase())
            forceExit();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reuse project 
/// </summary>
/// <remarks>keyin = pid PlantProject reuse</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>09/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ReuseProject
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfAi.AppInterface.Workspace.Reuse ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Returns components in model after filtering out PID_ASSEMBLY, PID_DOCUMENT and
/// BORDER objects
/// </summary>
/// <returns> filtered list of components in model</returns>
/// <author>Quratulain.Shaheen</author>                               <date>01/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static Bom.IComponentCollection getFilteredComponentsForExport
(
string modelName
)
    {
    Bom.IComponentCollection filterComponents = new Bom.ComponentCollection ();
    Bom.IComponentCollection components = null;
    if (string.IsNullOrEmpty(modelName))
        components = Bom.Workspace.ActiveModel.FindAllComponentsInActiveModel();
    else
        components = Bom.Workspace.ActiveModel.FindAllComponentsInSpecificModelByClassName (modelName, "BMF_COMPONENT");

    if (null != components)
        {
        foreach (Bom.IBaseObject currentObject in components)
            {
            //Exclude any BMF_ASSEMBLY, PID_DOCUMENT and PIDBorder classes
            if (!(currentObject is Bom.IAssemblyComponent || currentObject.ClassDefinition.Name.Equals ("PID_DOCUMENT") || currentObject is PidInt.IPidBorder))
                filterComponents.Add (currentObject);
            }
        }
    return filterComponents;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Export the active model to drawing.
/// </summary>
/// <remarks>keyin = pid model export drawing</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>02/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ExportModelToDrawing
(
string unparsed
)
    {
    if (!isApplicationInitialized)
        return;

    initializePidKeyinCommand ();

    Bom.IComponentCollection filterComponents = getFilteredComponentsForExport (unparsed);
    Bom.IComponentCollection filterEmbeddedComponents = Bom.Workspace.ActiveModel.FilterOutEmbeddedTargets (filterComponents);
    Bom.Workspace.ActiveModel.ShowExportOptionsDialog (Bom.ImportExportMode.Drawing, filterEmbeddedComponents);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Export the active model to drawing.
/// </summary>
/// <remarks>keyin = pid model export ECXML</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>02/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ExportModelToECXML
(
string unparsed
)
    {
    if (!isApplicationInitialized)
        return;

    initializePidKeyinCommand ();

    string filterOnModelName = "";
    if (!string.IsNullOrEmpty(unparsed))
        filterOnModelName = unparsed;

    Bom.IComponentCollection filterComponents = getFilteredComponentsForExport (filterOnModelName);
    Bom.Workspace.ActiveModel.ShowExportOptionsDialog (Bom.ImportExportMode.ECXML, filterComponents);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Import into the active model.
/// </summary>
/// <remarks>keyin = pid model import drawing</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ed.Becnel</author>                                   <date>05/2008</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ImportModel
(
string unparsed
)
    {
    if (!isApplicationInitialized)
        return;

    initializePidKeyinCommand ();

    Bom.Workspace.ActiveModel.ShowImportOptionsDialog ();
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show Start Up Settings
/// </summary>
/// <remarks>keyin = pid model settings</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowStartUpSettings
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();  // should this stop Insert?

        if (s_initialized)
            forceShowStartUp (false, 0);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Registry state for whether or not to show Standard Preferences dialog
/// </summary>
/// <author>Diana.Fisher</author>                               <date>02/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              StateShowStandardPreferencesDialog
    {
    get
        {
        int val = (int)BmfUt.StartupRegistryUtilities.GetRegistryValue (m_constShowAssociateItems, 1);
        if (val == 1)
            return true;

        return false;
        }
    set
        {
        if (value) //show
            BmfUt.StartupRegistryUtilities.SetRegistryValue (m_constShowAssociateItems, 1);
        else
            BmfUt.StartupRegistryUtilities.SetRegistryValue (m_constShowAssociateItems, 0);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show Standard Preferences dialog
/// </summary>
/// <author>Diana.Fisher</author>                               <date>01/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowStandardPreferencesDialog
(
)
    {
    if (null == m_standardPreferencesDialog)
        {
        LicenseManager.MarkFeature (PidFeature.StandardPrefrencesDialog);
        BmfMsCad.IBmfMsHost msHost = BmfGs.GFXSystem.ActiveHost as BmfMsCad.IBmfMsHost;
        if (msHost != null)
            m_standardPreferencesDialog = new StandardPreferencesDialog (Bom.Workspace.ActiveModel, msHost.MatchLifetime);
        else
            m_standardPreferencesDialog = new StandardPreferencesDialog (Bom.Workspace.ActiveModel);

        m_standardPreferencesDialog.Show ();
        m_standardPreferencesDialog.Activate ();
        StateShowStandardPreferencesDialog = true;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Close Standard Preferences dialog if it is open
/// </summary>
/// <author>Diana.Fisher</author>                               <date>01/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CloseStandardPreferencesDialog
(
)
    {
    if (null != m_standardPreferencesDialog)
        m_standardPreferencesDialog.CloseAdapter ();
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets m_standardPreferencesDialog to null
/// </summary>
/// <author>Diana.Fisher</author>                               <date>01/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetStandardPreferencesDialogToNull
(
)
    {
    m_standardPreferencesDialog = null;
    // if dialog is closed due to model closing, should not set flag
    // this ensures next open model follows the last user setting of flag
    if (!s_closeOccurring)
        StateShowStandardPreferencesDialog = false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refresh given category and optionally specific property of Standard Preferences dialog, if it is open
/// </summary>
/// <author>Diana.Fisher</author>                               <date>01/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RefreshStandardPreferencesDialog
(
string category,
string propertyName
)
    {
    if (null == m_standardPreferencesDialog)
        return;

    // if category not specified, refresh entire dialog
    if (string.IsNullOrEmpty (category))
        {
        //m_standardPreferencesDialog.Refresh (); does not work
        CloseStandardPreferencesDialog ();
        ShowStandardPreferencesDialog ();
        }
    else
        {
            // TODO: need mechanism to either
            //          refresh all properties 
            //      or 
            //          only a specific category property 
            //  within Standard Preferences dialog
            // until then, refresh entire dialog
        CloseStandardPreferencesDialog ();
        ShowStandardPreferencesDialog ();
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refresh Standard Preferences dialog if it is open
/// </summary>
/// <author>Diana.Fisher</author>                               <date>01/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RefreshStandardPreferencesDialog
(
)
    {
    RefreshStandardPreferencesDialog (null, null);
    }

/*------------------------------------------------------------------------------------**/
/// <summary> Registry access key for show associated items</summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
private const string            m_constShowAssociateItems = "ShowAssociatedItems";

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Show standard preferences dialog
/// </summary>
/// <remarks>keyin = pid model standardPreferences</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowStandardPreferencesDialog
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();  // should this stop Insert?

        ShowStandardPreferencesDialog ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refreshes the standard preferences dialog.
/// </summary>
/// <param name="unparsed">The unparsed.</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RefreshStandardPreferencesDlg
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand();

        string[] insertParams = unparsed.Split(' ');
        if(insertParams.Count() < 3)
            return;
        string WBSClassName = insertParams[0];
        string businessKeyNew = insertParams[1];
        string businessKeyOld = insertParams[2];
        m_standardPreferencesDialog.RefreshWBSItemsList(WBSClassName, businessKeyNew, businessKeyOld); 
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Create Instrument line(s) from dumb lines and/or linestrings
/// </summary>
/// <remarks>keyin = pid conform instrumentline</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                              <date>11/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateUserInstrumentLine
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.CreateUserInstrumentLines ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Create Pipe Run(s) from dumb lines and/or linestrings
/// </summary>
/// <remarks>keyin = pid piperun create</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                              <date>09/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateUserRun
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.CreateUserPipeRuns ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Splits Run into two runs using DataChange like behavior.
/// </summary>
/// <remarks>keyin = pid piperun split</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SplitRun
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient runSplitTool = new SchApp.RunSplitTool ();
        runSplitTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Merges two runs into a single run using DataChange like behavior.
/// </summary>
/// <remarks>keyin = pid piperun merge</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              MergeRun
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient runMergeTool = new SchApp.RunMergeTool ();
        runMergeTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates To From information for all pipe runs in model
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>09/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             updateToFromForAllRuns
(
)
    {
    Bom.IComponentCollection allComponents = Bom.Workspace.ActiveModel.FindAllComponentsInActiveModelByClassName ("RUN");
    if (null == allComponents)
        return;

    string completionBarText = CTLog.FrameworkLogger.Instance.GetLocalizedString("ToFromUpdate_ProgressMessage");
    MsUtils.CompletionBar completionBar = new MsUtils.CompletionBar();

    try
        {
        completionBar.OpenCompletionBar(completionBarText, allComponents.Count);   

        foreach (Bom.IComponent currentComponent in allComponents)
            {
            completionBar.IncrementCompletionBar (completionBarText);

            if (currentComponent is SchInt.ISchematicsRun)
                {
                SchInt.ISchematicsRun currentRun = (SchInt.ISchematicsRun)currentComponent;
                currentRun.UpdateRunWithConnectedCompData ();
                BmfPer.PersistenceManager.Update (currentComponent);
                }
            }
        }
    catch
        {
        }
    finally
        {
        completionBar.CloseCompletionBar ();
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates To From information for all pipe runs in model
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                             <date>09/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateToFrom
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        updateToFromForAllRuns ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Change an existing Pipe Run from one pipeline to another
/// </summary>
/// <remarks>keyin = pid piperun ChangePipeline</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ChangeRunPipelineTool
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient changeRunPipelineTool = new ChangeRunPipelineTool (null);
        changeRunPipelineTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Evaluate the run to add jumpers over components.
/// </summary>
/// <remarks>keyin = pid piperun JumpComponents</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>03/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              JumpComponentsForRun
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.IComponentCollection selectedComponents = Bom.Workspace.ActiveModel.SelectedComponents;
        if (selectedComponents == null)
            return;

        // ReevaluateJumpers will force the selection to update, by removing and adding back the run
        // therefore must maintain new collection to avoid exception with modified collection
        SchInt.IRunCollection runs = new SchCat.RunCollection ();
        foreach (Bom.IComponent currentComponent in selectedComponents)
            {
            if (currentComponent is SchInt.ISchematicsRun)
                runs.Add ((SchInt.ISchematicsRun)currentComponent);
            }

        foreach (SchInt.ISchematicsRun currentRun in runs)
            currentRun.ReevaluateJumpers ();

        // in future might do something like RunMergeTool that allows selection of runs
        //BmfCad.ICadElementEditorClient runJumpComponentsTool = new SchApp.RunJumpComponentsTool ();
        //RunJumpComponentsTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Evaluate the run to add jumpers over components.
/// </summary>
/// <remarks>keyin = pid piperun ValidateBreaks</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>11/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ValidateBreaksForRun
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.IComponentCollection selectedComponents = Bom.Workspace.ActiveModel.SelectedComponents;
        if (selectedComponents == null)
            return;

        // ValidateBreaks will force the selection to update, by removing and adding back the run
        // therefore must maintain new collection to avoid exception with modified collection
        SchInt.IRunCollection runs = new SchCat.RunCollection ();
        foreach (Bom.IComponent currentComponent in selectedComponents)
            {
            if (currentComponent is SchInt.ISchematicsRun)
                runs.Add ((SchInt.ISchematicsRun)currentComponent);
            }

        foreach (SchInt.ISchematicsRun currentRun in runs)
            currentRun.ValidateBreaks ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set the run interior vertex handle behavior to slide
/// </summary>
/// <remarks>keyin = pid piperun InteriorVertexHandle add</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>09/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetRunInteriorVertexHandleToAdd
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        //initializePidKeyinCommand ();  // should this stop Insert?

        SchApp.SchematicsApplication.RunInteriorVertexEditHandleSlides = false;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set the run interior vertex handle behavior to slide
/// </summary>
/// <remarks>keyin = pid piperun InteriorVertexHandle slide</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>09/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetRunInteriorVertexHandleToSlide
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        //initializePidKeyinCommand ();  // should this stop Insert?

        SchApp.SchematicsApplication.RunInteriorVertexEditHandleSlides = true;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Adds an existing Component to an existing Loop
/// </summary>
/// <remarks>keyin = pid loop AddComponent</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AddComponentToLoopTool
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient addComponentToLoopTool = new AddComponentToLoopTool (null);
        addComponentToLoopTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Removes a component from an existing Loop
/// </summary>
/// <remarks>keyin = pid loop RemoveComponent</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RemoveComponentFromLoopTool
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient removeComponentFromLoopTool = new RemoveComponentFromLoopTool (null);
        removeComponentFromLoopTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Creates a relationship between two objects
/// </summary>
/// <remarks>keyin = pid relationship create</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateRelationship
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient createRelationshipTool = new CreateRelationshipTool (null);
        createRelationshipTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Removes a relationship between two desiginated components.
/// </summary>
/// <remarks>keyin = pid relationship remove</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RemoveRelationship
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient removeRelationshipTool = new RemoveRelationshipTool (null);
        removeRelationshipTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Joins MicroStation elements to an existing component
/// </summary>
/// <remarks>keyin = pid join</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Dustin.Parkman</author>                             <date>09/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              Join
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient joinTool = new JoinTool (null);
        joinTool.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Removes join of MS elements from a component that has MS elements previously joined.
/// </summary>
/// <remarks>keyin = pid join remove</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Asif.Khan</author>                             <date>10/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RemoveJoin
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        RemoveJoinList rjl = RemoveJoinList.GetInstance (Bom.Workspace.ActiveModel, (Bom.IWorkspace)BmfAi.AppInterface.Workspace);
        rjl.Show ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Displays the Base Component Annotation dialog.  This dialog lists all component
/// annotation classes that derive from a base component annotation class.
/// </summary>
/// <remarks>if unparsed is empty, default base class is pid:COMPONENT_ANNOTATION</remarks>
/// <param name="unparsed">command line string, optional base class for component annotation list</param>
/// <author>Asif.Khan</author>                             <date>11/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AnnotateComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        string[] insertParams = unparsed.Split (' ');

        if (1 != insertParams.Length)
            return;

        String baseClass = PIDConstants.BasePidComponentAnnotationClass;
        if (String.Empty != insertParams[0])
            baseClass = insertParams[0];

        using (PidAnnotationListDialog annotDlg = new PidAnnotationListDialog(baseClass))
            {
            if (DialogResult.OK == annotDlg.ShowDialog())
                {
                PidInt.IPidComponentAnnotation annotation = annotDlg.AnnotationComponent;
                if (null != annotation)
                    Bom.Workspace.ActiveModel.InsertComponent(annotation, null);
                }
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Creates a component from selected native elements
/// pid component create [LIQUID_PUMPS] ANNO
/// pid component create [LIQUID_PUMPS]
/// pid component create [EQUIPMENT]
/// pid component create CONE_ROOF_TANK
/// pid component create
/// </summary>
/// <remarks>keyin = pid component create</remarks>
/// <param name="unparsed">command line string, this can contain class names and annotation flag</param>
/// <author>Steve.Morrow</author>                               <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateUserComponent
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        //TODO this needs to be made into a class to support selecting components, like ChangeRunPipelineTool
        PIDUtilities.CreateUserComponent (unparsed);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Diagnostic method to check if relationship blocks exist on both source and target
/// components for all instances in drawing
/// </summary>
/// <remarks>This method only detects issues and logs them in message center.
/// Not fix them.
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ValidateRelationshipsInModel
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.ValidateRelationshipsInModel ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Diagnostic method to ensure all schematics runs have CP1 and CP2 connect points
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>09/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              FixRunConnectPoints
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand();

        SchApp.SchematicsApplication.VerifyRunConnectPoints ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException(BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Diagnostic method to move the origin and other component points, 
/// to the location of the ms element graphics.
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Qarib.Kazmi</author>                                <date>05/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              FixComponentPoints
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand();

        BmfVal.AdjustedComponentPointsForm adjustedForm = BmfVal.AdjustedComponentPointsForm.GetInstance (Bom.Workspace.ActiveModel,
            (Bom.IWorkspace)BmfAi.AppInterface.Workspace);

        adjustedForm.ValidateModel ();
        int invalidCount = adjustedForm.InvalidComponents.Count;
        if (invalidCount > 0)
            {
            adjustedForm.Show ();
            adjustedForm.PopulateDataGrid ();
            }
        else
            {
            CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.INFO,
                CTLog.FrameworkLogger.Instance.GetLocalizedString ("AdjustedComponents"),
                CTLog.FrameworkLogger.Instance.GetLocalizedString ("NoAdjustedComponents"),
                    CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException(BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// A wrapper key in for hidden key in of items browser Items FixCorruptedRelationships
/// </summary>
/// <remarks> This method is only a wrapper and doesn't do any additional work</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              FixCorruptRelationships
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfGs.GFXSystem.ActiveHost.UiUtils.ExecuteKeyIn ("Items FixCorruptedRelationships");
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Redraws all text/annotation elements in file with their default/configured
/// text style. Also refreshes text style on elements if definition in xml hasn't changed
/// </summary>
/// <remarks> This is useful if user wants to change the text style of specific types
/// of components by updating the default text style against their ECClass in schema.
/// The method would change text style for every component whose text style is different
/// from the one specified in the ECSchema.</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>02/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RedrawAllAnnotationComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.IAnnotationCollection annotationsRequiringApplicationProcessing;

        Bom.Workspace.ActiveModel.RedrawAnnotationComponents (Bom.ComponentsRedrawMode.All,
            out annotationsRequiringApplicationProcessing);

        PIDUtilities.RedrawAnnotationActuators (annotationsRequiringApplicationProcessing);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Redraws selected text/annotation elements in file with their default/configured
/// text style.
/// </summary>
/// <remarks> This is useful if user wants to change the text style of specific types
/// of components by updating the default text style against their ECClass in schema.
/// The method would change text style for every component whose text style is different
/// from the one specified in the ECSchema.</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
///
/// <author>Ali.Aslam</author>                                  <date>02/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RedrawSelectedAnnotationComponents
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.IAnnotationCollection annotationsRequiringApplicationProcessing;
        Bom.Workspace.ActiveModel.RedrawAnnotationComponents (Bom.ComponentsRedrawMode.Selected,
            out annotationsRequiringApplicationProcessing);

        PIDUtilities.RedrawAnnotationActuators (annotationsRequiringApplicationProcessing);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Provides user with various component selection options
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectBy
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();  // should this stop insert?

        ComponentSelectionDialog selectionDialog = new ComponentSelectionDialog ();
        selectionDialog.Show ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Selects a component on drawing with given persistent element path
/// </summary>
/// <remarks> Would only select if corresponding element is graphical</remarks>
/// <param name="unparsed">persistent element path of element to be selected</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectByPEP
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.SelectByPEP (unparsed);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Selects a component on drawing with given global id
/// </summary>
/// <param name="unparsed">global id in registry format</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectByGUID
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.SelectByGuid (unparsed);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Selects a component on drawing with given element id
/// </summary>
/// <param name="unparsed">element id in registry format</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectByElementId
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        int elementId = -1;
        if (int.TryParse (unparsed, out elementId))
            PIDUtilities.SelectByElementID (elementId);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Selects a component on drawing with given instance id
/// </summary>
/// <param name="unparsed">command line string, this can contain full or partial
/// instance id</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectByInstanceId
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        PIDUtilities.SelectByInstanceId (unparsed);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Diagnostic method to detect duplicate GUID values in drawing
/// </summary>
/// <remarks>Instances with duplicate guids would be logged as error in message center
/// </remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>06/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CheckForDuplicateGUIDs
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        //TODO this needs to be made into a class to support selecting components, like ChangeRunPipelineTool
        PIDUtilities.CheckForDuplicateGUIDs ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Check page connectors.  Ensure linked connectors have required data
/// </summary>
/// <remarks>must be done post upgrade, because UpdateInstanceUsingWorkDGN only works with CE instances</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>9/2018</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CheckSsSixPageConnectors
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        SchApp.PageConnectorUtility.VerifyUpgradedLinks ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates GUID property of selected components with new GUID values.
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Ali.Aslam</author>                                  <date>7/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateComponentGUID
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.UpdateComponentGUID ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Batch convert
/// </summary>
/// <remarks>keyin = pid batch convert</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Steve.Morrow</author>                             <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              BatchConvert
(
string unparsed
)
    {
    string savedSuppressStartupDialog = "";
    bool savedInteractiveMode = CTLog.FrameworkLogger.Instance.Logger.InteractiveMode;
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        ConfigurationVariableUtility.GetValue (out savedSuppressStartupDialog, "PID_SUPPRESS_STARTUP_DIALOG");
        // Supress startup dialog for this session (do not persist)
        ConfigurationVariableUtility.Add ("PID_SUPPRESS_STARTUP_DIALOG", "1", false);
        // Ensure batch mode is enabled.
        CTLog.FrameworkLogger.Instance.Logger.InteractiveMode = false;

        Bentley.DgnPlatformNET.DgnDocumentList filesToConvert = null;

        if (string.IsNullOrEmpty (unparsed))
            filesToConvert = BmfAi.AppInterface.Workspace.RunBatchConversion ();
        else
            filesToConvert = readAndCreateBatchFileList (unparsed);

        if (null == filesToConvert || filesToConvert.Count == 0)
            return;

        Bentley.Plant.MF.Conversion.ConversionUtilities.InBatch = true;
        SetLogModeBatch (string.Empty);

       // using (filesToConvert)
            {
            int numSelected = filesToConvert.Count;
            for (int entryIndex = 0; entryIndex < numSelected; entryIndex++)
                {
                using (DPNet.DgnDocument document = filesToConvert[entryIndex])
                    {
                    string fileName = document.FileName ;
                    bool isDMSDocument = document.DocState == DPNet.DgnDocument.State.InDMS;//document.IsDMSDocument ();// Todo: AliA_Vancouver confirm
                    if (isDMSDocument)
                        document.Fetch (DPNet.DgnDocument.FetchMode.Write, DPNet.DgnDocument.FetchOptions.Default);
                    BmfGs.GFXSystem.ActiveHost.CadDocumentServer.OpenDrawingDocument (fileName, "DEFAULT", false);
                    Bom.Workspace.ActiveModel.OpenForBatchConvert ();
                    BmfAi.AppInterface.Workspace.SaveModel (fileName, BmfCad.FileSaveReason.SaveAtClose);
                    BmfPer.PersistenceManager.CloseModel ();
                    if (isDMSDocument)
                        document.Put(DPNet.DgnDocument.PutAction.Checkin, DPNet.DgnDocument.PutOptions.Silent,"Checking in batch converted document");
                    }
                }
            }

        Bentley.Plant.MF.Conversion.ConversionUtilities.BatchComplete ();
        SetLogModeInteractive (string.Empty);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    finally
        {
        // Restore previous startup dialog value for this session (do not persist)
        ConfigurationVariableUtility.Add ("PID_SUPPRESS_STARTUP_DIALOG", savedSuppressStartupDialog, false);
        // Restore previous value of interactive mode
        CTLog.FrameworkLogger.Instance.Logger.InteractiveMode = savedInteractiveMode;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reads the and create batch file list.
/// </summary>
/// <param name="fileName">Name of the file.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                   <date>3/10/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/
private static DPNet.DgnDocumentList readAndCreateBatchFileList
(
string fileNameParameter
)
    {
    string fileName = "";
    DPNet.DgnDocument inputParameterDocument = null;
    DPNet.DgnDocumentList files = new DPNet.DgnDocumentList();

    try
        {
        using (DPNet.DgnDocumentMoniker monikerFromFileInputParameter =
                         DPNet.DgnDocumentMoniker.CreateFromFileName(fileNameParameter, ""))
            {
            if (null != monikerFromFileInputParameter)
                {
                    DPNet.StatusInt status;
                inputParameterDocument =
                                  DPNet.DgnDocument.CreateFromMoniker(out status,
                                          monikerFromFileInputParameter, 0,
                                         DPNet.DgnDocument.FetchMode.Read, DPNet.DgnDocument.FetchOptions.Default);
                if (null != inputParameterDocument)
                    {
                    if (inputParameterDocument.DocState == DPNet.DgnDocument.State.InDMS)//.IsDMSDocument ())// Todo:AliA_Vancouver confirm
                        inputParameterDocument.Fetch (DPNet.DgnDocument.FetchMode.Read, DPNet.DgnDocument.FetchOptions.Default);
                    fileName = inputParameterDocument.FileName ;
                    }
                }
            if (!File.Exists (fileName))
                {
                CTLog.FrameworkLogger.Instance.SubmitResponse
                    (BsiLog.SEVERITY.LOG_FATAL,
                    CTLog.FrameworkLogger.Instance.GetLocalizedString ("BatchConversion"),
                    string.Format (CTLog.FrameworkLogger.Instance.GetLocalizedString (
                                   "BatchConvertParameterFileFind"), fileName),
                                   CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
                return null;
                }


            using (StreamReader reader = new StreamReader (fileName))
                {
                string line = string.Empty;
                while ((line = reader.ReadLine ()) != null)
                    {
                    if (!string.IsNullOrEmpty (line))
                        {
                        using (DPNet.DgnDocumentMoniker monikerFromFile =
                                 DPNet.DgnDocumentMoniker.CreateFromFileName(line, ""))
                            {
                            DPNet.StatusInt status;
                                DPNet.DgnDocument documentManagerDocument =
                              DPNet.DgnDocument.CreateFromMoniker(out status,
                                      monikerFromFile, 0,
          DPNet.DgnDocument.FetchMode.Write, DPNet.DgnDocument.FetchOptions.Default );
                            if (null == documentManagerDocument)
                                CTLog.FrameworkLogger.Instance.SubmitFindFile (BsiLog.SEVERITY.LOG_WARNING, line);
                            else
                                files.Add (documentManagerDocument);
                            }
                        }
                    }
                }
            }
        }
    catch (Exception e)
        {
        CTLog.FrameworkLogger.Instance.SubmitResponse
            (BsiLog.SEVERITY.LOG_FATAL,
            CTLog.FrameworkLogger.Instance.GetLocalizedString ("BatchConversion"), e.Message,
            CTLog.MessageBoxButtons.OK,
            CTLog.DialogResponse.OK);
        return null;
        }
    finally
        {
        if (null != inputParameterDocument)
            inputParameterDocument.Dispose ();
        }

    return files;

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Synchronizes with alternate repository. Shows dialog based on config variable
/// </summary>
/// <remarks>keyin = pid PlantProject sync</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                             <date>03/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SynchronizeWithAlternateRepository
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        if (Bom.Workspace.InDwg && Bom.Workspace.IsActiveModelSheet)
            return;

        Bom.Workspace.ActiveModel.SynchronizeWithAlternateRepository ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refreshes active model and update title sheets after a successful sync
/// if sync direction is from alternate repository to dgn.
/// Caller should ensure that this method only gets called when direction is alternate
/// repostitory to dgn.
/// </summary>
/// <author>Ali.Aslam</author>                               <date>06/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             refreshActiveModelAndUpdateTitleSheets
(
)
    {
    BmfAi.AppInterface.Workspace.RefreshActiveModel ();
    if (!Bom.Workspace.ActiveModel.ModelName.Equals ("unknown", StringComparison.InvariantCultureIgnoreCase))
        {
        Bom.IComponentCollection borders = Bom.Workspace.ActiveModel.ExistingBorders;
        if (null == borders)
            return;

        foreach (Bom.IComponent borderComponent in borders)
            {
            borderComponent.UpdateGraphics ();
            if (null != borderComponent.GraphicsElement)
                Bom.Workspace.ActiveModel.ManipulateCompleteUpdateMicroStationTags (borderComponent.GraphicsElement);
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Disassociates the drawing with the associated Plant Project
/// </summary>
/// <remarks>keyin = pid PlantProject disassociate</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                               <date>06/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DisAssociateWithAlternateRepository
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.DisAssociateWithAlternateRepository ();
        string bodymsg = CTLog.FrameworkLogger.Instance.GetLocalizedString ("DisAssociateWithAlternateRepositoryBody");
        string caption = CTLog.FrameworkLogger.Instance.GetLocalizedString ("DisAssociateWithAlternateRepositoryCaption");
        CTLog.FrameworkLogger.Instance.SubmitResponse (BsiLog.SEVERITY.INFO, caption,
            bodymsg, CTLog.MessageBoxButtons.OK, CTLog.DialogResponse.OK);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Toggle between the work offline and online states with alternate repository
/// </summary>
/// <remarks>keyin = pid PlantProject workoffline</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                             <date>03/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetWorkOffline
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetWorkOfflineMode ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set logging mode to batch
/// </summary>
/// <remarks>keyin = pid logMode batch</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>07/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetLogModeBatch
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();  // should this stop insert?

        CTLog.FrameworkLogger.Instance.Logger.InteractiveMode = false;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set logging mode to interactive
/// </summary>
/// <remarks>keyin = pid logMode interactive</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Diana.Fisher</author>                               <date>07/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetLogModeInteractive
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();  // should this stop insert?

        CTLog.FrameworkLogger.Instance.Logger.InteractiveMode = true;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the Differencing dialog option to display it while synchronization or not
/// It toggles the current state of this option
/// </summary>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                             <date>05/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ShowDiffDialogOnSync
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ShowDiffDialogOnSync ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the option to synchronize DGN with Plant Project on file open
/// It toggles the current state of this option
/// </summary>
/// <remarks>keyin = pid PlantProject synconstartup</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/// <author>Shoaib.Ehsan</author>                             <date>05/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizeOnStartUp
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetSynchronizeOnStartUp ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the source and target repository for synchronization i.e. from DGN to SQL or vice versa
/// if user specified:
/// 0 means Always Ask
/// 1 means DGN to DB
/// 2 means DB to DGN
/// </summary>
/// <remarks>
/// keyin = pid PlantProject syncdirection
/// defaults to Always Ask
/// </remarks>
/// <param name="unparsed">command line string, sync direction</param>
/// <author>Shoaib.Ehsan</author>                             <date>05/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizationDirection
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        int syncDirection = 0;
        if (!string.IsNullOrEmpty (unparsed))
            {
            if (!Int32.TryParse (unparsed, out syncDirection))
                syncDirection = 0;

            if ((syncDirection < 0) || (syncDirection > 2))
                syncDirection = 0;
            }

        Bom.Workspace.ActiveModel.SetSynchronizationDirection (syncDirection);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the source and target repository for synchronization to Always Ask
/// </summary>
/// <remarks>keyin = pid PlantProject syncdirection alwaysAsk</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizationDirectionAlwaysAsk
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetSynchronizationDirection (0);  // 0 means always ask
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the source and target repository for synchronization to dgn to db
/// </summary>
/// <remarks>keyin = pid PlantProject syncdirection DgnToDb</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizationDirectionDgnToDb
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetSynchronizationDirection (1);  // 1 means dgn to db
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the source and target repository for synchronization to db to dgn
/// </summary>
/// <remarks>keyin = pid PlantProject syncdirection DbToDgn</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Diana.Fisher</author>                               <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizationDirectionDbToDgn
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetSynchronizationDirection (2);  // 2 means db to dgn
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Launches dialog to specify synchronization settings
/// </summary>
/// <remarks>keyin = pid PlantProject settings</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Shoaib.Ehsan</author>                             <date>06/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetSynchronizationSettings
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.SetSynchronizationSettings ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Open Plant Sight
/// </summary>
/// <remarks>keyin = pid PlantProject OpenPlantSight</remarks>
/// <param name="unparsed">command line string, ignored by this command</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              OpenPlantSight
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

		string plantSight = s_connectApi.GetBuddiUrl("PLANTSIGHT_PORTAL");
        if (string.IsNullOrWhiteSpace(plantSight))
        {
            return;
        }

        string projectGuidWs = MsUtils.ActiveParamsWrapper.GetConnectProjectGUID();
        if (string.IsNullOrWhiteSpace(projectGuidWs))
        {
            System.Diagnostics.Process.Start(plantSight);
            return;
        }

		string plantSightProj = string.Format(@"{0}/{1}/home", plantSight, projectGuidWs);
		System.Diagnostics.Process.Start(plantSightProj);		
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Open Project Portal
/// </summary>
/// <remarks>keyin = pid PlantProject OpenPlantSight</remarks>
/// <param name="unparsed">command line string, used in keyin</param>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              OpenProjectPortal
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();
        string cfgProj = string.Empty;
        ConfigurationVariableUtility.GetValue(out cfgProj, "CONNECTPROJECTGUID");
        if (string.IsNullOrWhiteSpace(cfgProj))
            return;

        string url = getProjectPortalUrl(new System.Guid (cfgProj));
        System.Diagnostics.Process.Start(url);

        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>Gets the project portal URL.</summary>
/// <param name="projectId">The project identifier.</param>
/// <returns>url</returns>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static string getProjectPortalUrl(Guid projectId)
{
    try
    {
        //Gets the url of Project portal from specified ProjectId
        string url = s_connectApi.GetProjectHomepage(projectId.ToString());

        // Replace the url in passive state to bypass the login screen if user is already login
        url = s_connectApi.GetA2PUrl(url);

        return url;
    }
    catch (Exception exp)
    {
        string msg = exp.Message;
    }
    return "";
}


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Assigns each component in the model a new GUID
/// </summary>
/// <remarks>keyin = pid model regenguid</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Dustin.Parkman</author>                             <date>07/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RegenGuid
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bentley.Plant.MFPersist.PersistenceManager.UpdateGlobalIds (true);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Select current OpenPlant Spec Provider (AutoPlant, OpenPlant, PlantSpace, PDX)
/// </summary>
/// <remarks>keyin = pid OpenPlantSpecs selectProvider</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Dan.Nichols</author>                             <date>06/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SelectOpenPlantSpecsProvider
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfOpenPlant.SpecManager.Instance.SelectSpecProvider ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Refresh Properties from key in command
/// </summary>
/// <remarks>keyin = pid model refreshProperties</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Steve.Morrow</author>                                  <date>08/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              RefreshAndRedrawProperties
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        refreshPropertiesAndRedrawOnStartup (true);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates current drawing from instances in given idgn list file
/// </summary>
/// <remarks>keyin = pid model import idgnlist</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Ali.Aslam</author>                                  <date>09/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ImportIDgnList
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ImportIDgnList ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates current drawing from instances in given idgn list file
/// </summary>
/// <remarks>keyin = pid model import owllist</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Ali.Aslam</author>                                  <date>09/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ImportOwlList
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        Bom.Workspace.ActiveModel.ImportOwlList ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Converts drawing to OP dgn and then publishes it as dgn
/// </summary>
/// <remarks>
/// keyin = pid model export publishOpenPlantIDgnToOwl
/// if the file is not an i-model and is  OPPID create i-model and process with
/// OWL writer behind the scenes
/// if the file is an i-model then you just process that file with the
/// OWL writer behind the scenes
/// If not i-model and not OPPID then prompt user that this option is not
/// available for this file type.
/// </remarks>
/// <param name="unparsed">command line string, optional.  if "1" show dialog</param>
/// <author>Ali.Aslam</author>                                  <date>09/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              PublishOpenPlantIDgnToOwl
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadDocument activeDocument = BmfGs.GFXSystem.ActiveCadDocument;
        string filename = activeDocument.FileName;
        string activeFileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension (filename);
        string directoryName = System.IO.Path.GetDirectoryName (filename);
        string outputFileName = directoryName + "\\" + activeFileNameWithoutExtension + ".owl";
        if (filename.ToLower ().EndsWith (".i.dgn"))
            Bom.Workspace.ActiveModel.PublishOpenPlantIDgnToOwl (filename, outputFileName);
        else
            {
            string publishKeyInCommand = string.Empty;
            if (unparsed.Equals ("1"))
                publishKeyInCommand = "publishdgn dialog";
            else
                publishKeyInCommand = "publishdgn publish force";

            BmfGs.GFXSystem.ActiveHost.UiUtils.ExecuteKeyIn (publishKeyInCommand);

            string[] idgnFiles = System.IO.Directory.GetFiles (directoryName, "*.i.dgn");
            foreach (string idgnFile in idgnFiles)
                {
                string idgnFileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension (idgnFile);
                if (idgnFileNameWithoutExtension.ToLower ().CompareTo (activeFileNameWithoutExtension.ToLower () + ".i") == 0 ||
                    idgnFileNameWithoutExtension.ToLower ().CompareTo (activeFileNameWithoutExtension.ToLower () + ".dgn.i") == 0)
                    {
                    Bom.Workspace.ActiveModel.PublishOpenPlantIDgnToOwl (idgnFile, outputFileName);
                    continue;
                    }
                }
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }
#endregion

#region Utility methods

/*------------------------------------------------------------------------------------**/
/// <summary>Pid catalog name </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            PidCatalogName
    {
    get
        {
        return PIDConstants.PidCatalogName;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>Pid Behavior catalog name </summary>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            PidBehaviorCatalogName
    {
    get
        {
        // TODO PidBehaviorCatalogName limits the customer to only 1 behavioral catalog
        return s_behavioralSchemaName;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set the components active scale
/// </summary>
/// <remarks>keyin = pid component scale set</remarks>
/// <param name="unparsed">command line string, double value for scale</param>
/// <author>Steve.Morrow</author>                             <date>08/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetScale
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        //initializePidKeyinCommand ();  // should this stop insert?

        string[] insertParams = new string[1];
        insertParams[0] = string.Empty;
        insertParams = unparsed.Split (' ');

        if (insertParams.Length == 0)
            return;

        double scale = 1.0;
        Double.TryParse (insertParams[0], out scale);

        Bom.Workspace.ActiveModel.CellComponentScaleFactor = scale;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reset the components active scale to 1
/// </summary>
/// <remarks>keyin = pid component scale reset</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Steve.Morrow</author>                             <date>08/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ResetScale
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        //initializePidKeyinCommand ();  // should this stop insert?

        BmfUt.UnitConversionUtilities.ResetOnDemandApplyScaleFactor ();
        Bom.Workspace.ActiveModel.CellComponentScaleFactor = 1.0;
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Change Pipeline or Instrument line Type
/// if no string is passed, Pipeline change is done
/// </summary>
/// <remarks>keyin = pid component ChangeLineType</remarks>
/// <param name="unparsed">command line string, either PIPELINE or INSTRUMENTLINE</param>
/// <author>Steve.Morrow</author>                             <date>09/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ChangeLineType
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        if (unparsed.Length == 0 || unparsed.Equals ("PIPELINE", StringComparison.InvariantCultureIgnoreCase))
            {
            BmfCad.ICadElementEditorClient chgPipelineType = new ChangePipelineType ();
            chgPipelineType.InitDialog ();
            chgPipelineType.StartClient ();
            }
        else if (unparsed.Equals ("INSTRUMENTLINE", StringComparison.InvariantCultureIgnoreCase))
            {
            BmfCad.ICadElementEditorClient chgInstrumentLineType = new ChangeInstrumentLineType ();
            chgInstrumentLineType.InitDialog ();
            chgInstrumentLineType.StartClient ();
            }
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Change Pipeline Type
/// </summary>
/// <remarks>keyin = pid component ChangeLineType Pipeline</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Diana.Fisher</author>                              <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ChangePipelineType
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient chgPipelineType = new ChangePipelineType ();
        chgPipelineType.InitDialog ();
        chgPipelineType.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Change Instrument line Type
/// </summary>
/// <remarks>keyin = pid component ChangeLineType instrumentline</remarks>
/// <param name="unparsed">command line string, ignored by this method</param>
/// <author>Diana.Fisher</author>                              <date>03/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              ChangeInstrumentLineType
(
string unparsed
)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();

        BmfCad.ICadElementEditorClient chgInstrumentLineType = new ChangeInstrumentLineType ();
        chgInstrumentLineType.InitDialog ();
        chgInstrumentLineType.StartClient ();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }


#endregion

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Event Handler for component changed events.  Notifies addins of these changes.
/// </summary>
/// <author>Dustin.Parkman</author>                              <date>07/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private void                    componentChanged
(
object sender,
BmfPer.ComponentChangedEventArgs e
)
    {
    if (s_powerPIDAddins==null)
        return;

    foreach(PidInt.IPowerPIDAddin addin in s_powerPIDAddins)
        {
        switch (e.ComponentChangeType)
            {
            case BmfPer.ComponentChangeType.ComponentAdded:
                addin.ComponentAdded((Bom.IComponent)sender);
                break;
            case BmfPer.ComponentChangeType.ComponentUpdated:
                 addin.ComponentUpdated((Bom.IComponent)sender);
                break;
            case BmfPer.ComponentChangeType.ComponetDeleted:
                 addin.ComponentDeleted((Bom.IComponent)sender);
                break;
            case BmfPer.ComponentChangeType.RelationshipAdded:
                 addin.RelationshipAdded((IECRelationshipInstance)sender);
                break;
            case BmfPer.ComponentChangeType.RelationshipDeleted:
                 addin.RelationshipDeleted((IECRelationshipInstance)sender);
                break;
            }

        }
    }

public static void ImportFromAI (string unparsed)
    {
    try
        {
        if (!isApplicationInitialized)
            return;

        initializePidKeyinCommand ();
        AIImport.Import import = new AIImport.Import();
        import.Run();
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        }
    }


}
}
