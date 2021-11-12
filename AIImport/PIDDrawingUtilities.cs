
/*--------------------------------------------------------------------------------------+
* 
*      $Source: /plantint-root/plantint/bentley.plant/Application/PID/AutoPLANTPIDConversion/APPIDUtilities.cs,v $
*     $RCSfile: APPIDUtilities.cs,v $
*    $Revision: 1.95.2.1.2.4.2.1 $
*        $Date: 2021/04/29 10:25:34 $
*      $Author: Qarib.Kazmi $
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
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;
using Bentley.ECObjects.Instance;
using Bentley.ECObjects.Schema;
using BsiLog = Bentley.Logging;
using Bom = Bentley.Plant.ObjectModel;
using MsDgn = Bentley.Interop.MicroStationDGN;
using BmfGe = Bentley.Plant.MF.Geometry;
using CTLog = Bentley.Plant.CommonTools.Logging;
using BmfPr = Bentley.Plant.MF.Primitives;
using BmfCad = Bentley.Plant.CadSys;
using BmfUt = Bentley.Plant.MF.Utilities;
using BmfCM = Bentley.Plant.MF.ComponentManager;
using BmfCat = Bentley.Plant.MF.Catalog;
using BmfGs = Bentley.Plant.MF.GraphicsSystem;
using BmfPer = Bentley.Plant.MFPersist;
using Bentley.Plant.ConfigVariable;
//using BmfMsCad = Bentley.Plant.MF.CadSystem.MicroStation;
using SchInt = Bentley.Plant.AppFW.Schematics.Interfaces;
using SchCat = Bentley.Plant.AppFW.Schematics.Catalog;
using SchApp = Bentley.Plant.AppFW.Schematics;
using Bentley.Plant.MF.Conversion;
using Bentley.Plant.App.Pid;
using PidInt = Bentley.Plant.App.Pid.Interfaces;
//using BmfCSI = Bentley.Plant.ModelingFramework.CadSystem.CadSystemInterfaces;
using BCT = Bentley.Plant.CommonTools;

namespace Bentley.Plant.App.Pid.AIImport
{
    
/*====================================================================================**/
/// <summary>
/// APPID Utilities
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*==============+===============+===============+===============+===============+=====**/
public class                    APPIDUtilities
    {

#region Properties
/*------------------------------------------------------------------------------------**/
/// <summary>
/// APPID Components
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static Hashtable         APPIDComponents
    {
    get
        {
        return m_APPIDComponents;
        }
    set
        {
        m_APPIDComponents = value;
        }
    }
private static Hashtable m_APPIDComponents = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// APPID ProjectID
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            APPIDProjectID
    {
    get
        {
        return m_APPIDProjectID;
        }
    set
        {
        m_APPIDProjectID = value;
        }
    }
private static string m_APPIDProjectID = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// APPID DocumentID
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            APPIDDocumentID
    {
    get
        {
        return m_APPIDDocumentID;
        }
    set
        {
        m_APPIDDocumentID = value;
        }
    }
private static string m_APPIDDocumentID = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Conversion catalog
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static BmfCat.Catalog    ConversionCatalog
    {
    get
        {
        return m_conversionCatalog;
        }
    set
        {
        m_conversionCatalog = value;
        }
    }
private static BmfCat.Catalog m_conversionCatalog = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Project Schema Name
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            ProjectSchemaName
    {
    get
        {
        return ConversionUtilities.GetMainProjectSchemaName ();
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Metric Scaling Is Being Used
/// <author>Steve.Morrow</author>                                <date>11/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              MetricScalingIsBeingUsed
    {
    get
        {
        return m_metricScalingIsBeingUsed;
        }
    set
        {
        m_metricScalingIsBeingUsed = value;
        }
    }
private static bool m_metricScalingIsBeingUsed = false;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Invalid Or Missing Xdata List
/// <author>Steve.Morrow</author>                                <date>03/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static StringCollection  InvalidOrMissingXdataList
    {
    get
        {
        return m_invalidOrMissingXdataList;
        }
    set
        {
        m_invalidOrMissingXdataList = value;
        }
    }
private static StringCollection m_invalidOrMissingXdataList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Piperun Xlass
/// <author>Steve.Morrow</author>                                <date>12/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static IECClass         s_pipeRunClass
    {
    get
        {
        if (null == m_pipeRunClass)
            m_pipeRunClass = BmfCat.CatalogServer.GetECClass (BmfPer.PersistenceManager.PipingNetworkSegmentName);

        return m_pipeRunClass;
        }
    }
private static IECClass m_pipeRunClass = null;

#endregion

#region MS Application 
/*------------------------------------------------------------------------------------**/
/// <summary>
/// MS Application
/// <author>Steve.Morrow</author>                                <date>12/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static MsDgn.Application MsDgnApplicationInstance
    {
    get
        {
        if (null == m_msDgnApplicationInstance)
            m_msDgnApplicationInstance = BmfCad.CadApplication.Instance.MSDGNApplication;

        return m_msDgnApplicationInstance;
        }
    }
private static MsDgn.Application m_msDgnApplicationInstance = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Pid DisplayList Schema Utilities
/// <author>Steve.Morrow</author>                                <date>02/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static PidDisplayListSchemaUtilities ListUtilities
    {
    get
        {
        if (null == m_listUtil)
            m_listUtil = new Pid.PidDisplayListSchemaUtilities (Pid.PidApplication.PidBehaviorCatalogName, 
                                 Pid.PIDConstants.PidAttributeClass);

        return m_listUtil;
        }
    }
private static PidDisplayListSchemaUtilities m_listUtil = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Pipe sizes list
/// <author>Steve.Morrow</author>                                <date>02/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static IList<KeyValuePair<string, string>> pipeSizes
    {
    get
        {
        if (null == m_sizesList)
            {
            m_sizesList = new List<KeyValuePair<string, string>> ();
            string sizeClassName = BmfUt.UnitConversionUtilities.AlternateActiveSizeListClassName;
            if (null == sizeClassName || sizeClassName.Length == 0)
                {
                sizeClassName = ListUtilities.GetCurrentListMode (PIDConstants.ActiveSizesList);
                if (null == sizeClassName || sizeClassName.Length == 0)
                    sizeClassName = PIDConstants.ImperialPipeSizes;
                }
            m_sizesList = ListUtilities.GetSizeNominalDisplay (sizeClassName);

            if (null == m_sizesList)
                m_sizesList = new List<KeyValuePair<string, string>> ();
            }

        return m_sizesList;
        }
    }
private static IList<KeyValuePair<string, string>> m_sizesList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Adds the cloud.
/// </summary>
/// <param name="apComp">The ap comp.</param>
/// <param name="warningText">Warning text for validation</param>
/// <author>Steve.Morrow</author>                                    <date>1/15/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              AddCloud 
(
APPIDComponent apComp, 
string warningText
)
    {
    try
        {
        Bom.ICloud cloud = Bom.BmfObjectFactory.CreateCloud ();
        if (cloud != null)
            {
            apComp.OPPIDComponent.AddCloud (cloud, true);
            apComp.OPPIDComponent.ValidationState = 2;
            if (!string.IsNullOrEmpty (warningText))
                apComp.OPPIDComponent.ValidationFailureDescription = warningText;
            Bom.Workspace.ActiveModel.UpdateComponent (apComp.OPPIDComponent);
            }
        }
    catch
        {
        }
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Sets the active warning level.
/// </summary>
/// <param name="warningLevelName">Name of the warning level.</param>
/// <param name="color">The color.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                    <date>1/15/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public static MsDgn.Level SetActiveWarningLevel(string warningLevelName, int color)
    {
    try
        {
        if (string.IsNullOrEmpty(warningLevelName))
            return null;

        MsDgn.Level activeLevel = MsDgnApplicationInstance.ActiveSettings.Level;
        MsDgn.Level warningLevel = MsDgnApplicationInstance.ActiveDesignFile.Levels.Find(warningLevelName, null);
        if (warningLevel == null)
            {
            MsDgnApplicationInstance.ActiveDesignFile.AddNewLevel(warningLevelName);
            warningLevel = MsDgnApplicationInstance.ActiveDesignFile.Levels.Find(warningLevelName, null);
            if (warningLevel != null)
                {
                warningLevel.IsDisplayed = true;
                warningLevel.ElementColor = color; //red
                MsDgnApplicationInstance.ActiveDesignFile.Levels.Rewrite();
                }
            }
        if (warningLevel != null)
            MsDgnApplicationInstance.ActiveSettings.Level = warningLevel;
        return activeLevel;
        }
    catch
        {
        return null;
        }
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Resets the warning level.
/// </summary>
/// <param name="activeLevel">The active level.</param>
/// <author>Steve.Morrow</author>                                    <date>1/15/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public static void ResetWarningLevel(MsDgn.Level activeLevel)
    {
    try
        {
        if (activeLevel != null)
            MsDgnApplicationInstance.ActiveSettings.Level = activeLevel;
        }
    catch { }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set Units
/// APPID Metric Dwg are scaled up very large. OPPID needs to have these drawings in real world units.
/// This methods scales the metric drawing down by using a variable that would be set in the project.pcf file
/// Set flag indicating that metric scaling is being used. This is needed for instrument bubbles.
/// Instrument bubbles have their cell origin stored in xdata. This cannot be used becuase of the scaling. Need to use the annotation attribute
/// to get the instrument bubble origin.
/// </summary>
/// <author>Steve.Morrow</author>                            <date>08/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetUnitsAndScaleElements
(
)
    {
    MsDgnApplicationInstance.ActiveSettings.AngleLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.AssociationLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.AxisLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.DepthLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.GraphicGroupLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.GridLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.IsometricLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.LevelLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.ScaleLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.SnapLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.TextNodeLockEnabled = false;
    MsDgnApplicationInstance.ActiveSettings.UnitLockEnabled = false;
    
    // Line weights are disabled by default for drawings coming from APPID. Enable them
    // to give a consistent view with out of box environment.
    // Currently only donig that in open views
    foreach (MsDgn.View thisView in MsDgnApplicationInstance.ActiveDesignFile.Views)
        if (thisView.IsOpen)
            thisView.DisplaysLineWeights = true;

    MsDgn.MeasurementUnit masterUnits = new MsDgn.MeasurementUnit ();
    MsDgn.MeasurementUnit subUnits = new MsDgn.MeasurementUnit ();

    string seedFileName = string.Empty;
    if (ConfigurationVariableUtility.GetValue (out seedFileName, "MS_DWGSEED"))
        {
        BmfGs.GFXSystem.ActiveHost.CadDocumentServer.GetUnitsFromFile (seedFileName, ref masterUnits, ref subUnits);
        MsDgnApplicationInstance.ActiveModelReference.set_MasterUnit (ref masterUnits);
        MsDgnApplicationInstance.ActiveModelReference.set_SubUnit (ref subUnits);
        }

    string globalLineStyleScale = string.Empty;
    if (ConfigurationVariableUtility.GetValue(out globalLineStyleScale, "APPID_CONVERT_GLOBAL_LINESTYLE_SCALE"))
        {
        string keyIn = string.Format ("active linestylescale {0}", globalLineStyleScale);
        BmfGs.GFXSystem.ActiveHost.UiUtils.ExecuteKeyIn (keyIn);
        }

    string scaleValue = string.Empty;
    if (!ConfigurationVariableUtility.GetValue(out scaleValue, "APPID_CONVERT_METRIC_SCALE_VALUE"))
        return;

    //string scaleValue = "0.20";
    string keyin = "choose all;active scale " + scaleValue + ";scale original;xy=0,0;xy=0,0;active scale 1.0;choose none;NULL";
    BmfGs.GFXSystem.ActiveHost.UiUtils.ExecuteKeyIn (keyin);

    BmfGs.GFXSystem.ActiveHost.ViewManager.FitView (0, true);

    //Cfg variable to set grid units
    string gridUnit = string.Empty;
    if (ConfigurationVariableUtility.GetValue(out gridUnit, "APPID_CONVERT_METRIC_GRID_UNIT_VALUE"))
        {
        double gridUnitValue = 0.01500;
        if (Double.TryParse (gridUnit, out gridUnitValue))
            MsDgnApplicationInstance.ActiveSettings.GridUnits = gridUnitValue;
        }

    MetricScalingIsBeingUsed = true;

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update APPID Text Style Fonts To Match OPPID font names
/// </summary>
/// <remarks>
/// This mapping is defined in the mapping schema. Each item is a paired list: 
///     Text Style Name
///     Font Name
/// The Font will replace the existing Font in the the text style
/// 
/// This is needed to ensure the Fonts styles are consistent from APPID to OPPID
/// In DWG drawing, it seems that even though the font is set on the text during conversion, the font
/// displayed in the element info is that of the one of the original existing fonts: STANDARD, ANNOTATIVE...
/// 
/// So to stay consistent, the Font defined on APPID text style "STANDARD" must be the same as the font defined on the OPPID Text style "PID"
/// This goes for all native APPID text style and fonts
/// </remarks> 
/// <author>Steve.Morrow</author>                            <date>02/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateAPPIDTextStyleFontsToMatchOPPID
(
)
    {
    SortedList styleFontList = ConversionMappingUtilities.GetTextStyleFontUpdateMapping ("TEXT_STYLE_FONT_UPDATE_MAPPING", ConversionCatalog);
    if (null == styleFontList || styleFontList.Count == 0)
        return;

    for (int count = 0; count < styleFontList.Count; count++)
        {
        string textStyleName = styleFontList.GetKey (count).ToString ();
        string fontName = styleFontList.GetByIndex (count).ToString ();

        if (string.IsNullOrEmpty (textStyleName) || string.IsNullOrEmpty (fontName))
            continue;

        MsDgn.TextStyle targetStyle = getTextStyleFromName (textStyleName);
        if (null == targetStyle || targetStyle.IsLocked)
            continue;

        for (int i = 1; i <= MsDgnApplicationInstance.ActiveDesignFile.Fonts.Count; i++)
            {
            if (MsDgnApplicationInstance.ActiveDesignFile.Fonts[i].Name.Equals (fontName, StringComparison.InvariantCultureIgnoreCase))
                {
                targetStyle.Font = MsDgnApplicationInstance.ActiveDesignFile.Fonts[i];
                break;
                }
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// getTextStyleFromName
/// </summary>
/// <param name="textStyleName">text style</param> 
/// <author>Steve.Morrow</author>                            <date>02/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static MsDgn.TextStyle  getTextStyleFromName
(
string textStyleName
)
    {
    for (int count = 1; count <= MsDgnApplicationInstance.ActiveDesignFile.TextStyles.Count; count++)
        {
        if (MsDgnApplicationInstance.ActiveDesignFile.TextStyles[count].Name.Equals (textStyleName, StringComparison.InvariantCultureIgnoreCase))
            return MsDgnApplicationInstance.ActiveDesignFile.TextStyles[count];
        }
    return null;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete element
/// </summary>
/// <param name="elementIDs">element Id to remove</param> 
/// <author>Steve.Morrow</author>                            <date>08/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteElement
(
UInt32 id
)
    {

    MsDgn.Element element = ElementFromElementUnit (id);
    if (null != element)
        MsDgnApplicationInstance.ActiveModelReference.RemoveElement (element);

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete elements
/// </summary>
/// <param name="elementIDs">element Ids to remove</param> 
/// <author>Steve.Morrow</author>                            <date>08/2010</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteElements
(
ArrayList elementIDs
)
    {
    if (null == elementIDs || elementIDs.Count == 0)
        return;

    foreach (UInt32 id in elementIDs)
        DeleteElement (id);

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Element from ElementID
/// </summary>
/// <param name="elementID">element ID</param> 
/// <returns>MsDgn.Element</returns>
/// <author>Steve.Morrow</author>                            <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static MsDgn.Element     ElementFromElementID
(
long elementID
)
    {

    if (elementID == 0)
        return null;

    MsDgn.Element element = null;
    // A try catch is needed because the xdata handle from AutoCAD might not always be valid
    try
        {
        element = MsDgnApplicationInstance.ActiveModelReference.GetElementByID (ref elementID);
        }
    catch
        {
        return null;
        }
    return element;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Element from ElementID
/// </summary>
/// <param name="elementID">element ID</param> 
/// <returns>MsDgn.Element</returns>
/// <author>Steve.Morrow</author>                            <date>04/2009</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static MsDgn.Element     ElementFromElementUnit
(
UInt32 elementID
)
    {

    if (elementID == 0)
        return null;

    MsDgn.Element element = null;
    // A try catch is needed because the xdata handle from AutoCAD might not always be valid
    try
        {
        element = MsDgnApplicationInstance.ActiveModelReference.GetElementByID64 (elementID);
        }
    catch
        {
        return null;
        }
    return element;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is Element Text
/// </summary>
/// <param name="id">element Id</param> 
/// <returns>true if element is text</returns> 
/// <author>Steve.Morrow</author>                            <date>08/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsElementText
(
UInt32 id,
out bool embeddedTag
)
    {
    embeddedTag = false;
    MsDgn.Element element = ElementFromElementUnit (id);
    if (null != element)
        {
        if (element.IsTagElement ())
            return true;
        if (element.IsTextElement ())
            return true;
        if (element.IsTextNodeElement ())
            return true;
        if (element.HasAnyTags)
            {
            //this block has embedded attribute text and needed to be converted to native text.
            //this is the case with a couple of the actuator blocks
            //use the IGNORE_EMBEDDED_ATTRIBUTE_BLOCK_NAMES_LIST list to determine this cells
            if (element.IsSharedCellElement () &&
                IgnoreBlockNameWithTagSet (element.AsSharedCellElement ().Name))
                {
                embeddedTag = true;
                return false;
                }

            return true;
            }
        }
    return false;
    }

#endregion

#region Scan methods

/*------------------------------------------------------------------------------------**/
/// <summary>
/// There are several DIN lines in APPID which have special graphic cells in their line 
/// styles which are not linked to their pipe run component element via some xdata. 
/// This method very explicitly tries to find and returns list of such elements for 
/// particular line styles only. It also moves these graphics to a special level it 
/// creates just for these graphics and turns it off.
/// 
/// This method works for shared cell graphics only.
/// Search is cell name and range based which are required
/// Level based search is optional. 
/// Resulting cell must not have any xdata to be considered. 
/// </summary>
/// <param name="apComponent">appid run component</param> 
///
/// <author>Ali.Aslam</author>                                      <date>10/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static List<UInt32>      MoveSpecialDumbComponentsInSomeLinesToDeletedLevel
(
APPIDComponent apPipeRunComponent
)
    {
    List<UInt32> returnList = new List<uint> ();

    if 
    (
    null == apPipeRunComponent ||
    !apPipeRunComponent.IsPipeRun ||
    null == apPipeRunComponent.SourceData ||
    !apPipeRunComponent.SourceData.Contains("PINT1")
    )
    return returnList;

    string pipeRunCode = apPipeRunComponent.SourceData["PINT1"].ToString ().Trim ().ToUpper ();
    string cellNameToSearch = "";
    string levelNameToSearch = "";
    if (pipeRunCode == "DIN-514")
        {
        cellNameToSearch = "dindot";
        levelNameToSearch = "din4";
        }
    else if (pipeRunCode == "DIN-513")
        {
        cellNameToSearch = "dinsfuel";
        levelNameToSearch = "din4";
        }
    else if (pipeRunCode == "DIN-515")
        {
        cellNameToSearch = "dinbar";
        levelNameToSearch = "din4";
        }
    if (string.IsNullOrEmpty (cellNameToSearch))
        return returnList;

    Bentley.Interop.MicroStationDGN.Level levelToSearch = null;

    // Level based search is optional, but if specified, level must exist
    if (levelNameToSearch != "")
        {
        levelToSearch = MsDgnApplicationInstance.ActiveDesignFile.Levels.Find (levelNameToSearch, null);
        if (null == levelToSearch)
            return returnList;
        }

    if (null == apPipeRunComponent.OPPIDComponent ||
        null == apPipeRunComponent.OPPIDComponent.GraphicsElement ||
        null == apPipeRunComponent.OPPIDComponent.GraphicsElement.Range)
        return returnList;


    // We are going to move identified cells to a new level. Create the level if it does not exist already
    string deletedComponentsLevelName = "ComponentsDeletedDuringAPPIDConversion";
    MsDgn.Level deletedComponentsLevel = MsDgnApplicationInstance.ActiveDesignFile.Levels.Find (deletedComponentsLevelName, null);
    if (null == deletedComponentsLevel)
        MsDgnApplicationInstance.ActiveDesignFile.AddNewLevel (deletedComponentsLevelName);

    deletedComponentsLevel = MsDgnApplicationInstance.ActiveDesignFile.Levels.Find (deletedComponentsLevelName, null);
    

    BmfGe.IRange scanRange = apPipeRunComponent.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;
    
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    scanCriteria.ExcludeAllTypes ();
    scanCriteria.IncludeType (MsDgn.MsdElementType.SharedCell);
    scanCriteria.IncludeOnlyCell (cellNameToSearch);
    if (null != levelToSearch)
        {
        scanCriteria.ExcludeAllLevels ();
        scanCriteria.IncludeLevel (levelToSearch);
        }

    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    while (elemEnum.MoveNext ())
        {
        MsDgn.SharedCellElement currentCell = elemEnum.Current.AsSharedCellElement ();
        // Ensure this cell has no xdata
        MsDgn.XDatum[] xdatas = currentCell.GetXData ("AT_");
        if (null != xdatas)
            continue;

        returnList.Add ((UInt32)elemEnum.Current.ID);
        currentCell.Level = deletedComponentsLevel;
        currentCell.Rewrite ();
        }

    
    deletedComponentsLevel.IsDisplayed = false;
    MsDgnApplicationInstance.ActiveDesignFile.Levels.Rewrite ();

    return returnList;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the leader line for tie in or specialty Item.
/// </summary>
/// <param name="elementID">The element ID.</param>
/// <param name="apComponent">The ap component.</param>
/// <author>Steve.Morrow</author>                                    <date>1/6/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public static void GetLeaderLineForBubbleIdentification(long elementID, APPIDComponent apComponent)
    {
    UInt32 id = (UInt32)elementID;
    MsDgn.Element elemEnum = ElementFromElementUnit(id);
    if (elemEnum == null)
        return;
        
    if (elemEnum.Type == MsDgn.MsdElementType.Line)
        {
        if (apComponent.InstrumentBubbleLeaderHandles == null)
            apComponent.InstrumentBubbleLeaderHandles = new ArrayList();

        apComponent.InstrumentBubbleLeaderHandles.Add(elementID);
        }
    
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete APPID bubble elements
/// This method determines which elements should be deleted. APPID allows for extra "instrument" lines to be drawn
//  This lines are not grouped. So they will not be deleted and will not be intelligent.
/// </summary>
/// <param name="bubble">bubble component</param> 
/// <param name="apComponent">appid component</param> 
/// <param name="elementIDs">element id</param> 
/// <author>Steve.Morrow</author>                            <date>09/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteAPPIDBubbleElements
(
SchInt.IBubble bubble,
APPIDComponent apComponent,
ArrayList elementIDs
)
    {
    //Determine which lines to remove. APPID allows for extra lines to be drawn. Also the Control Panel/Local panel is a line(s)
    //These extra lines do not need to be erased and cannot be made intelligent.
    MsDgn.Point3d p1 = new MsDgn.Point3d ();
    MsDgn.Point3d p2 = new MsDgn.Point3d ();
    p1.X = bubble.Origin.X - bubble.BubbleRadius / 4;
    p1.Y = bubble.Origin.Y - bubble.BubbleRadius / 4;
    p2.X = bubble.Origin.X + bubble.BubbleRadius / 4;
    p2.Y = bubble.Origin.Y + bubble.BubbleRadius / 4;

    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low = p1;
    range.High = p2;

    double subUnitsPerMasterUnit = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;

    range.Low.X /= subUnitsPerMasterUnit;
    range.Low.Y /= subUnitsPerMasterUnit;
    range.Low.Z /= subUnitsPerMasterUnit;

    range.High.X /= subUnitsPerMasterUnit;
    range.High.Y /= subUnitsPerMasterUnit;
    range.High.Z /= subUnitsPerMasterUnit;

    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    scanCriteria.ExcludeAllTypes ();
    scanCriteria.IncludeType (MsDgn.MsdElementType.Line);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    //remove the elements that make up the Location lines (panel,control...)
    ArrayList elementIDsToDelete = new ArrayList ();
    while (elemEnum.MoveNext ())
        {

        long el = elemEnum.Current.ID64;
        UInt32 id = (UInt32)el;
        if (isElementIDLeaderLine (apComponent, id))
            continue;

        //ensure scanned elements are lines
        if (elemEnum.Current.Type == MsDgn.MsdElementType.Line)
            {
            elementIDsToDelete.Add (id);
            elementIDs.Remove (id);
            }
        }
    DeleteElements (elementIDsToDelete);

    bool leaderDeleted = false;
    foreach (UInt32 id in elementIDs)
        {
        //deleted the leader
        if (isElementIDLeaderLine (apComponent, id))
            {
            leaderDeleted = true;
            DeleteElement (id);
            continue;
            }

        //don't delete the extra lines that could be drawn
        //these elements are lines
        //we want to delete LineString
        MsDgn.Element element = ElementFromElementUnit(id);
        if (null == element || element.Type == MsDgn.MsdElementType.Line ||
            element.Type == MsDgn.MsdElementType.CellHeader ||
            element.Type == MsDgn.MsdElementType.SharedCellDefinition ||
            element.Type == MsDgn.MsdElementType.SharedCell) 
            {
            //todo should there be a tolerance?
            //double length = element.AsLineElement ().Length;
            continue;
            }

        //delete everything else
        DeleteElement(id);
        }

    if (null != apComponent.InstrumentBubbleHandles)
        {
        //delete the bubble parts. These parts occur when the text breaks up the APPID bubble
        foreach (long elementID in apComponent.InstrumentBubbleHandles)
            {
            UInt32 id = (UInt32)elementID;

            //already been deleted
            if (elementIDs.Contains(id))
                continue;

            MsDgn.Element element = ElementFromElementUnit(id);
            if (null == element)
                continue;
            DeleteElement(id);
            }
        }
    //ensure leader line is deleted 
    if (leaderDeleted || 
        null == apComponent.InstrumentBubbleLeaderHandles || apComponent.InstrumentBubbleLeaderHandles.Count == 0)
        return;

    foreach (long elementLeader in apComponent.InstrumentBubbleLeaderHandles)
        {
        UInt32 ileader = (UInt32)elementLeader;
        DeleteElement (ileader);
        }

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Section Index for the Inrun connected to a PipeRun based on Origin as well as the.
/// Connect points and Also set the Insert by Mechanism accordingly.
/// </summary>
/// <param name="run">Pipe Run</param> 
/// <param name="apComponent">appid component</param> 
/// <author>Qarib.Kazmi</author>                            <date>08/2019</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static int               GetSectionIndexSetMechanism
(
SchInt.ISchematicsRun run,
APPIDComponent apComponent
)
{
    if ((null == run) || (null == apComponent))
        return -1;

    int index = -1;
    SchInt.ISchematicsInsertableComponent insertableComponent = apComponent.OPPIDComponent as SchInt.ISchematicsInsertableComponent;
    index = run.GetSectionIndex(apComponent.CellOrigin);
    if (index != -1)
        {
        insertableComponent.SetInsertByMechanism(SchInt.InsertByMechanism.Origin, null, SchInt.RangePointLocation.UpperLeft);
        return index;
        }
    //If no match for the Origin then Try out the Connect points for Connection
    Bom.IConnectPointCollection conPoints = apComponent.OPPIDComponent.ConnectPoints;
    if (null == conPoints)
        return index;

    foreach (Bom.IConnectPoint conPt in conPoints)
        {
        index = run.GetSectionIndex(conPt.Location);
        if (index != -1)
            {
            insertableComponent.SetInsertByMechanism(SchInt.InsertByMechanism.ConnectPoint, conPt.NameKey, SchInt.RangePointLocation.UpperLeft);
            return index;
            }
        }
    return index;
}
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get PipeRun For NonIntelligent InRun
/// </summary>
/// <param name="inRun">inRun component</param>
/// <param name="piperunComponents">List of piperuns</param> 
/// <returns> Pipe Run</returns>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static SchInt.ISchematicsRun GetPipeRunForNonIntelligentInRun
(
APPIDComponent inRun,
ArrayList piperunComponents
)
    {
    BmfGe.IRange scanRange = inRun.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;

    // Use scan criteria to get selection set and return
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    string guidValue = string.Empty;
    while (elemEnum.MoveNext ())
        {
        Bom.IComponent component = Bom.Workspace.ActiveModel.FindComponentById (elemEnum.Current.ID64);
        if (null != component)
            {
            if (!component.ClassDefinition.Is (s_pipeRunClass))
                continue;
 
            guidValue = GetComponentGuid (component);
            if (string.IsNullOrEmpty (guidValue))
                continue;
            if (string.IsNullOrEmpty(guidValue))
                return null;
            // Iterate all PipeRun components to find the matching guidValue
            foreach (APPIDComponent apPipeRun in piperunComponents)
                {
                if (null == apPipeRun.OPPIDComponent)
                    continue;

                string pipeGuid = GetComponentGuid(apPipeRun.OPPIDComponent);
                if (string.IsNullOrEmpty(pipeGuid))
                    continue;

                if (pipeGuid.Equals(guidValue, StringComparison.InvariantCultureIgnoreCase))
                    {
                    // If Matched then check if the inRun is located on any segment of the Current Piperun
                    SchInt.ISchematicsRun run = apPipeRun.OPPIDComponent as SchInt.ISchematicsRun;
                    //get index of run and Set Mechanism
                    int index = GetSectionIndexSetMechanism(run, inRun);
                    if (index == -1)
                        continue;
                    return run;
                    }
                }
            }
        }
    return null;
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Returns the elements within a given range
/// </summary>
/// <param name="angleComponentInPiperun">angle component</param> 
/// <param name="piperunElementIDs">piperun element id</param> 
/// <returns> The collection of intersecting elements or null if not intersecting </returns>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdatePiperunLineEndpointsBasedOnIntersectionRange
(
APPIDComponent angleComponentInPiperun,
ArrayList piperunElementIDs
)
    {

    BmfGe.IRange scanRange = angleComponentInPiperun.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;

    // Use scan criteria to get selection set and return
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    while (elemEnum.MoveNext ())
        {
        if (!elemEnum.Current.IsLineElement () || !piperunElementIDs.Contains (elemEnum.Current.ID64))
            continue;

        MsDgn.LineElement line = elemEnum.Current.AsLineElement ();
        BmfGe.IPoint aPt = angleComponentInPiperun.CellOrigin.Clone ();
        SchApp.SchematicsMicroStationSettings.ConvertSubUnitsToMasterUnits (aPt, out aPt);
        MsDgn.Point3d anglePt = new MsDgn.Point3d ();
        anglePt.X = aPt.X;
        anglePt.Y = aPt.Y;
        anglePt.Z = aPt.Z;

        int index = line.GetClosestVertex (ref anglePt);
        MsDgn.Point3d newPt = line.get_Vertex (index + 1);
        line.set_Vertex (index + 1, ref anglePt);
        line.Rewrite ();
        line.Redraw (MsDgn.MsdDrawingMode.Normal);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get PipeRuns from Scan Range
/// </summary>
/// <param name="range">element range in master units</param> 
/// <param name="piperunComponents">pipeRun components List</param>
/// <returns>array of Pipe Runs</returns>
/// <author>Qarib.Kazmi</author>                                <date>02/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static ArrayList         GetPipeRunsFromRange
(
MsDgn.Range3d range,
ArrayList pipeRunComponents,
int numberRunsRequired
)
    {
    // this scan will find every run whose total range, which includes all vertices, 
    // overlaps the given range
    SchInt.IRunCollection runs = new SchCat.RunCollection ();

    // scan range for runs
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);
    while (elemEnum.MoveNext ())
        {
        MsDgn.Element el = elemEnum.Current;
        Bom.IComponent component = Bom.Workspace.ActiveModel.FindComponentById (elemEnum.Current.ID64);
        if ((null != component) && (component is SchInt.ISchematicsRun))
            {
            SchInt.ISchematicsRun run = component as SchInt.ISchematicsRun;
            runs.Add (run);
            }
        }

    if (null == runs || runs.Count == 0)
        return null;

    // runs found within range (convert range to subunits)
    double subpermaster = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    BmfGe.IPoint low = new BmfGe.Point (range.Low.X * subpermaster, range.Low.Y * subpermaster);
    BmfGe.IPoint high = new BmfGe.Point (range.High.X * subpermaster, range.High.Y * subpermaster);
    BmfGe.IRange geRange = new BmfGe.Range (low, high);

    // filter out any runs that do not start or end within the range
    string guidValue = string.Empty;
    ArrayList guids = new ArrayList ();
    foreach (SchInt.ISchematicsRun currentRun in runs)
        {
        BmfGe.IPointCollection vertices = currentRun.Vertices;
        if (vertices.Count >= 2)
            {
            if (geRange.ContainsXY (vertices[0]) ||
                geRange.ContainsXY (vertices[vertices.Count - 1]))
                {
                guidValue = GetComponentGuid (currentRun);
                if (!string.IsNullOrEmpty (guidValue))
                    guids.Add (guidValue);
                }
            }
        }

    if (null == guids || guids.Count != numberRunsRequired)
        return null;

    ArrayList pipeRuns = new ArrayList ();
    foreach (APPIDComponent apPipeRun in pipeRunComponents)
        {
        if (null != apPipeRun.OPPIDComponent)
            {
            string pipeGuid = GetComponentGuid (apPipeRun.OPPIDComponent);
            if (!string.IsNullOrEmpty (pipeGuid))
                {
                if (guids.Contains (pipeGuid))
                    {
                    pipeRuns.Add (apPipeRun);
                    if (pipeRuns.Count == numberRunsRequired)
                        break;
                    }
                }
            }
        }

    return pipeRuns;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Scan Spec Break
/// </summary>
/// <param name="specBreak">specBreak element</param> 
/// <returns> Pipe Run</returns>
/// <param name="piperunComponents">pipeRun components List</param> 
/// <author>Qarib.Kazmi</author>                                <date>02/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static ArrayList         ScanSpecBreak
(
APPIDComponent specBreak,
ArrayList pipeRunComponents
)
    {
    //MsDgn.Range3d range = element.Range;
    BmfGe.IRange scanRange = specBreak.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;
    // Use scan criteria to get selection set and return
    ArrayList piperuns = GetPipeRunsFromRange (range, pipeRunComponents, 2);
    return piperuns;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Scan Run Terminator
/// </summary>
/// <param name="elementID">element id</param> 
/// <param name="piperunComponents">pipeRun components List</param>
/// <param name="numberRunsRequired">Number of runs required</param> 
/// <returns> Pipe Run</returns>
/// <modified>Qarib.Kazmi</author>                                <date>02/2015</date> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static ArrayList         ScanRunsForComponent
(
UInt32 elementID,
ArrayList pipeRunComponents,
int numberRunsRequired
)
    {
    MsDgn.Element element = ElementFromElementUnit (elementID);
    if (null == element)
        return null;

    MsDgn.Range3d range = element.Range;
    // Use scan criteria to get selection set and return
    ArrayList piperuns = GetPipeRunsFromRange (range, pipeRunComponents, numberRunsRequired);
    return piperuns;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Scan Run Terminator
/// </summary>
/// <param name="runTerm">runterm element</param> 
/// <returns> Pipe Run</returns>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static ArrayList         ScanReducer
(
APPIDComponent reducer
)
    {
    BmfGe.IRange scanRange = reducer.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;

    // Use scan criteria to get selection set and return
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    ArrayList runTerms = new ArrayList ();
    while (elemEnum.MoveNext ())
        {
        APPIDComponent apComponent = GetAPPIDComponentFromElementID (elemEnum.Current.ID64);

        if (null != apComponent && apComponent.IsRunTerm) 
            runTerms.Add (apComponent);

        if (runTerms.Count == 2)
            break;

        }

    return runTerms;
    }


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Connection Points
/// This is currently just used for reducers
/// </summary>
/// <param name="apComponent">angle component</param> 
/// <returns> The collection of intersecting elements or null if not intersecting </returns>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              GetAndSetConnectionPoints
(
APPIDComponent apComponent
)
    {

    if (!apComponent.IsReducer)
        return;

    BmfGe.IRange scanRange = apComponent.OPPIDComponent.GraphicsElement.Range;
    MsDgn.Range3d range = new MsDgn.Range3d ();
    range.Low.X = scanRange.Low.X;
    range.Low.Y = scanRange.Low.Y;
    range.Low.Z = scanRange.Low.Z;
    range.High.X = scanRange.High.X;
    range.High.Y = scanRange.High.Y;
    range.High.Z = scanRange.High.Z;

    // Converting the range from sub units to Master Units    
    double ratio = MsDgnApplicationInstance.ActiveModelReference.SubUnitsPerMasterUnit;
    range.Low.X /= ratio;
    range.Low.Y /= ratio;
    range.Low.Z /= ratio;
    range.High.X /= ratio;
    range.High.Y /= ratio;
    range.High.Z /= ratio;

    // Use scan criteria to get selection set and return
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.IncludeOnlyWithinRange (ref range);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);

    //BmfGe.IPoint cp1 = new BmfGe.Point ();
    //BmfGe.IPoint cp2 = new BmfGe.Point (mscp2.X, mscp2.Y);
    BmfGe.IPoint startPt = null;
    BmfGe.IPoint endPt = null;

    while (elemEnum.MoveNext ())
        {
        if (!elemEnum.Current.IsLineElement ())
            continue;

        MsDgn.LineElement line = elemEnum.Current.AsLineElement ();
        BmfGe.IPoint aPt = apComponent.CellOrigin.Clone ();
        SchApp.SchematicsMicroStationSettings.ConvertSubUnitsToMasterUnits (aPt, out aPt);
        MsDgn.Point3d dataChangePt = new MsDgn.Point3d ();
        dataChangePt.X = aPt.X;
        dataChangePt.Y = aPt.Y;
        dataChangePt.Z = aPt.Z;

        int index = line.GetClosestVertex (ref dataChangePt);
        MsDgn.Point3d pt1 = line.get_Vertex (1);
        MsDgn.Point3d pt2 = line.get_Vertex (2);
        if (0 == index)
            {
            startPt = new BmfGe.Point (pt1.X, pt1.Y);
            endPt = new BmfGe.Point (pt2.X, pt2.Y);
            }
        else
            {
            startPt = new BmfGe.Point (pt2.X, pt2.Y);
            endPt = new BmfGe.Point (pt1.X, pt1.Y);
            }

        BmfGe.IVector dir = new BmfGe.Vector (startPt, endPt);
        dir.Normalize ();
        SchApp.SchematicsMicroStationSettings.ConvertMasterUnitsToSubUnits (startPt, out startPt);

        if (null == apComponent.ConnectionPtCP1)
            {
            apComponent.ConnectionPtCP1 = startPt.Clone ();
            apComponent.DirectionCP1 = dir.Clone ();
            }
        else if (null == apComponent.ConnectionPtCP2)
            {
            apComponent.ConnectionPtCP2 = startPt.Clone ();
            apComponent.DirectionCP2 = dir.Clone ();
            }
        else if (null != apComponent.ConnectionPtCP1 && null != apComponent.ConnectionPtCP2)
            break;

        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete APPID Settings Block
/// </summary>
/// <author>Steve.Morrow</author>                                <date>00/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DeleteAPPIDSettingsBlock
(
)
    {
    deleteAPPIDBlock("_at_asi_ldwg");
    deleteAPPIDBlock("AT_pid_SETT");
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete APPID Settings Block
/// </summary>
/// <author>Steve.Morrow</author>                                <date>00/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             deleteAPPIDBlock
(
string cellName
)
    {
    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.ExcludeAllTypes ();
    scanCriteria.IncludeOnlyCell (cellName);
    scanCriteria.IncludeType (MsDgn.MsdElementType.SharedCell);
    MsDgn.ElementEnumerator elemEnum = MsDgnApplicationInstance.ActiveModelReference.Scan (scanCriteria);
    ArrayList elementIDsToDelete = new ArrayList ();
    while (elemEnum.MoveNext ())
        {
        if (!elemEnum.Current.IsSharedCellElement ())
            continue;
        long el = elemEnum.Current.ID64;
        UInt32 id = (UInt32)el;
        if (!elementIDsToDelete.Contains (id))
            elementIDsToDelete.Add (id);
        }

    DeleteElements (elementIDsToDelete);
    }
#endregion

#region SQL methods

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Process Keytag From PipeRun Keytag
/// </summary>
/// <param name="inRunKeytag">PipeRun Keytag</param>
/// <returns>Process keytag</returns> 
/// <author>Steve.Morrow</author>                            <date>12/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetProcessKeytagFromPipeRunKeytag
(
string pipeRunKeytag
)
    {
    string sqlStatement = string.Format ("select LINE_ID from PIPE_RUN where KEYTAG = '{0}'", pipeRunKeytag);
    //return getSqlStringValue (sqlStatement, "LINE_ID");
    return "test";
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get Piperun Keytag From InRun Keytag
/// </summary>
/// <param name="inRunKeytag">inRun Keytag</param>
/// <returns>pipe run keytag</returns> 
/// <author>Steve.Morrow</author>                            <date>08/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetPiperunKeytagFromInRunKeytag
(
string inRunKeytag
)
    {
    string sqlStatement = string.Format ("select RUN_ID from RUN_CONN where KEYTAG = '{0}'", inRunKeytag);
    //return getSqlStringValue (sqlStatement, "RUN_ID");
    return "test2";
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get RUN_IDs from the passed in keytag
/// * Keytag can be a reducer and the returns would be Run Terminator ids
///      Should be 2 run term components
/// * Keytag can be a 3 way or four way valve and the return would be multiple run ids
/// </summary>
/// <param name="inRunKeytag">in run Keytag</param>
/// <param name="runKeyTags">run ids found from DB(out parameter)</param>
/// <returns>apcomponent list</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
//public static ArrayList         GetRUN_IDComponentsFromInRunKeytag
//(
//string inRunKeytag,
//out ArrayList runKeyTags
//)
//    {
//    //string sqlStatement = string.Format ("select RUN_ID from RUN_CONN where KEYTAG = '{0}'", inRunKeytag);
//    //ADODB.Recordset recordSet = GetRecordset (sqlStatement);
//    ArrayList runidAPComponentList = new ArrayList ();
//    //runKeyTags = new ArrayList ();

//    //try
//    //    {

//    //    //Do query to get value;
//    //    if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 0)
//    //        return null;

//    //    recordSet.MoveFirst ();
//    //    while (!recordSet.EOF)
//    //        {
//    //        string runidKeytag = string.Empty;
//    //        foreach (ADODB.Field selectedField in recordSet.Fields)
//    //            {
//    //            if (selectedField.Name.Equals ("RUN_ID") && !DBNull.Value.Equals (selectedField.Value))
//    //                runidKeytag = selectedField.Value.ToString ().Trim ();
//    //            }
//    //        if (string.IsNullOrEmpty (runidKeytag))
//    //            return null;

//    //        runKeyTags.Add (runidKeytag);
//    //        APPIDComponent comp = GetAPPIDComponentFromKeytag (runidKeytag);
//    //        if (null == comp)
//    //            return null;

//    //        runidAPComponentList.Add (comp);
//    //        recordSet.MoveNext ();
//    //        }
//    //    }
//    //catch (Exception ex)
//    //    {
//    //    CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
//    //    CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
//    //    return null;
//    //    }
//    //finally
//    //    {
//    //    if (null != recordSet && recordSet.State == 1)
//    //        recordSet.Close ();
//    //    }

//    return runidAPComponentList;
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get RUN_IDs from the passed in keytag
/// * Keytag can be a reducer and the returns would be Run Terminator ids
///      Should be 2 run term components
/// * Keytag can be a 3 way or four way valve and the return would be multiple run ids
/// </summary>
/// <param name="inRunKeytag">in run Keytag</param>
/// <returns>apcomponent list</returns> 
/// <author>Naveed.Khalid</author>                            <date>09/2014</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
//public static ArrayList         GetRUN_IDComponentsFromInRunKeytag
//(
//string inRunKeytag
//)
//    {
//    ArrayList runKeyTags = null;
//    return GetRUN_IDComponentsFromInRunKeytag (inRunKeytag, out runKeyTags);
//    }
/*------------------------------------------------------------------------------------**/
/// <summary>
///The document inwhich the link to the component is on has been converted to an OPPID document
/// </summary>
/// <param name="apComponent">apComponent</param>
/// <returns>true if the document has converted</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool            ComponentsParentDocumentBeenConverted
(
APPIDComponent apComponent
)
    {

    ////query keylink to get all link ids
    //string sqlStatement = string.Format ("SELECT LINK_ID from KEY_LINK where KEYTAG = '{0}'", apComponent.Keytag);
    //ADODB.Recordset recordSet = null;

    //try
    //    {
    //    recordSet = GetRecordset (sqlStatement);
    //    if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 1)
    //        return false;

    //    recordSet.MoveFirst ();
    //    while (!recordSet.EOF)
    //        {
    //        string componentLinkID = string.Empty;
    //        foreach (ADODB.Field selectedField in recordSet.Fields)
    //            {
    //            if (selectedField.Name.Equals ("LINK_ID") && !DBNull.Value.Equals (selectedField.Value))
    //                componentLinkID = selectedField.Value.ToString ().Trim ();
    //            }

    //        if (string.IsNullOrEmpty (componentLinkID) || apComponent.KeyLink.Equals (componentLinkID))
    //            {
    //            recordSet.MoveNext ();
    //            continue;
    //            }

    //        //Check document
    //        sqlStatement = string.Format ("SELECT doc_reg.REGAPP FROM doc_reg INNER JOIN {0} ON doc_reg.DOC_ID = {0}.DWG_NAME WHERE {0}.LINK_ID='{1}'",
    //                                             apComponent.LinkTab, componentLinkID);
    //        //string regapp = getSqlStringValue (sqlStatement, "REGAPP");
    //        string regapp = "test3";
    //        //if the regapp is null or equal to AT_PID, drawing has not been converted
    //        if (string.IsNullOrEmpty (regapp) || regapp.Equals ("AT_PID"))
    //            {
    //            recordSet.MoveNext ();
    //            continue;
    //            }

    //        //if the regapp is equal to OPPID, drawing has been converted
    //        if (regapp.Equals ("OPPID"))
    //            return true;

    //        recordSet.MoveNext ();
    //        }
    //    }
    //catch (Exception ex)
    //    {
    //    CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
    //    CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
    //    return false;
    //    }
    //finally
    //    {
    //    if (null != recordSet && recordSet.State == 1)
    //        recordSet.Close ();
    //    }

    return false;

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get Value From SqlString
/// </summary>
/// <param name="sqlStatement">sqlStatement</param> 
/// <param name="keytag">equipment or piperun/pipeline keytag</param> 
/// <returns>string value</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetDBFieldValueFromForeignKey
(
string sqlStatement,
string fieldName
)
    {
    //return getSqlStringValue (sqlStatement, fieldName);
    return "test3";
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Keytag From Link Values
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <returns>keytag</returns> 
/// <author>Steve.Morrow</author>                            <date>09/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static string           getKeytagFromLinkValues
(
APPIDComponent apComponent
)
    {
    string sqlStatement = string.Format("select KEYTAG from KEY_LINK where LINK_ID = '{0}' and LINK_TAB = '{1}'",apComponent.KeyLink, apComponent.LinkTab );
    //return getSqlStringValue (sqlStatement, "KEYTAG");
    return "Test4";
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the document values for page connectors.
/// </summary>
/// <param name="keytag">The keytag.</param>
/// <param name="linkedName">Name of the linked.</param>
/// <param name="linkedFileName">Name of the linked file.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                    <date>1/11/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
//public static bool GetDocumentValuesForPageConnectors
//(
//string keytag,
//out string linkedName,
//out string linkedFileName
//)
//    {
//    //linkedName = string.Empty;
//    //linkedFileName = string.Empty;

//    //try
//    //    {
//    //    string sqlStatement = string.Format("SELECT PID_TOFR.KEYTAG, doc_reg.DOC_NAME, doc_reg.DOC_FNAME FROM PID_TOFR INNER JOIN ((KEY_LINK INNER JOIN PID_TLNK ON KEY_LINK.LINK_ID = PID_TLNK.LINK_ID) INNER JOIN doc_reg ON PID_TLNK.DWG_NAME = doc_reg.DOC_ID) ON PID_TOFR.MATCH_KEY = KEY_LINK.KEYTAG WHERE (((PID_TOFR.KEYTAG)='{0}'));", keytag);
//    //    ADODB.Recordset recordSet = GetRecordset(sqlStatement);
//    //    //Do query to get value;
//    //    if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 0)
//    //        return false;

//    //    recordSet.MoveFirst();
//    //    ADODB.Field fld = recordSet.Fields["DOC_NAME"];
//    //    if (!DBNull.Value.Equals(fld.Value))
//    //        linkedName = fld.Value.ToString();

//    //    fld = recordSet.Fields["DOC_FNAME"];
//    //    if (!DBNull.Value.Equals(fld.Value))
//    //        linkedFileName = fld.Value.ToString();

//    //    recordSet.Close();
//    //    }
//    //catch
//    //    {
//    //    return false;
//    //    }

//    return true;
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update Run Term NEQUIP and NLINE values
/// </summary>
/// <param name="apComponent">appid component runterm</param> 
/// <param name="piperuns">Piperuns to get keytags from</param> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateRunTermDatabaseValues
(
APPIDComponent runTerm,
ArrayList piperuns
)
    {
    if (piperuns == null || piperuns.Count != 2)
        return;

    APPIDComponent runAPPIDcomponent1 = piperuns[0] as APPIDComponent;
    APPIDComponent runAPPIDcomponent2 = piperuns[1] as APPIDComponent;
    string NLINE = runAPPIDcomponent1.Keytag;
    string NEQUIP = runAPPIDcomponent2.Keytag;
    string NINT1 = "RUN_TERM";
    string NDESC = "Run Term";

    //check to see if the runterm records has been applied to the nozzle table. Do not want to commit twice.
    //this could happen twice on a spec break.
    string sqlStatement = string.Format ("select keytag from NOZZLE where NEQUIP = '{0}' and NLINE = '{1}' or NLINE = '{2}' and NEQUIP = '{3}'", 
                               NEQUIP, NLINE, NLINE, NEQUIP);
    bool updateRunterm = isRuntermLinkAlreadyCommitted (sqlStatement);

    if (!updateRunterm)
        {
        sqlStatement = string.Format ("Update NOZZLE set NLINE = '{0}' , NEQUIP = '{1}', NINT1 = '{2}', NDESC = '{3}' where KEYTAG = '{4}'",
                          NLINE, NEQUIP, NINT1, NDESC, runTerm.Keytag);
        ExecuteSqlStatement (sqlStatement);
        }
    else
        {
        sqlStatement = string.Format ("Delete from NOZZLE where KEYTAG = '{0}' ", runTerm.Keytag);
        ExecuteSqlStatement (sqlStatement);
        }

    DeleteRunTermKeylinkDatabaseValues (runTerm);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Check to see if the run term has been committed to the nozzle table
/// </summary>
/// <param name="sqlStatement">sql string</param>
/// <returns>true if has, false it it has not</returns> 
/// <author>Steve.Morrow</author>                            <date>02/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             isRuntermLinkAlreadyCommitted
(
string sqlStatement 
)
    {
    //ADODB.Recordset recordSet = GetRecordset (sqlStatement);
    bool status = false;

    //try
    //    {

    //    //Do query to get value;
    //    if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 0)
    //        status = false;
    //    else if (recordSet.RecordCount >= 0)
    //        status = true;

    //    }
    //catch (Exception ex)
    //    {
    //    CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
    //    CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
    //    return status;
    //    }
    //finally
    //    {
    //    if (null != recordSet && recordSet.State == 1)
    //        recordSet.Close ();
    //    }

    return status;
    }


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update Reducer Run Term values
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <param name="piperuns">Piperuns to get keytags from</param> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdateReducerRunTermDatabaseValues
(
APPIDComponent reducer,
APPIDComponent runTerm
)
    {
    //This is the ReducerKeytag
    string NEQUIP = reducer.Keytag;

    //Currently these values are required as they are defined in "REDUCER_CONNECTS_TO_SEGMENT" in file PlantProjectSchema.01.03_Autoplant_PIW.01.03.mapping.xml
    //TODO Should these be made into a cfg variable value???
    //Because these constants could be changed in "REDUCER_CONNECTS_TO_SEGMENT" 
    string NINT1 = "RED_TERM";
    string NDESC = "RED Term";

    string sqlStatement = string.Format ("Update NOZZLE set NEQUIP = '{0}', NINT1 = '{1}', NDESC = '{2}' where KEYTAG = '{3}'", NEQUIP, NINT1, NDESC, runTerm.Keytag);
    ExecuteSqlStatement (sqlStatement);

    DeleteRunTermKeylinkDatabaseValues (runTerm);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Create Reducer Data Change Record
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              CreateReducerDataChangeRecord
(
APPIDComponent reducer
)
    {
    string guid = GetComponentGuid (reducer.OPPIDComponent);
    if (string.IsNullOrEmpty (guid))
        return;

    string sqlStatement = string.Format("insert into DATACHANGE (KEYTAG, TAG_TYPE, COMP_ID) VALUES ('{0}', '{1}', '{2}')",reducer.Keytag, "AT_PID_REDUCER", "{" + guid + "}");
    ExecuteSqlStatement (sqlStatement);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Delete RunTerm Database link Values, these are only needed for APPID
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void             DeleteRunTermKeylinkDatabaseValues
(
APPIDComponent apComponent
)
    {
    string sqlStatement = string.Format( "Delete from KEY_LINK where LINK_ID = '{0}' and LINK_TAB = '{1}'",apComponent.KeyLink, apComponent.LinkTab);
    ExecuteSqlStatement (sqlStatement);

    sqlStatement = string.Format( "Delete from " + apComponent.LinkTab + " where LINK_ID = '{0}' and TAG_TYPE = '{1}'",apComponent.KeyLink, apComponent.TagType );
    ExecuteSqlStatement (sqlStatement);
    }


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Update Piperun Nominal Values
/// update the NOMINAL_DIAMETER, this is required
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <param name="run">run to update</param> 
/// <author>Steve.Morrow</author>                            <date>02/2012</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              UpdatePiperunNominalValues
(
APPIDComponent apComponent,
PidInt.IPipeRun run
)
    {

    IECPropertyValue sizeProperty = BCT.ECPropertyValueUtility.Get(run, 
                                  "SIZE", "", false);

    if (null == sizeProperty || sizeProperty.IsNull)
        return;
    string size = sizeProperty.StringValue;  
    if (string.IsNullOrEmpty (size))
        return;

    string nominalSize = "";
    foreach (KeyValuePair<string, string> data in pipeSizes)
        if (data.Key.Equals (size, StringComparison.OrdinalIgnoreCase))
            {
            nominalSize = data.Value.ToString ();
            break;
            }

    //set a default
    if (string.IsNullOrEmpty (nominalSize))
        nominalSize = "1";

    PIDUtilities.UpdatePipeRunNominalPropertyFromSize (sizeProperty, nominalSize);
    string sqlStatement = string.Format ("Update PIPE_RUN set NOM_DIAMETER = '{0}' where KEYTAG = '{1}'", 
                                         nominalSize, apComponent.Keytag);
    ExecuteSqlStatement (sqlStatement);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get Sql Values
/// </summary>
/// <param name="sqlStatement">Sql statement</param> 
/// <returns>object of values from query</returns> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
//private static string           getSqlStringValue
//(
//string sqlStatement,
//string fieldName
//)
//    {
//    ADODB.Recordset recordSet = null;
//    string sqlValue = string.Empty;
//    try
//        {
//        recordSet = BmfPer.PersistenceManager.GetDataFromSql (sqlStatement);
//        if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 0)
//            return sqlValue;

//        ADODB.Field fld = recordSet.Fields[fieldName];
//        if (!DBNull.Value.Equals (fld.Value))
//            sqlValue = fld.Value.ToString ();

//        }
//    catch (Exception ex)
//        {
//        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
//        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
//        return sqlValue;
//        }
//    finally
//        {
//        if (null != recordSet && recordSet.State == 1)
//            recordSet.Close ();
//        }
//    return sqlValue; 
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get Sql Recordset
/// </summary>
/// <param name="sqlStatement">Sql statement</param> 
/// <returns>Recordset</returns> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
//public static ADODB.Recordset  GetRecordset
//(
//string sqlStatement
//)
//    {
//    ADODB.Recordset recordSet = null;
//    try
//        {
//        recordSet = BmfPer.PersistenceManager.GetDataFromSql (sqlStatement);
//        }
//    catch (Exception ex)
//        {
//        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
//        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
//        }
//    return recordSet;
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Execute Sql Values
/// </summary>
/// <param name="sqlStatement">Sql statement</param> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void             ExecuteSqlStatement
(
string sqlStatement
)
    {
    try
        {
        //BmfPer.PersistenceManager.GetDataFromSql (sqlStatement);
        }
    catch (Exception ex)
        {
        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
        CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
        }
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates the GUID in Doc_Reg Table of ProjectDB .
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
//public static void              UpdateDocumentIDGuid
//(
//)
//    {
//    try
//        {
//        string globalIDPropertyName = BmfUt.GlobalIDUtilities.GetGlobalIDPropertyName (Bom.Workspace.ActiveModel);
//        if (string.IsNullOrEmpty (globalIDPropertyName))
//            return;

//        IECPropertyValue globalIDPropertyValue = Bom.Workspace.ActiveModel.FindPropertyValue (globalIDPropertyName, true, false, false);
//        if (null == globalIDPropertyValue || globalIDPropertyValue.IsNull || string.IsNullOrEmpty (globalIDPropertyValue.StringValue))
//            return;

//        BmfPer.PersistenceManager.UpdateDOC_REGGuidValues ("DOC_ID", APPIDDocumentID, "DOC_ID_GUID_PK", globalIDPropertyValue.StringValue, "OPPID");
//        string sqlStatement = string.Format ("Update DOC_REG set TAG_TYPE = 'AT_DWG_NAME' where DOC_ID = '{0}'", APPIDDocumentID);
//        ExecuteSqlStatement (sqlStatement);

//        string filePath = System.IO.Path.GetDirectoryName (MsDgnApplicationInstance.ActiveDesignFile.FullName);
//        BmfPer.PersistenceManager.EnsureDocumentPathAndUpdateIfIncorrect (APPIDDocumentID, filePath);
//        }
//    catch (Exception e)
//        {
//        CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.WARNING, e);
//        }
//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// assign Pipeline Class From Piperun Pint1
/// </summary>
/// <remarks>
/// Must trim the string as it has leading spaces "   MAJOR"
/// </remarks> 
/// <author>Steve.Morrow</author>                            <date>08/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static string           assignPipelineClassFromPiperun
(
string pipelineKeytag
)
    {
    //string sqlStatement = string.Format( "select PINT1 from PIPE_RUN where LINE_ID = '{0}'",pipelineKeytag);
    //ADODB.Recordset recordSet = GetRecordset (sqlStatement);

    //try
    //    {
    //    if (null == recordSet || recordSet.State == 0 || recordSet.RecordCount <= 0)
    //        return APPIDProviderUtilities.DefaultAPPIDProcessLineClass;

    //    recordSet.MoveFirst ();
    //    if (recordSet.RecordCount == 1)
    //        {
    //        ADODB.Field fld = recordSet.Fields["PINT1"];
    //        if (!DBNull.Value.Equals (fld.Value))
    //            {
    //            string pint1 = fld.Value.ToString ().Trim ();
    //            if (string.IsNullOrEmpty (pint1))
    //                return APPIDProviderUtilities.DefaultAPPIDProcessLineClass;

    //            return pint1;
    //            }
    //        }

    //    //return list can contain multiple appid class names "MAJOR,MINOR,..."
    //    //Need to have a schema list to determine which appid class takes Priority
    //    string highestPriorityClass = string.Empty;
    //    int currentHighestPriority = -1;
    //    while (!recordSet.EOF)
    //        {
    //        string pint1 = string.Empty;
    //        foreach (ADODB.Field selectedField in recordSet.Fields)
    //            {
    //            if (selectedField.Name.Equals ("PINT1") && !DBNull.Value.Equals (selectedField.Value))
    //                pint1 = selectedField.Value.ToString ().Trim ();
    //            }

    //        //move to the next record of recordset
    //        recordSet.MoveNext ();
    //        if (!string.IsNullOrEmpty (pint1))
    //            {
    //            int priority = GetProcessLinePriority (pint1);
    //            if (priority != -1)
    //                {
    //                if (currentHighestPriority == -1)
    //                    {
    //                    currentHighestPriority = priority;
    //                    highestPriorityClass = pint1;
    //                    }
    //                else if (priority < currentHighestPriority)
    //                    {
    //                    currentHighestPriority = priority;
    //                    highestPriorityClass = pint1;
    //                    }
    //                }
    //            }
    //        }
    //    if (!string.IsNullOrEmpty (highestPriorityClass))
    //        return highestPriorityClass;
    //    }
    //catch (Exception ex)
    //    {
    //    CTLog.FrameworkLogger.Instance.SubmitException (BsiLog.SEVERITY.ERROR, ex);
    //    CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.ERROR, sqlStatement);
    //    return null;
    //    }
    //finally
    //    {
    //    if (null != recordSet && recordSet.State == 1)
    //        recordSet.Close ();
    //    }

    //return APPIDProviderUtilities.DefaultAPPIDProcessLineClass;
    return "Pipeline";
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get OPPID Instrument Class From APPID Instrument IINT1 Value
/// </summary>
/// <param name="instrumentKeytag">instrument keytag</param> 
/// <param name="defaultOPPIDClassName">default OPPID class name</param> 
/// <returns> OPPID class name</returns> 
/// <remarks>
/// Must trim the string as it has leading spaces "   ISADCD"
/// </remarks> 
/// <author>Steve.Morrow</author>                            <date>09/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetOPPIDInstrumentClassFromAPPIDInstrumentIINT1Value
(
string instrumentKeytag,
string defaultOPPIDClassName
)
    {
    return getInt1ValueFromDatabase(instrumentKeytag, defaultOPPIDClassName, "INSTR", "IINT1", false);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the OPPID valve class from VINT1 value.
/// </summary>
/// <param name="valveKeyTag">The valve key tag.</param>
/// <param name="defaultOPPIDClassName">Default name of the OPPID class.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                   <date>12/8/2015</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public static string GetOPPIDValveClassFromAPPIDVINT1Value
(
string valveKeytag,
string defaultOPPIDClassName
)
    {
    return getInt1ValueFromDatabase(valveKeytag, defaultOPPIDClassName, "VALVE", "VINT1",true);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Determines whether [is equipment in parametric list] [the specified keytag].
/// </summary>
/// <param name="keytag">The keytag.</param>
/// <returns>
///   <c>true</c> if [is equipment in parametric list] [the specified keytag]; otherwise, <c>false</c>.
/// </returns>
/// <author>Steve.Morrow</author>                                    <date>1/4/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public static bool IsEquipmentInParametricList
(
APPIDComponent apComponent
)
    {

    string keytag = apComponent.Keytag;
    string sourceTable = getSourceTableName (apComponent);
    if (string.IsNullOrEmpty (sourceTable))
        sourceTable = "EQUIP";

    string sqlStatement = string.Format ("select {0} from {1} where KEYTAG = '{2}'", "EINT1", sourceTable, keytag);
    string keyName = "Test5:"; //getSqlStringValue(sqlStatement, "EINT1");
    if (keyName == null)
        keyName = string.Empty;

    keyName = keyName.ToUpper ().Trim ();
    if (string.IsNullOrEmpty (keyName))
        {
        //have found that some parametric components (AT_EQEXCH) do not have a function name.
        //Possibly require user to fill out this field??
        //TODO log this as a warning
        return true;
        }

    if (null == m_equipmentParametricList)
        m_equipmentParametricList = ConversionMappingUtilities.GetStringCollection("EQUIPMENT_PARAMETRICS_LIST", ConversionCatalog);

    if (null != m_equipmentParametricList)
        return m_equipmentParametricList.Contains(keyName);

    return false;
    }
private static ArrayList m_equipmentParametricList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the name of the source table.
/// </summary>
/// <param name="apComponent">The ap component.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                    <date>3/3/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
private static string getSourceTableName
(
APPIDComponent apComponent
)
    {
    //Dictionary<string, string> tagTypeTables = BmfPer.PersistenceManager.GetTagTypeFromSchemaCollection (apComponent.TagType);

    //if (null == tagTypeTables || tagTypeTables.Count == 0)
    //    return null;

    //if (tagTypeTables.ContainsKey ("SOURCE_TAB"))
    //    return tagTypeTables["SOURCE_TAB"].ToString ();

    return null;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the int1 value from database.
/// </summary>
/// <param name="keytag">The keytag.</param>
/// <param name="defaultOPPIDClassName">Default name of the OPPID class.</param>
/// <param name="tableName">Name of the table.</param>
/// <param name="columnName">Name of the column.</param>
/// <returns></returns>
/// <author>Steve.Morrow</author>                                    <date>12/8/2015</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
private static string getInt1ValueFromDatabase
(
string keytag,
string defaultOPPIDClassName,
string tableName,
string columnName,
bool isIsolationKey

)
    {
    string sqlStatement = string.Format("select {0} from {1} where KEYTAG = '{2}'", columnName,tableName, keytag);
    string name = "test6"; // getSqlStringValue(sqlStatement, columnName);

    //Do query to get value;
    if (string.IsNullOrEmpty(name))
        return defaultOPPIDClassName;

    name = name.Trim();
    if (string.IsNullOrEmpty(name))
        return defaultOPPIDClassName;

    //Test name to determine if the value is in the ISOLATION_VALVE_LIST list
    //this is needed because some isolation values have a block name and key name that are the same
    //"FLOAT" is a block name for an isolation valve (but is not intelligent). The AP_34100 block is a GATE_VALVE_FLOAT which has a VINT1 value of "FLOAT"
    //this test will hel solve this issue of a block name and a key name being the same
    //FLOAT should not be in the ISOLATION_VALVE_LIST list. The original APPID key value for block name of FLOAT is IFLOAT 
    if (isIsolationKey)
        {
        if (!IsIntelligentIsolationValve(name))
            return defaultOPPIDClassName;
        }

    IConversionMappingDefinition conversionMappingDefintion = ConversionMappingUtilities.GetMappingDefintionFromSourceClass
            (ConversionCatalog, name, false);

    if (null != conversionMappingDefintion)
        return conversionMappingDefintion.TargetClassName;

    return defaultOPPIDClassName;
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get APPID Instrument ILOC_PID Value
/// </summary>
/// <param name="instrumentKeytag">instrument keytag</param> 
/// <returns> Location Value</returns> 
/// <remarks>
/// The value in the database is L,BUB. L is the code that is needed
/// </remarks> 
/// <author>Steve.Morrow</author>                            <date>09/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
//public static string            GetAPPIDLocationValue
//(
//string instrumentKeytag
//)
//    {
//    //L = Field Panel
//    string defaultLocation = "L";
//    string sqlStatement = string.Format ("select ILOC_PID from INSTR where KEYTAG = '{0}'", instrumentKeytag);
//    string locationValue = getSqlStringValue (sqlStatement, "ILOC_PID");

//    //Do query to get value;
//    if (string.IsNullOrEmpty(locationValue))
//        return defaultLocation;

//    return locationValue.Substring (0, 1);

//    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get OPPID COmponent From Keytag
/// Scan conversion list and get APPIDComponent object 
/// </summary>
/// <param name="keytag">keytag</param> 
/// <returns>component</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static Bom.IComponent    GetOPPIDComponentFromKeytag
(
string keytag
)
    {
    APPIDComponent apComponent = GetAPPIDComponentFromKeytag (keytag);
    if (null == apComponent)
        return null;

    return apComponent.OPPIDComponent;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get APPIDComponent From Keytag
/// Scan conversion list and get APPIDComponent object 
/// </summary>
/// <param name="keytag">keytag</param> 
/// <returns>AP component</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static APPIDComponent    GetAPPIDComponentFromKeytag
(
string keytag
)
    {
    foreach (APPIDComponent apComponent in APPIDComponents.Values)
        {
        if (apComponent.Keytag.Equals (keytag))
            return apComponent;
        }
    return null;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get the current rotation existent on the element 
/// </summary>
/// <param name="elementID"> The element ID of the source element </param>
///<author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static BmfGe.IRotMatrix  GetRotationfromSharedCell
(
long elementID
)
    {
    MsDgn.Element element = BmfUt.MiscellaneousUtilities.GetElementFromElementID(elementID);
    if (null == element)
        return null;

    MsDgn.SharedCellElement cellElement = element as MsDgn.SharedCellElement;
    if (null == cellElement)
        return null;

    BmfGe.IRotMatrix rotMat = new BmfGe.RotMatrix();

    rotMat = new BmfGe.RotMatrix(cellElement.Rotation.RowX.X, cellElement.Rotation.RowX.Y, cellElement.Rotation.RowX.Z,
                                   cellElement.Rotation.RowY.X, cellElement.Rotation.RowY.Y, cellElement.Rotation.RowY.Z,
                                   cellElement.Rotation.RowZ.X, cellElement.Rotation.RowZ.Y, cellElement.Rotation.RowZ.Z);

    return rotMat;
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get APPIDComponent From ID
/// Scan conversion list and get APPIDComponent object 
/// </summary>
/// <param name="long">id</param> 
/// <returns>AP component</returns> 
/// <author>Steve.Morrow</author>                            <date>10/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static APPIDComponent    GetAPPIDComponentFromElementID
(
long id
)
    {
    foreach (APPIDComponent apComponent in APPIDComponents.Values)
        {
        if (apComponent.ID == id)
            return apComponent;
        }
    return null;
    }
#endregion

#region Methods
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Dispose All Objects
/// </summary>
///<author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              DisposeObjects
(
)
    {
    if (null != m_APPIDComponents)
        m_APPIDComponents.Clear ();
    m_APPIDComponents = null;

    m_msDgnApplicationInstance = null;
    m_conversionCatalog = null;

    if (null != m_attributeBlockNameTagList)
        m_attributeBlockNameTagList.Clear ();
    m_attributeBlockNameTagList = null;

    if (null != m_tagTypesToUseAlternateClass)
        m_tagTypesToUseAlternateClass.Clear ();
    m_tagTypesToUseAlternateClass = null;

    if (null != m_attributeBlockNameTagList)
        m_attributeBlockNameTagList.Clear ();
    m_attributeBlockNameTagList = null;

    if (null != m_blockNames)
        m_blockNames.Clear ();
    m_blockNames = null;

    if (null != m_nonIntelligentInRunBlockNames)
        m_nonIntelligentInRunBlockNames.Clear ();
    m_nonIntelligentInRunBlockNames = null;

    if (null != m_doNotConvertblockNames)
        m_doNotConvertblockNames.Clear ();
    m_doNotConvertblockNames = null;

    if (null != APPIDBubbleTagTypeList)
        APPIDBubbleTagTypeList.Clear ();
    APPIDBubbleTagTypeList = null;

    if (null != AdditionalXdataItems)
        AdditionalXdataItems.Clear();
    AdditionalXdataItems = null;

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get OPPID Guid Property String Value
/// <param name="oppidComponent">OPPID component</param> 
/// </summary>
/// <returns>GUID Value</returns> 
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetComponentGuid
(
Bom.IComponent oppidComponent
)
    {
    string globalIDPropertyName = BCT.GlobalIdUtility.GetValue (oppidComponent);
    if (string.IsNullOrEmpty (globalIDPropertyName))
        return null;

    IECPropertyValue globalIDPropertyValue = oppidComponent.FindPropertyValue (globalIDPropertyName, true, false, false);
    if (null == globalIDPropertyValue || globalIDPropertyValue.IsNull || string.IsNullOrEmpty (globalIDPropertyValue.StringValue))
        return null;

    return globalIDPropertyValue.StringValue;

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Updates the loop GUID.
/// </summary>
/// <param name="apComponent">The ap component.</param>
/// <author>Steve.Morrow</author>                                    <date>4/1/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public static void                  UpdateLoopGuid 
(
APPIDComponent apComponent
)
    {
    try
        {
        if (apComponent.OPPIDComponent == null || string.IsNullOrEmpty (apComponent.Keytag))
            return;

        Bom.IComponent component = apComponent.OPPIDComponent;
        string guid = BCT.ECPropertyValueUtility.GetString (component, "GUID", string.Empty);
        if (string.IsNullOrEmpty (guid))
            return;

        //guid = "{" + guid.ToUpper () + "}";
        //string sqlStatement = string.Format ("Update LOOP set KEYTAG_GUID_PK = {0} where KEYTAG = '{1}'", guid, apComponent.Keytag);
        
        string sqlStatement = string.Format ("Update LOOP set KEYTAG_GUID_PK = '{0}' where KEYTAG = '{1}'", guid.ToUpper (), apComponent.Keytag);
        ExecuteSqlStatement (sqlStatement);
        }
    catch { }

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ContainsDefintionBlock
/// </summary>
/// <param name="blockNameToSearchFor"></param> 
/// <author>Steve.Morrow</author>                                <date>09/2009</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              ContainsDefintionBlock
(
)
    {

    //Check to see if dwg contains the _at_asi_ldwg cell
    //This call deletes the xdata.. This is done so it would not be seen as a APPID document
    if (MasterAPPIDBlockdefinitionFound (out m_APPIDProjectID, out m_APPIDDocumentID, true))
        return true;

    return false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// is ElementID a Leader Line
/// </summary>
/// <param name="apComponent">appid component</param> 
/// <param name="elementID">element id</param> 
/// <returns>true is element id is leader id</returns> 
/// <author>Steve.Morrow</author>                            <date>09/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             isElementIDLeaderLine
(
APPIDComponent apComponent,
UInt32 elementID
)
    {
    if (null == apComponent.InstrumentBubbleLeaderHandles || apComponent.InstrumentBubbleLeaderHandles.Count == 0)
        return false;

    foreach (long elementLeader in apComponent.InstrumentBubbleLeaderHandles)
        {
        UInt32 ileader = (UInt32)elementLeader;
        if (ileader == elementID)
            return true;
        }

    return false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set Element Template From Parent Component
/// <param name="sourceComponent">Source OPPID component</param> 
/// <param name="targetComponent">Target OPPID component</param> 
/// </summary>
/// <author>Steve.Morrow</author>                                <date>01/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetElementTemplateFromParentComponent
(
Bom.IComponent sourceComponent,
Bom.IComponent targetComponent
)
    {
    IECInstance inst = sourceComponent.ClassDefinition.GetCustomAttributes ("SCHEMATICS_CAD_CUSTOM_ATTRIBUTES");
    if (null == inst)
        return;

    string template = BCT.ECPropertyValueUtility.GetString (inst, "ElementTemplate", String.Empty);
    if (string.IsNullOrEmpty (template))
        return;

    Bom.Workspace.ActiveModel.SetElementTemplateForComponent (targetComponent, template);
    }

#endregion

#region Schema Lists
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Matching Attribute Block TagString
/// Special case for valve size or any inline that has a tag.TagDefinitionName = PSZ
/// Need to create a schema definition, but for now hard code
/// The issue is that PSZ is stored on the attribute and not in any mapping. The PSZ value is from the PIPERUN
/// and in APPID that was filled out with a lisp function (valve_size). 
/// </summary>
/// <param name="fieldName">attribute tag Field name</param> 
/// <returns>OPPID matching property name</returns> 
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static string            GetMatchingAttributeBlockTagString
(
string fieldName
)
    {
    if (null == m_attributeBlockNameTagList)
        m_attributeBlockNameTagList = ConversionMappingUtilities.GetMappingListFromClassName (ConversionCatalog, "ATTRIBUTE_TEXT_MAPPING");

    if (null == m_attributeBlockNameTagList)
        {
        m_attributeBlockNameTagList = new SortedList ();
        return string.Empty;
        }

    if (!m_attributeBlockNameTagList.ContainsKey (fieldName))
        return string.Empty;

    return m_attributeBlockNameTagList[fieldName].ToString ();

    }
private static SortedList m_attributeBlockNameTagList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Use Alternate Target class
/// This is used for control valves AT_CVALVE. To get a specific OPPID class, the block name in the conversion schema has an alternate OPPID class name
/// </summary>
/// <<param name="tagType">tagType name</param> 
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              UseAlternateTargetClass
(
string tagType
)
    {

    if (null == m_tagTypesToUseAlternateClass)
        m_tagTypesToUseAlternateClass = ConversionMappingUtilities.GetStringCollection ("ALTERNATE_SEARCH_TAG_TYPES", ConversionCatalog);

    if (null != m_tagTypesToUseAlternateClass)
        {
        bool status = m_tagTypesToUseAlternateClass.Contains (tagType.ToUpper ());
        return status;
        }
    return false;
    }
private static ArrayList m_tagTypesToUseAlternateClass = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Ignore this cell name from the ap cell name 
/// This cell will not be considered for a schema match
/// </summary>
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IgnoreBlockName
(
string cellName
)
    {

    if (null == m_blockNames)
        m_blockNames = ConversionMappingUtilities.GetStringCollection ("BLOCK_ATTRIBUTE_NAMES", ConversionCatalog);

    if (null != m_blockNames)
        return m_blockNames.Contains (cellName.ToUpper ());

    return false;
    }
private static ArrayList m_blockNames = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
///  Is Field Name In Tag Registry Table
/// </summary>
/// <<param name="fieldName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsFieldNameInTagRegistryTable
(
string fieldName
)
    {

    if (null == m_tagRegistryTable)
        m_tagRegistryTable = ConversionMappingUtilities.GetStringCollection ("TAG_REGISTRY_COLUMN_NAMES", ConversionCatalog);

    if (null != m_tagRegistryTable)
        return m_tagRegistryTable.Contains (fieldName.ToUpper ());

    return false;
    }
private static ArrayList m_tagRegistryTable = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Do not convert the block names in this list
/// </summary>
/// <remarks>
/// These are the block names that will not be converted, but left in their native state. 
/// This includes actuator blocks on control valves. These block are greater than the extents of the inrun break. This causes the line break to be greater than the inrun size.
/// </remarks> 
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsBlockToBeConverted
(
string cellName
)
    {

    if (null == m_doNotConvertblockNames)
        m_doNotConvertblockNames = ConversionMappingUtilities.GetStringCollection ("DO_NOT_CONVERT_BLOCK_NAMES", ConversionCatalog);

    if (null != m_doNotConvertblockNames)
        return !m_doNotConvertblockNames.Contains (cellName.ToUpper ());

    return true;
    }
private static ArrayList m_doNotConvertblockNames = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is an actuator block, to be converted and used with control valves
/// </summary>
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>01/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsAnActuatorBlock
(
string cellName
)
    {

    if (null == m_actutorBlockNames)
        m_actutorBlockNames = ConversionMappingUtilities.GetStringCollection ("ACTUATOR_BLOCK_NAMES", ConversionCatalog);

    if (null != m_actutorBlockNames)
        return m_actutorBlockNames.Contains (cellName.ToUpper ());

    return false;
    }
private static ArrayList m_actutorBlockNames = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Determines whether [is an auto nozzle block] [the specified cell name].
/// </summary>
/// <param name="cellName">Name of the cell.</param>
/// <returns>
///   <c>true</c> if [is an auto nozzle block] [the specified cell name]; otherwise, <c>false</c>.
/// </returns>
/// <author>Steve.Morrow</author>                                    <date>3/1/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public static bool IsAutoNozzleBlock
(
string cellName
)
    {

    if (null == m_autoNozzleBlockNames)
        m_autoNozzleBlockNames = ConversionMappingUtilities.GetStringCollection ("AUTONOZZLE_CELLNAMES", ConversionCatalog);

    if (null != m_autoNozzleBlockNames)
        return m_autoNozzleBlockNames.Contains (cellName.ToUpper ());

    return false;
    }
private static ArrayList m_autoNozzleBlockNames = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Determines whether [is intelligent isolation valve] [the specified key name].
/// </summary>
/// <param name="keyName">Name of the key.</param>
/// <returns>
///   <c>true</c> if [is intelligent isolation valve] [the specified key name]; otherwise, <c>false</c>.
/// </returns>
/// <author>Steve.Morrow</author>                                   <date>12/30/2015</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public static bool IsIntelligentIsolationValve
(
string keyName
)
    {
    if (null == m_intelligentIsolationValve)
        m_intelligentIsolationValve = ConversionMappingUtilities.GetStringCollection("ISOLATION_VALVE_LIST", ConversionCatalog);

    if (null != m_intelligentIsolationValve)
        return m_intelligentIsolationValve.Contains(keyName.ToUpper());

    return false;
    }
private static ArrayList m_intelligentIsolationValve = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a multi line pipeline
/// </summary>
/// <<param name="pipelineCode">pipelineCode</param> 
/// <author>Ali.Aslam</author>                                <date>10/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsMultiLinePipeLineCode
(
string pipelineCode
)
    {
    if (null == pipelineCode)
        return false;

    if (null == m_MultiLinePipelineCodes)
        m_MultiLinePipelineCodes = ConversionMappingUtilities.GetStringCollection ("MULTI_LINE_PIPELINE_CODES", ConversionCatalog);

    if (null != m_MultiLinePipelineCodes)
        return m_MultiLinePipelineCodes.Contains (pipelineCode.ToUpper ());

    return false;
    }
private static ArrayList m_MultiLinePipelineCodes = null;
/*------------------------------------------------------------------------------------**/
/// <summary>
/// These block names have embedded attributes that will be ignored. 
/// This is the case for actuators with attributes.
/// </summary>
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>02/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IgnoreBlockNameWithTagSet
(
string cellName
)
    {

    if (null == m_ignoreBlockNameWithTagSet)
        m_ignoreBlockNameWithTagSet = ConversionMappingUtilities.GetStringCollection ("IGNORE_EMBEDDED_ATTRIBUTE_BLOCK_NAMES_LIST", ConversionCatalog);

    if (null != m_ignoreBlockNameWithTagSet)
        return m_ignoreBlockNameWithTagSet.Contains (cellName.ToUpper ());

    return false;
    }
private static ArrayList m_ignoreBlockNameWithTagSet = null;



/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get adjustment type
/// Instrument tags INUM, ITYP have a different justification. This requires a modified 
/// adjustment type to compensate for the justification.
/// </summary>
/// <param name="tag">tagset object</param> 
/// <returns>adjustment type</returns> 
/// <author>Steve.Morrow</author>                                <date>12/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static double            TagTextAdjustmentType
(
string fieldName
)
    {
    //default adjustment type is 1.0 This denotes to just use the standard offset
    double adjustmentType = 1.0;
    if (null == m_textAdjustmentList)
        m_textAdjustmentList = ConversionMappingUtilities.GetTextAdjustmentMapping
            ("APPID_TEXT_ADJUSTMENT_MAPPING", APPIDUtilities.ConversionCatalog);

    if (null == m_textAdjustmentList)
        m_textAdjustmentList = new SortedList ();

    if (null == m_textAdjustmentList || m_textAdjustmentList.Count == 0 ||
        !m_textAdjustmentList.ContainsKey (fieldName))
        return adjustmentType;

    string adjValue = m_textAdjustmentList[fieldName].ToString ();
    if (string.IsNullOrEmpty (adjValue))
        return adjustmentType;

    if (!double.TryParse (adjValue, out adjustmentType))
        return adjustmentType;

    return adjustmentType;
    }
private static SortedList m_textAdjustmentList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is A Non Intelligent InRun Block Name
/// These are the non-intelligent inline block names. This list is used to determine which block to use for the origin.
/// </summary>
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsANonIntelligentInRunBlockName
(
string cellName
)
    {

    if (null == m_nonIntelligentInRunBlockNames)
        m_nonIntelligentInRunBlockNames = ConversionMappingUtilities.GetStringCollection ("INLINE_NON_INTELLIGENT_BLOCK_NAMES", ConversionCatalog);

    if (null != m_nonIntelligentInRunBlockNames)
        {
        bool status = m_nonIntelligentInRunBlockNames.Contains (cellName.ToUpper ());
        return status;
        }

    return false;
    }
private static ArrayList m_nonIntelligentInRunBlockNames = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is A Instrument Bubble Annotation Block
/// These are the instrument block annotation block names. This list is used to determine which block to use for the origin when Metric scaling is donw.
/// </summary>
/// <<param name="cellName">cell name</param> 
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              IsAInstrumentBubbleAnnotationBlock
(
string cellName
)
    {

    if (null == m_instrumentBubbleAnnotationBlock)
        m_instrumentBubbleAnnotationBlock = ConversionMappingUtilities.GetStringCollection ("INSTRUMENT_BUBBLE_ANNOTATION_BLOCK_NAME", ConversionCatalog);

    if (null != m_instrumentBubbleAnnotationBlock)
        {
        bool status = m_instrumentBubbleAnnotationBlock.Contains (cellName.ToUpper ());
        return status;
        }

    return false;
    }
private static ArrayList m_instrumentBubbleAnnotationBlock = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Bubble tag types
/// The list of tag types that will use the intelligent OPPID bubble. Control valves and flow elements use an annotation bubble that is not the main tagged component.
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              SetBubbleTagTypes
(
)
    {
    APPIDBubbleTagTypeList = ConversionMappingUtilities.GetStringCollection ("BUBBLE_TAG_TYPES", APPIDUtilities.ConversionCatalog);
    if (null == APPIDBubbleTagTypeList || APPIDBubbleTagTypeList.Count == 0)
        {
        APPIDBubbleTagTypeList = new ArrayList ();
        APPIDBubbleTagTypeList.Add ("AT_INST_");
        APPIDBubbleTagTypeList.Add ("AT_TIEIN");
        APPIDBubbleTagTypeList.Add ("AT_SPEC_ITEM");
        }
    }
public static ArrayList APPIDBubbleTagTypeList = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Process Line Priority
/// In APPID the Pipeline class name is stored in the PIPE_RUN.PINIT1 column. 
/// There could be multiple Piperuns with multiple Pipeline class names (MAJOR,MIINOR,...) 
/// that belong to one Pipeline. This custom attribute can be set to determine which conversion class to use. 0 is the TOP priority.
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static int               GetProcessLinePriority
(
string APPIDClassName
)
    {
    int priority = -1;

    // Since it is possible for input parameter to contain dashes ('-') which is not premitted in a class name
    // we need to replace it with _ as this is the norm used in conversion schema. As an example the APPIDClassName
    // DIN-512 (for non combustible gases pipeline) should be converted to DIN_512 as it resides in conversion schema
    // with that name (DIN_512). No need to use further alternate checks (like starting with number) since the limited
    // set of possible autoplant pid pipelines only have - as a known case so far. 
    string appidClassNameToUse = APPIDClassName.Replace ("-", "_");;

    IECClass processClassName = ConversionMappingUtilities.GetECClassFromCatalog (ConversionCatalog, appidClassNameToUse);
    if (null == processClassName)
        return priority;

    IECInstance customAttributeInstance = processClassName.GetCustomAttributes ("PROCESSLINE_PRIORITY");
    if (null == customAttributeInstance)
        return priority;

    return BCT.ECPropertyValueUtility.GetInt (customAttributeInstance, "ORDER", priority);
    }


#endregion

#region Xdata
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Build Component List
/// </summary>
/// <author>Steve.Morrow</author>                                <date>00/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static Hashtable         BuildComponentList
(
)
    {

    //create list of associated items properties
    //APPIDProviderUtilities.CreateAssociatedItemsList (m_conversionCatalog);

    ////line annotation definition mapping
    //APPIDProviderUtilities.CreateLineNumberAnnotationMapping (m_conversionCatalog);

    //create list of instrument tag_types
    SetBubbleTagTypes ();

    m_APPIDComponents = new Hashtable ();
    MsDgn.Application msApp = MsDgnApplicationInstance;

    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();

    //*Scan all elements in drawing. 
    //One reason: need this because of equipment vessel heads are complex chains

    //scanCriteria.ExcludeAllTypes ();
    //scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.SharedCell);
    //scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.LineString);
    //scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.Line);
    //scanCriteria.IncludeType(Bentley.Interop.MicroStationDGN.MsdElementType.Arc);
    //scanCriteria.IncludeType(Bentley.Interop.MicroStationDGN.MsdElementType.ComplexShape);
    //scanCriteria.IncludeType(Bentley.Interop.MicroStationDGN.MsdElementType.ComplexString);
    //scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.Text);
    //scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.TextNode);

    MsDgn.ModelReference modelRef = msApp.ActiveModelReference;
    MsDgn.ElementEnumerator elemEnum = modelRef.Scan (scanCriteria);
    getXdataAndBuildList (elemEnum);

    foreach (MsDgn.Attachment attachment in msApp.ActiveModelReference.Attachments)
        {
        elemEnum = attachment.Scan (scanCriteria);
        getXdataAndBuildList (elemEnum);
        }

    return m_APPIDComponents;
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Project ID and Document ID
/// </summary>
/// <param name="projectID">Project ID</param> 
/// <param name="documentID">Document ID</param> 
/// <param name="deleteXdata">delete xdata</param> 
/// <author>Steve.Morrow</author>                                <date>00/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public static bool              MasterAPPIDBlockdefinitionFound
(
out string projectID,
out string documentID,
bool deleteXdata
)
    {
    projectID = string.Empty;
    documentID = string.Empty;
    string masterAPPIDBlock = "_at_asi_ldwg";
    MsDgn.Application msApp = MsDgnApplicationInstance;

    MsDgn.ElementScanCriteria scanCriteria = new MsDgn.ElementScanCriteriaClass ();
    scanCriteria.ExcludeAllTypes ();
    scanCriteria.IncludeOnlyCell (masterAPPIDBlock);
    scanCriteria.IncludeType (Bentley.Interop.MicroStationDGN.MsdElementType.SharedCell);

    MsDgn.ModelReference modelRef = msApp.ActiveModelReference;
    MsDgn.ElementEnumerator elemEnum = modelRef.Scan (scanCriteria);

    while (elemEnum.MoveNext ())
        {
        if (!elemEnum.Current.IsSharedCellElement ())
            return false;

        MsDgn.SharedCellElement masterCell = elemEnum.Current.AsSharedCellElement ();

        if (null == masterCell || !masterCell.IsGraphical || !masterCell.HasXData ("AT_") ||
            !masterCell.Name.Equals (masterAPPIDBlock, StringComparison.InvariantCultureIgnoreCase))
            return false;

        bool indicatorFound = false;
        MsDgn.XDatum[] xdatas = masterCell.GetXData ("AT_");
        if (null == xdatas)
            return false;

        foreach (MsDgn.XDatum xdata in xdatas)
            {
            switch (xdata.Type)
                {
                case MsDgn.MsdXDatumType.Int16:
                case MsDgn.MsdXDatumType.Int32:
                    //"1" is the indicator that next value (string) is the project and document ID
                    if (xdata.Value.ToString ().Equals ("1"))
                        indicatorFound = true;
                    break;
                case MsDgn.MsdXDatumType.String:
                    if (indicatorFound)
                        {
                        string[] docProj = xdata.Value.ToString ().Split (',');
                        projectID = docProj[0];
                        documentID = docProj[1];

                        if (deleteXdata)
                            {
                            masterCell.DeleteAllXData ();
                            masterCell.Rewrite ();
                            //masterCell.DeleteXData ("AT_");
                            }
                        return true;
                        }
                    break;
                }
            }
        }
    return false;
    }


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Extract parent and child element information from given line element. 
/// </summary>
/// <param name="lineElementID">Line element to analyze</param> 
/// <param name="mainParents17000">MAIN parent pipe run element which is usually a shared
/// cell element</param>
/// <param name="mainChildrenLinesIfMainPipeRunElement17001">Line element ids which are 
/// directly associated with main pipe run element. Usually these are only main pipe lines
/// but sometime include additional elements so need to be analyzed. Usually contain all
/// the starting lines for a pipe run</param>
/// <param name="immediateParentLineHandle17100">Immediate parent handle. Some lines are 
/// linked to main run element via an intermediary. This is usually case of multi lines
/// that appear after an in run component. They are usually associated with a main line 
/// after in run (and not directly to pipe run element)</param>
/// <param name="immediateChildrenLineHandles17101or17102">Immediate children of current
/// line element. This usually is present on main run after an in run only.</param>
///
/// <author>Ali.Aslam</author>                            <date>10/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
public static void              GetHierarchyInfoForLineElement
(
long lineElementID,
out List<long> mainParents17000,
out List<long> mainChildrenLinesIfMainPipeRunElement17001,
out List<long> immediateParentLineHandle17100,
out List<long> immediateChildrenLineHandles17101or17102
)
    {
    mainParents17000 = new List<long> ();
    mainChildrenLinesIfMainPipeRunElement17001 = new List<long> ();
    immediateParentLineHandle17100 = new List<long> ();
    immediateChildrenLineHandles17101or17102 = new List<long> ();

    MsDgn.Element msElement = APPIDUtilities.ElementFromElementID (lineElementID);
    if (null == msElement)
        return;

    // 17000 is main parent pipeline or pipe run element id depending on if its the main run component or a child line
    // For children of pipe run, it would remain same = appiperun.id for all child lines at every level
    // 17001 are main immediate children of pipe run, should appear only once on pipe run element id only
    // 17100 would be the immediate parent line handle in every child line at every level
    // 17101 and 17102 would be the children of a child line
    bool found17000 = false;
    bool found17001 = false;
    bool found17100 = false;
    bool found17101 = false;
    bool found17102 = false;

    MsDgn.XDataObject xdatas = msElement.GetXData1 ("AT_");
    if (null == xdatas)
        return;

    for (int i = 0; i < xdatas.Count; i++)
        {
        MsDgn.MsdXDatumType xdType = xdatas.GetXDatumType (i);
        switch (xdType)
            {
            case MsDgn.MsdXDatumType.WorldSpacePosition:
            case MsDgn.MsdXDatumType.WorldSpaceDisplacement:
            case MsDgn.MsdXDatumType.WorldDirection:
            case MsDgn.MsdXDatumType.Real:
            case MsDgn.MsdXDatumType.Unsupported:
                continue;

            case MsDgn.MsdXDatumType.Int16:
            case MsDgn.MsdXDatumType.Int32:
                MsDgn.XDatum xdataInt = xdatas.GetXDatum (i);
                if (null == xdataInt.Value)
                    continue;

                string typeVal = xdataInt.Value.ToString ();
                if (string.IsNullOrEmpty (typeVal))
                    continue;

                found17000 = false;
                found17001 = false;
                found17100 = false;
                found17101 = false;
                found17102 = false;
                if (typeVal.Equals ("-17000"))
                    found17000 = true;
                else if (typeVal.Equals ("-17001"))
                    found17001 = true;
                else if (typeVal.Equals ("-17100"))
                    found17100 = true;
                else if (typeVal.Equals ("-17101"))
                    found17101 = true;
                else if (typeVal.Equals ("-17102"))
                    found17102 = true;
                break;
            case MsDgn.MsdXDatumType.DatabaseHandle:
                MsDgn.XDatum xdata = xdatas.GetXDatum (i);
                if (null == xdata.Value)
                    continue;

                if (!found17000 && !found17001 && !found17100 && !found17101 && !found17102)
                    continue;

                string handle = xdata.Value.ToString ();
                if (string.IsNullOrEmpty (handle))
                    continue;

                long elementID = BmfUt.MiscellaneousUtilities.ConvertHexStringToLong (handle);
                if (elementID == 0)
                    continue;

                if (found17000)
                    {
                    if (!mainParents17000.Contains (elementID))
                        mainParents17000.Add (elementID);
                    }
                else if (found17001)
                    {
                    if (!mainChildrenLinesIfMainPipeRunElement17001.Contains (elementID))
                        mainChildrenLinesIfMainPipeRunElement17001.Add (elementID);
                    }
                else if (found17100)
                    {
                    if (!immediateParentLineHandle17100.Contains (elementID))
                        immediateParentLineHandle17100.Add (elementID);
                    }
                else if (found17101 || found17102)
                    {
                    if (!immediateChildrenLineHandles17101or17102.Contains (elementID))
                        immediateChildrenLineHandles17101or17102.Add (elementID);
                    }
                break;
            }
        }

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get Xdata And BuildList
/// </summary>
/// <param name="elemEnum">ElementEnumerator</param>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             getXdataAndBuildList
(
MsDgn.ElementEnumerator elemEnum
)
    {

    //Bentley.UI.Controls.WinForms.WaitForm.ShowForm
    //   (CTLog.FrameworkLogger.Instance.GetLocalizedString ("AutoPLANTPIDConversion"),
    //    CTLog.FrameworkLogger.Instance.GetLocalizedString ("AutoPLANTPIDConversionExtract"));

    try
        {
        while (elemEnum.MoveNext ())
            {

            AddToAdditionalXdataItems(elemEnum.Current);
            if (!elemEnum.Current.IsSharedCellElement())
                {                
                continue;
                }

            MsDgn.SharedCellElement msElement = elemEnum.Current.AsSharedCellElement ();
            if (msElement.Name.Equals("lnbrk"))
                {
                setAPComponentSpecBreak (msElement);
                continue;
                }
            if (null == msElement || !msElement.IsGraphical || !msElement.HasXData ("AT_LINK"))
                {
                setPID2Values(msElement);
                //non intelligent components: fittings
                if (IsANonIntelligentInRunBlockName (msElement.Name))
                    setNonIntelligentInRunAPComonentData (msElement);
                else
                    //create list of elements that are mapped but do not have xdata
                    //in the future these block types without xdata can be converted
                    displayInvalidElementMessage (msElement);

                continue;
                }

            //create a new entry for a APPID component
            setAPComponentData (msElement);
            }
        }
    catch { }
    //finally
    //    {
    //    Bentley.UI.Controls.WinForms.WaitForm.CloseForm ();
    //    }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Reads the pidset2 block and extracts the value that is stored in the SCALE attribute.
/// The Scale attribute is then used in scaling components that can be placed after conversion.
/// This Scale value is NOT used in the conversion
/// </summary>
/// <param name="msElement">The ms element.</param>
/// <author>Steve.Morrow</author>                                    <date>1/20/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
private static void setPID2Values(MsDgn.SharedCellElement msElement)
    {
    if (!msElement.Name.Equals("pidset2", StringComparison.InvariantCultureIgnoreCase))
        return;
    if (!msElement.HasAnyTags)
        return;
    MsDgn.TagElement[] tags = msElement.GetTags();
    if (null == tags)
        return;

    foreach (MsDgn.TagElement tag in tags)
        {
        if (tag == null || tag.Value == null)
            continue;

        string tagValue = tag.Value.ToString();
        if (string.IsNullOrEmpty(tagValue))
            continue;

        switch (tag.TagDefinitionName.ToLower())
            {
            case "scale":
                double scale = 1.0;
                if (double.TryParse(tagValue, out scale))
                    m_APPIDScale = scale;
                break;
            case "snap":
            case "sfi":
            case "brad":
            case "txth":
                break;
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets the APPID scale.
/// This is NOT used for converion
/// It is used for Placement from Task Menu
/// </summary>
/// <author>Steve.Morrow</author>                                <date>1/20/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public static double APPIDScale 
    {
    get
        {
        if (m_APPIDScale <= 0.0)
            m_APPIDScale = 1.0;

        return m_APPIDScale;
        }
    }
private static double m_APPIDScale = 1.0;    

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Adds to additional xdata items.
/// </summary>
/// <param name="elm">The elm.</param>
/// <author>Steve.Morrow</author>                                    <date>1/4/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
private static void AddToAdditionalXdataItems(MsDgn.Element elm)
    {

    try
        {
        //do not check any elements that have the intelligent 
        if (elm.HasXData("AT_LINK"))
            return;

        //must have a grouping xdata link
        if (!elm.HasXData("AT_"))
            return;

        MsDgn.XDataObject xdatas = elm.GetXData1("AT_");
        if (null == xdatas)
            return;

        //must contain only two items: type and value
        if (xdatas.Count != 2)
            return;

        if (AdditionalXdataItems == null)
            AdditionalXdataItems = new Dictionary<string, IList<long>>();

        //check xdata type. Must -17000
        MsDgn.MsdXDatumType xdataT = xdatas.GetXDatumType(0);
        switch (xdataT)
            {
            case Bentley.Interop.MicroStationDGN.MsdXDatumType.Int16:
            case Bentley.Interop.MicroStationDGN.MsdXDatumType.Int32:
                MsDgn.XDatum xdataInt = xdatas.GetXDatum(0);
                if (!xdataInt.Value.ToString().Equals("-17000"))
                    return;
                break;
            default:
                return;
            }

        //get element handle
        xdataT = xdatas.GetXDatumType(1);
        switch (xdataT)
            {
            case Bentley.Interop.MicroStationDGN.MsdXDatumType.DatabaseHandle:
                break;
            default:
                return;
            }
        MsDgn.XDatum xdataV = xdatas.GetXDatum(1);

        if (xdataV.Value == null)
            return;

        string elHandle = xdataV.Value.ToString();
        if (string.IsNullOrEmpty(elHandle))
            return;

        //convert handle from hex to long
        //add to collection. The key handle is the master element handle, which is defined on the xdata AT_LINK
        long xHandle = Convert.ToInt64(elHandle, 16);
        elHandle = xHandle.ToString();
        if (AdditionalXdataItems.ContainsKey(elHandle))
            {
            IList<long> lc = AdditionalXdataItems[elHandle];
            lc.Add(elm.ID64);
            }
        else
            {
            IList<long> lc = new List<long>();
            lc.Add(elm.ID64);
            AdditionalXdataItems.Add(elHandle, lc);
            }
        }
    catch 
        { 
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// AdditionalXdataItems 
/// Used for equipment trim with parametric equipment
/// </summary>
/// <author>Steve.Morrow</author>                                    <date>1/4/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public static IDictionary<string, IList<long>> AdditionalXdataItems = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Display message if the cell found is in the mapping but does not have any xdata
/// The Xdata would contain APPID information about grouping
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <author>Steve.Morrow</author>                                <date>03/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             displayInvalidElementMessage
(
MsDgn.SharedCellElement msElement
)
    {
    string cellName = msElement.Name.ToUpper ();
    IECClass cls = ConversionMappingUtilities.GetECClassFromCatalog (ConversionCatalog, cellName);
    if (null == cls)
        cls = ConversionMappingUtilities.GetECClassFromCatalog (ConversionCatalog, "AP_" + cellName);

    if (null == cls || msElement.HasXData ("AT_"))
        return;

    if (null == m_invalidOrMissingXdataList)
        m_invalidOrMissingXdataList = new StringCollection (); 

    m_invalidOrMissingXdataList.Add( string.Format
        (CTLog.FrameworkLogger.Instance.GetLocalizedString ("MissingXdataOnComponentBlock"),
        msElement.Name.ToUpper (), msElement.ID64.ToString ()));
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Build info about master APPID entity
/// xdata idenifers: -18001, -18002, -18005,
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             setNonIntelligentInRunAPComonentData
(
MsDgn.SharedCellElement msElement
)
    {

    APPIDComponent apComponent = new APPIDComponent ();

    //Set specific component values
    apComponent.ElementName = msElement.Name;
    apComponent.ID = msElement.ID;
    apComponent.APPIDClassName = msElement.Name.ToUpper ();
    apComponent.SourceData = new Hashtable ();
    apComponent.IsNonIntelligentInRun = true;

    apComponent.CellNames = new StringCollection ();
    apComponent.CellNames.Add (apComponent.APPIDClassName);

    string masterKey = msElement.Name + "-" + msElement.ID.ToString ();
    m_APPIDComponents.Add (masterKey, apComponent);
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Build info about Spec Break
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <author>Qarib.Kazmi</author>                                <date>02/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             setAPComponentSpecBreak
(
MsDgn.SharedCellElement msElement
)
    {
    APPIDComponent apComponent = new APPIDComponent ();
    string specClassName = BmfCat.CatalogServer.StripClassName (BmfPer.PersistenceManager.SpecBreakName);

    //Set specific component values
    apComponent.ElementName = msElement.Name;
    apComponent.ID = msElement.ID;
    apComponent.APPIDClassName = specClassName;
    apComponent.SourceData = new Hashtable ();
    apComponent.IsSpecBreak = true;

    apComponent.CellNames = new StringCollection ();
    apComponent.CellNames.Add ("Spec_Break");

    apComponent.OPPIDClassName = specClassName;
    string masterKey = msElement.Name + "-" + msElement.ID.ToString ();
    m_APPIDComponents.Add (masterKey, apComponent);
    }
/*------------------------------------------------------------------------------------**/
/// <summary>
/// Build info about master APPID entity
/// xdata idenifers: -18001, -18002, -18005,
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             setAPComponentData
(
MsDgn.SharedCellElement msElement
)
    {

    //AT_LINK are only on master entities. These are the parents.
    MsDgn.XDatum[] xdatas = msElement.GetXData ("AT_LINK");
    if (null == xdatas)
        return;

    string masterKey = string.Empty;

    APPIDComponent apComponent = new APPIDComponent ();

    //Set specific component values
    apComponent.ElementName = msElement.Name;
    apComponent.ID = msElement.ID;    
    apComponent.SourceData = getSourceData (msElement);

    //-18001 contains the link tab and link ID 
    //(1070 . -18001) (1000 . "EQP_LNK") (1000 . "0000000038")
    bool found18001 = false;

    //-18002 contains intelligent annotation handle 
    //(1070 . -18002) (1005 . "821")
    bool found18002 = false;

    //-18005 contains the keytag
    //(1070 . -18005) (1000 . "0000000375")
    bool found18005 = false;

    //Build info about master data point
    foreach (MsDgn.XDatum xdata in xdatas)
        {
        switch (xdata.Type)
            {
            case MsDgn.MsdXDatumType.Int16:
            case MsDgn.MsdXDatumType.Int32:
                string typeVal = xdata.Value.ToString ();
                found18001 = false;
                found18002 = false;
                found18005 = false;
                if (typeVal.Equals ("-18001"))
                    found18001 = true;
                else if (typeVal.Equals ("-18002"))
                    found18002 = true;
                else if (typeVal.Equals ("-18005"))
                    found18005 = true;
                break;
            case MsDgn.MsdXDatumType.String:
                string xdataValue = xdata.Value.ToString ();

                //This value should uniquely idenify a component
                //(1000 . "&PROC_LNK|0000000178")
                if (!found18001 && !found18002 && !found18005 && xdataValue.StartsWith ("&"))
                    {
                    masterKey = xdataValue;
                    continue;
                    }
                if (found18001)
                    {
                    if (string.IsNullOrEmpty (apComponent.LinkTab))
                        apComponent.LinkTab = xdataValue;
                    else
                        apComponent.KeyLink = xdataValue;
                    }
                else if (found18005)
                    apComponent.Keytag = xdataValue;
                break;
            case MsDgn.MsdXDatumType.DatabaseHandle:
                if (!found18002)
                    continue;
                if (null == apComponent.AnnotationHandles)
                    apComponent.AnnotationHandles = new ArrayList ();

                string handle = xdata.Value.ToString ();
                if (string.IsNullOrEmpty (handle))
                    continue;

                long elementID = BmfUt.MiscellaneousUtilities.ConvertHexStringToLong (handle);
                if (elementID == 0)
                    continue;

                if (!apComponent.AnnotationHandles.Contains (elementID))
                    apComponent.AnnotationHandles.Add (elementID);

                break;
            }
        }

    if (string.IsNullOrEmpty(masterKey) || null == apComponent)
        return;

    //get keytag from KEY_LINK table
    string keytag = getKeytagFromLinkValues (apComponent);
    if (!string.IsNullOrEmpty (keytag))
        apComponent.Keytag = keytag;

    //get tagtype from KEY_LINK table
    apComponent.TagType = getTagTypeFromLinkValues (apComponent, msElement);

    //set specific component types
    switch (apComponent.TagType)
        {
        case "AT_PROCESS":
            apComponent.IsPipeline = true;
            apComponent.APPIDClassName = assignPipelineClassFromPiperun (apComponent.Keytag);
            CTLog.FrameworkLogger.Instance.Submit (BsiLog.SEVERITY.DEBUG, "PINT1- " + apComponent.APPIDClassName);
            break;
        case "AT_PIPERUN":
            apComponent.IsPipeRun = true;
            break;
        case "AT_DWG_NAME":
            apComponent.IsBorderSheet = true;
            break;
        case "AT_PID_RUNTERM":
            apComponent.IsRunTerm = true;
            break;
        case "AT_PID_REDUCER":
            apComponent.IsReducer = true;
            break;
        case "AT_PSV":
            apComponent.IsPSV = true;
            break;
        case "AT_PID_NOZZLE":
            apComponent.IsNozzle = true;
            break;
        case "AT_PID_TOFROM":
            apComponent.IsOffpageConnector = true;
            break;
        case "AT_INST_":
            apComponent.IsInstrument = true;
            break;
        case "AT_TIEIN":
        case "AT_SPEC_ITEM":
            apComponent.IsBubbleIdentification = true;
            break;
        default:              
            if (apComponent.IsEquipment && 
                IsEquipmentInParametricList(apComponent))
                apComponent.IsParametricEquipment = true;
            break;
        }

    //Build info about related (child) entities
    setRelatedChildData (msElement, ref apComponent, false, false);

    checkEquipmentForAdditionalItems (ref apComponent);

    //set bubble origin if metric scaling
    setBubbleOriginPoint (ref apComponent);

    //Flow elements xdata is slightly different from Control valves and Instrument bubbles
    //Need to cleanup the xdata handles
    cleanUpFlowElementHandles (ref apComponent);

    // Check if pipe run is multi line and if so, set appropiate properties
    checkIfMultilinePipeRun (ref apComponent);

    if (!m_APPIDComponents.Contains (masterKey))
        m_APPIDComponents.Add (masterKey, apComponent);

    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Checks the equipment for additional items.
/// This will set the IsParametricEquipment flag
/// </summary>
/// <param name="apComponent">The ap component.</param>
/// <author>Steve.Morrow</author>                                   <date>3/3/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
private static void checkEquipmentForAdditionalItems (ref APPIDComponent apComponent)
    {
    //equipment is already Parametric no need to check to see if component is parametric
    if (!apComponent.IsEquipment || apComponent.IsParametricEquipment)
        return;

    //Check to see if an auto nozzle has bee placed on a equipment cell
    //the list contains the nozzle names and is stored in the conversion schema
    //if the nozzle is part of the APPID component, treat the component as parametric 
    if (apComponent.CellNames != null)
        {
        foreach (string cellName in apComponent.CellNames)
            {
            if (IsAutoNozzleBlock (cellName))
                {
                apComponent.IsParametricEquipment = true;
                return;
                }
            }
        }

    //check to see if equipment has extra lines and graphics
    //this is most often used with TEMA components
    //if there are extra APPID lines or graphics, treat the component as parametric 
    if (apComponent.ChildHandles != null)
        {
        int elementCnt = 0;
        foreach (long elementID in apComponent.ChildHandles)
            {
            UInt32 id = (UInt32)elementID;
            MsDgn.Element element = ElementFromElementUnit (id);
            if (null != element)
                {
                if (element.IsLineElement () 
                    || element.IsArcElement () 
                    || element.IsEllipseElement() )
                    elementCnt++;
                }
            if (elementCnt > 0)
                {
                apComponent.IsParametricEquipment = true;
                return;
                }
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Check if pipe run is based off a pipeline which uses a multiline template
/// </summary>
/// <remarks>Method sets the IsMultiLinePipeRun property of component based on the 
/// pipeline code</remarks>
/// <param name="apPipeRunComponent">appid component which should be pipe run</param> 
/// 
/// <author>Ali.Aslam</author>                            <date>10/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             checkIfMultilinePipeRun
(
ref APPIDComponent apPipeRunComponent
)
    {
    if ((null == apPipeRunComponent) ||
        !apPipeRunComponent.IsPipeRun ||
        (null == apPipeRunComponent.SourceData))
        return;

    if (apPipeRunComponent.SourceData.ContainsKey("PINT1"))
        {
        string pipeLineCode = apPipeRunComponent.SourceData["PINT1"].ToString().Trim().ToUpper ();
        apPipeRunComponent.IsMultiLinePipeRun = IsMultiLinePipeLineCode (pipeLineCode);
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Flow elements xdata is slightly different from Control valves and Instrument bubbles
/// Need to cleanup the xdata handles
/// </summary>
/// <param name="apComponent">appid component</param> 
/// <author>Steve.Morrow</author>                            <date>11/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             cleanUpFlowElementHandles
(
ref APPIDComponent apComponent
)
    {
    if (null == apComponent.InstrumentBubbleHandles || apComponent.InstrumentBubbleHandles.Count == 0 ||
        null == apComponent.ChildHandles || apComponent.ChildHandles.Count == 0)
        return;

    foreach (long elementID in apComponent.InstrumentBubbleHandles)
        {
        if (apComponent.ChildHandles.Contains(elementID))
            apComponent.ChildHandles.Remove(elementID);
        }

    if (null == apComponent.InstrumentBubbleLeaderHandles || apComponent.InstrumentBubbleLeaderHandles.Count == 0)
        return;

    foreach (long elementID in apComponent.InstrumentBubbleLeaderHandles)
        {
        if (apComponent.ChildHandles.Contains (elementID))
            apComponent.ChildHandles.Remove (elementID);
        }
    } 

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set Bubble Origin point is metric scaling was used
/// </summary>
/// <param name="apComponent">appid component</param> 
/// <author>Steve.Morrow</author>                            <date>11/2011</date>
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             setBubbleOriginPoint
(
ref APPIDComponent apComponent
)
    {

    if (!APPIDUtilities.MetricScalingIsBeingUsed)
        return;

    if (null == apComponent.InstrumentBubbleHandles || apComponent.InstrumentBubbleHandles.Count == 0 ||
        null == apComponent.ChildHandles || apComponent.ChildHandles.Count == 0)
        return;

    //get all handles, so a search for instrument block attributes
    ArrayList elementIDs = new ArrayList ();
    UInt32 id = (UInt32)apComponent.ID;
    elementIDs.Add (id);
    foreach (long elementID in apComponent.ChildHandles)
        {
        id = (UInt32)elementID;

        if (!elementIDs.Contains (id))
            elementIDs.Add (id);

        }

    //go thru all elements and get center point from attibute block
    foreach (UInt32 uid in elementIDs)
        {
        MsDgn.Element element = APPIDUtilities.ElementFromElementUnit (uid);
        if (null == element)
            continue;

        if (!element.IsSharedCellElement ())
            continue;
        MsDgn.SharedCellElement cell = element.AsSharedCellElement ();

        string name = cell.Name.ToUpper();
        if (!IsAInstrumentBubbleAnnotationBlock (name))
            continue;

        apComponent.InstrumentBubbleCenterPoint = cell.Origin; 
        return;
        }
    } 

/*------------------------------------------------------------------------------------**/
/// <summary>
/// check xdata for 16007 bubble indicator
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <returns>true if 16007 is found</returns> 
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             hasBubbleData
(
MsDgn.Element msElement
)
    {

    MsDgn.XDataObject xdatas = msElement.GetXData1 ("AT_");
    if (null == xdatas)
        return false;

    for (int i = 0; i < xdatas.Count; i++)
        {

        MsDgn.MsdXDatumType xdType = xdatas.GetXDatumType (i);
        switch (xdType)
            {
            case MsDgn.MsdXDatumType.Int16:
            case MsDgn.MsdXDatumType.Int32:
                MsDgn.XDatum xdataInt = xdatas.GetXDatum (i);
                if (xdataInt.Value.ToString ().Equals ("-16007"))
                    return true;
                break;
            }
        }
    return false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// check xdata for Equipment UnderLine
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <returns>true if 17000 is found</returns> 
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static bool             hasEquipmentUnderLineData
(
MsDgn.Element msElement
)
    {

    MsDgn.XDataObject xdatas = msElement.GetXData1 ("AT_");
    if (null == xdatas)
        return false;

    for (int i = 0; i < xdatas.Count; i++)
        {

        MsDgn.MsdXDatumType xdType = xdatas.GetXDatumType (i);
        switch (xdType)
            {
            case MsDgn.MsdXDatumType.Int16:
            case MsDgn.MsdXDatumType.Int32:
                MsDgn.XDatum xdataInt = xdatas.GetXDatum (i);
                if (xdataInt.Value.ToString ().Equals ("-17000"))
                    return true;
                break;
            }
        }
    return false;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Build info about related (child) entities
/// xdata idenifers: -17000, -17001
/// </summary>
/// <param name="msElement">SharedCellElement</param>
/// <param name="apComponent">APPIDComponent</param>
/// <param name="bubbleDataOnly">Extract bubble data only</param> 
/// <param name="bubbleDataOnly">Extract equipment underline data only</param> 
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static void             setRelatedChildData
(
MsDgn.SharedCellElement msElement,
ref APPIDComponent apComponent,
bool annotationBubbleData,
bool annotationEquipUnderLineData
)
    {
    //-17000 contains the parent handle (process line would equal = 832)
    //(1070 . -17000) (1005 . "832")
    bool found17000 = false;

    //-17001 contains the child handles
    //(1070 . -17001) (1005 . "821") (1005 . "829")
    bool found17001 = false;

    bool found19001 = false;

    //Instrument bubble leader line
    bool found17004 = false;

    //instrument bubble info
    bool found16007 = false;

    MsDgn.XDataObject xdatas = msElement.GetXData1 ("AT_");
    if (null == xdatas)
        return;

    for (int i = 0; i < xdatas.Count; i++)
        {
        MsDgn.MsdXDatumType xdType = xdatas.GetXDatumType (i);
        switch (xdType)
            {
            case MsDgn.MsdXDatumType.WorldSpacePosition:
                if (found16007)
                    apComponent.InstrumentBubbleCenterPoint = xdatas.GetPoint3d (i);
                break;
            case MsDgn.MsdXDatumType.WorldSpaceDisplacement:
            case MsDgn.MsdXDatumType.WorldDirection:
            case MsDgn.MsdXDatumType.Real:
            case MsDgn.MsdXDatumType.Unsupported:
                continue;
            case MsDgn.MsdXDatumType.Int16:
            case MsDgn.MsdXDatumType.Int32:
                MsDgn.XDatum xdataInt = xdatas.GetXDatum (i);
                string typeVal = xdataInt.Value.ToString ();
                found17000 = false;
                found17001 = false;
                found17004 = false;
                found16007 = false;
                if (typeVal.Equals ("-17000"))
                    found17000 = true;
                else if (typeVal.Equals ("-17001"))
                    found17001 = true;
                else if (typeVal.Equals("-19001"))
                    found19001 = true;
                else if (typeVal.Equals ("-17004"))
                    found17004 = true;
                else if (typeVal.Equals ("-16007"))
                    found16007 = true;
                break;
            case MsDgn.MsdXDatumType.DatabaseHandle:
                MsDgn.XDatum xdata = xdatas.GetXDatum (i);
                if (null == xdata.Value)
                    continue;

                if (!found17000 && !found17001 && !found17004 && !found16007 && !found19001)
                    continue;

                string handle = xdata.Value.ToString ();
                if (string.IsNullOrEmpty (handle))
                    continue;

                long elementID = BmfUt.MiscellaneousUtilities.ConvertHexStringToLong (handle);
                if (elementID == 0)
                    continue;

                if (found17001)
                    {
                    if (annotationBubbleData)
                        {
                        if (null == apComponent.AdditionalHandles)
                            apComponent.AdditionalHandles = new ArrayList();

                        //used for function symbol relations
                        if (!apComponent.AdditionalHandles.Contains(elementID)) 
                            apComponent.AdditionalHandles.Add(elementID); 
                        continue;
                        }

                    //special case: get underline handles
                    //this is used to delete the underline after the text is replaced
                    if (annotationEquipUnderLineData)
                        {
                        //only equipment has this, don't want to mix any other tag types 
                        if (apComponent.IsPipeline || apComponent.IsPipeRun)
                            continue;

                        if (null == apComponent.EquipmentUnderLineHandles)
                            apComponent.EquipmentUnderLineHandles = new ArrayList ();

                        if (elementID != 0 && elementID != apComponent.ID &&
                            !apComponent.EquipmentUnderLineHandles.Contains (elementID)
                            )
                            {
                            UInt32 id = (UInt32)elementID;
                            apComponent.EquipmentUnderLineHandles.Add (id);
                            }
                        continue;
                        }

                    //standard xdata
                    if (null == apComponent.ChildHandles)
                        apComponent.ChildHandles = new ArrayList ();

                    if (elementID != 0 && !apComponent.ChildHandles.Contains (elementID))
                        {
                        apComponent.ChildHandles.Add (elementID);

                        //check to see if element is of SharedCellElement type
                        MsDgn.Element element = ElementFromElementID (elementID);
                        if (null != element && element.IsSharedCellElement ())
                            {
                            if (null == apComponent.CellNames)
                                apComponent.CellNames = new StringCollection ();

                            //get the cell name
                            //This name can be used to determine the specific OPPID class name
                            //Example: All valves use the AT_HVALVE  tagtype.. But the cell names are specific.
                            //         These cell names can match: 31200 = GATE_VALVE
                            string cellName = element.AsSharedCellElement ().Name.ToUpper ();

                            //insulation is stored in the same xdata space as the lines. Need to remove them
                            if (apComponent.IsPipeline || apComponent.IsPipeRun)
                                {
                                apComponent.ChildHandles.Remove (elementID);
                                continue;
                                }

                            //check to see if the block name is to be converted.
                            if (!IsBlockToBeConverted (cellName))
                                {
                                apComponent.ChildHandles.Remove (elementID);
                                continue;
                                }

                            //check to see if the block name is an actuator
                            if (IsAnActuatorBlock (cellName))
                                {
                                APPIDComponent actuator = new APPIDComponent ();
                                actuator.APPIDClassName = cellName;
                                actuator.ID = elementID;
                                actuator.ChildHandles = new ArrayList ();
                                actuator.ChildHandles.Add (elementID);
                                actuator.CellNames = new StringCollection ();
                                actuator.CellNames.Add (cellName.ToUpper());
                                //Assign Actuator cell Rotation Matrix to get Actuator Angle
                                actuator.CellRotationMatrix = GetRotationfromSharedCell(elementID);
                                apComponent.Actuator = actuator;
                                continue;
                                }

                            if (!string.IsNullOrEmpty (cellName) &&
                                !IgnoreBlockName (cellName) && !apComponent.CellNames.Contains (cellName))
                                apComponent.CellNames.Add (cellName);

                            //need to get instrument bubble on control valve
                            if (hasBubbleData (element))
                                setRelatedChildData (element.AsSharedCellElement (), ref apComponent, true, false);
                            else if (hasEquipmentUnderLineData (element))
                                setRelatedChildData (element.AsSharedCellElement (), ref apComponent, false, true);

                            }
                        }
                    }
                else if (found19001)
                    {
                    MsDgn.Element element = ElementFromElementID (elementID);
                    try
                        {
                        if (null != element && element.IsSharedCellElement())
                            apComponent.APPIDEndConditionName = element.AsSharedCellElement().Name.ToUpper();
                        }
                    catch { }
                    }
                else if (found17000)
                    {
                    //instrument bubble data (not using yet)
                    if (annotationBubbleData)
                        {
                        if (null == apComponent.InstrumentBubbleHandles)
                            apComponent.InstrumentBubbleHandles = new ArrayList();

                        if (elementID != 0 && elementID != apComponent.ID &&
                            !apComponent.InstrumentBubbleHandles.Contains(elementID)
                            )
                            apComponent.InstrumentBubbleHandles.Add(elementID);
                        }
                    else
                        {
                        //standard xdata
                        if (null == apComponent.ParentHandles)
                            apComponent.ParentHandles = new ArrayList();

                        if (elementID != 0 && !apComponent.ParentHandles.Contains(elementID))
                            apComponent.ParentHandles.Add(elementID);
                        }
                    }
                else if (found17004)
                    {
                    //instrument bubble leader line
                    if (null == apComponent.InstrumentBubbleLeaderHandles)
                        apComponent.InstrumentBubbleLeaderHandles = new ArrayList();
                    if (elementID != 0 && !apComponent.InstrumentBubbleLeaderHandles.Contains(elementID))
                        apComponent.InstrumentBubbleLeaderHandles.Add(elementID);
                    }
                else if (found16007)
                    {
                    //instrument bubble data (not using yet)
                    if (null == apComponent.InstrumentBubbleHandles)
                        apComponent.InstrumentBubbleHandles = new ArrayList();

                    if (elementID != 0 && !apComponent.InstrumentBubbleHandles.Contains(elementID))
                        apComponent.InstrumentBubbleHandles.Add(elementID);

                    }
                break;
            }
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// get TagType
/// </summary>
/// <param name="apComponent">appid component object</param> 
/// <param name="sharedCell">shared cell</param> 
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static string           getTagTypeFromLinkValues
(
APPIDComponent apComponent,
MsDgn.SharedCellElement sharedCell
)
    {

    string tagType = string.Empty;
    string sqlStatement = string.Format ("select TAG_TYPE from KEY_LINK where LINK_ID = '{0}' and LINK_TAB = '{1}'", apComponent.KeyLink, apComponent.LinkTab);
    tagType = "Tag"; // getSqlStringValue (sqlStatement, "TAG_TYPE");

    if(!string.IsNullOrEmpty(tagType))
        return tagType;

    if (!sharedCell.HasXData ("AT_SOURCEDATA"))
        return tagType;

    MsDgn.XDatum[] xdatas = sharedCell.GetXData ("AT_SOURCEDATA");
    bool found = false;
    if (null == xdatas)
        return tagType;
    foreach (MsDgn.XDatum xdata in xdatas)
        {
        switch (xdata.Type)
            {
            case MsDgn.MsdXDatumType.String:
                if (found)
                    return xdata.Value.ToString ();
                if (null == xdata.Value)
                    break;
                if (xdata.Value.ToString ().Equals ("^TAG_TYPE") ||
                    xdata.Value.ToString ().Equals ("TAG_TYPE"))
                    found = true;
                break;
            }
        }

    return tagType;
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Get Source data from Xdata
/// </summary>
/// <param name="elemEnum">ElementEnumerator</param>
/// <author>Steve.Morrow</author>                                <date>00/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
private static Hashtable        getSourceData
(
MsDgn.SharedCellElement sharedCell
)
    {

    Hashtable sourceData = new Hashtable ();
    if (!sharedCell.HasXData ("AT_SOURCEDATA"))
        return sourceData;

    MsDgn.XDatum[] xdatas = sharedCell.GetXData ("AT_SOURCEDATA");

    if (null == xdatas)
        return null;

    bool startFound = false;
    string key = string.Empty;
    object val = string.Empty;

    foreach (MsDgn.XDatum xdata in xdatas)
        {
        switch (xdata.Type)
            {
            case MsDgn.MsdXDatumType.String:
                if (startFound)
                    {
                    if (string.IsNullOrEmpty (key))
                        key = xdata.Value.ToString ();
                    else
                        val = xdata.Value;
                    }
                break;
            case MsDgn.MsdXDatumType.ControlString:
                startFound = false;
                if (xdata.Value.ToString ().Equals ("{"))
                    {
                    startFound = true;
                    key = string.Empty;
                    val = null;
                    }
                else if (xdata.Value.ToString ().Equals ("}"))
                    {
                    if (!string.IsNullOrEmpty(key) && !sourceData.ContainsKey (key))
                        sourceData.Add (key, val);

                    key = string.Empty;
                    val = null;
                    }
                break;
            }
        }

    return sourceData;
    }

#endregion

}
    
/*====================================================================================**/
/// <summary>
/// APPID Component object
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*==============+===============+===============+===============+===============+=====**/
public class                    APPIDComponent
{

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Keytag
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   Keytag
    {
    get
        {
        return m_keytag;
        }
    set
        {
        m_keytag = value;
        }
    }
private string m_keytag = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// LinkTab
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   LinkTab
    {
    get
        {
        return m_linkTab;
        }
    set
        {
        m_linkTab = value;
        }
    }
private string m_linkTab = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// KeyLink
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   KeyLink
    {
    get
        {
        return m_keyLink;
        }
    set
        {
        m_keyLink = value;
        }
    }
private string m_keyLink = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// TagType
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   TagType
    {
    get
        {
        return m_tagType;
        }
    set
        {
        m_tagType = value;
        }
    }
private string m_tagType = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// OPPID ClassName
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   OPPIDClassName
    {
    get
        {
        return m_OPPIDClassName;
        }
    set
        {
        m_OPPIDClassName = value;
        }
    }
private string m_OPPIDClassName = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// APPID ClassName
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   APPIDClassName
    {
    get
        {
        return m_APPIDClassName;
        }
    set
        {
        m_APPIDClassName = value;
        }
    }
private string m_APPIDClassName = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// TextStyle
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   TextStyle
    {
    get
        {
        return m_textStyle;
        }
    set
        {
        m_textStyle = value;
        }
    }
private string m_textStyle = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ElementName
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public string                   ElementName
    {
    get
        {
        return m_elementName;
        }
    set
        {
        m_elementName = value;
        }
    }
private string m_elementName = string.Empty;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// AnnotationHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                AnnotationHandles
    {
    get
        {
        return m_annotationHandles;
        }
    set
        {
        m_annotationHandles = value;
        }
    }
private ArrayList m_annotationHandles = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Annotation ElementIDs
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                AnnotationElementIDs
    {
    get
        {
        return m_annotationElementIDs;
        }
    set
        {
        m_annotationElementIDs = value;
        }
    }
private ArrayList m_annotationElementIDs = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// EquipmentUnderLineHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                EquipmentUnderLineHandles
    {
    get
        {
        return m_equipmentUnderLineHandles;
        }
    set
        {
        m_equipmentUnderLineHandles = value;
        }
    }
private ArrayList m_equipmentUnderLineHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ChildHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                ChildHandles
    {
    get
        {
        return m_childHandles;
        }
    set
        {
        m_childHandles = value;
        }
    }
private ArrayList m_childHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ParentHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                ParentHandles
    {
    get
        {
        return m_parentHandles;
        }
    set
        {
        m_parentHandles = value;
        }
    }
private ArrayList m_parentHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Cell names
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public StringCollection         CellNames
    {
    get
        {
        return m_cellNames;
        }
    set
        {
        m_cellNames = value;
        }
    }
private StringCollection m_cellNames = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// SourceData
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public Hashtable                SourceData
    {
    get
        {
        return m_sourceData;
        }
    set
        {
        m_sourceData = value;
        }
    }
private Hashtable m_sourceData = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Property Mapping
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public SortedList               PropertyMapping
    {
    get
        {
        return m_propertyMapping;
        }
    set
        {
        m_propertyMapping = value;
        }
    }
private SortedList m_propertyMapping = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Text Style Override Mapping
/// </summary>
/// <author>Steve.Morrow</author>                                <date>01/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public SortedList               TextStyleOverrideMapping
    {
    get
        {
        return m_textStyleOverrideMapping;
        }
    set
        {
        m_textStyleOverrideMapping = value;
        }
    }
private SortedList m_textStyleOverrideMapping = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ID of master component
/// </summary>
/// <author>Steve.Morrow</author>                                <date>07/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public long                     ID
    {
    get
        {
        return m_ID;
        }
    set
        {
        m_ID = value;
        }
    }
private long m_ID = 0;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Convert Annotation
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     ConvertAnnotationText
    {
    get
        {
        return m_convertAnnotationText;
        }
    set
        {
        m_convertAnnotationText = value;
        }
    }
private bool m_convertAnnotationText = true;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Set Active AssociatedItem Relationship
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     SetActiveAssociatedItemRelationship
    {
    get
        {
        return m_setActiveAssociatedItemRelationship;
        }
    set
        {
        m_setActiveAssociatedItemRelationship = value;
        }
    }
private bool m_setActiveAssociatedItemRelationship = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is In Run (valve, cvalve, flow element)
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsInRun
    {
    get
        {
        return m_isInRun;
        }
    set
        {
        m_isInRun = value;
        }
    }
private bool m_isInRun = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Pipeline. An APPID Process line
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsPipeline
    {
    get
        {
        return m_isPipeline;
        }
    set
        {
        m_isPipeline = value;
        }
    }
private bool m_isPipeline = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Pipe Run
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsPipeRun
    {
    get
        {
        return m_isPipeRun;
        }
    set
        {
        m_isPipeRun = value;
        }
    }
private bool m_isPipeRun = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a multi line Pipe Run
/// </summary>
/// <author>Ali.Aslam</author>                                <date>10/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsMultiLinePipeRun
    {
    get
        {
        return m_isMultiLinePipeRun;
        }
    set
        {
        m_isMultiLinePipeRun = value;
        }
    }
private bool m_isMultiLinePipeRun = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Nozzle
/// </summary>
/// <author>Steve.Morrow</author>                                <date>01/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsNozzle
    {
    get
        {
        return m_isNozzle;
        }
    set
        {
        m_isNozzle = value;
        }
    }
private bool m_isNozzle = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Reducer
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsReducer
    {
    get
        {
        return m_isAReducer;
        }
    set
        {
        m_isAReducer = value;
        }
    }
private bool m_isAReducer = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a RunTerm
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsRunTerm
    {
    get
        {
        return m_isARunTerm;
        }
    set
        {
        m_isARunTerm = value;
        }
    }
private bool m_isARunTerm = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Border sheet or Model type of component
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsBorderSheet
    {
    get
        {
        return m_isBorderSheet;
        }
    set
        {
        m_isBorderSheet = value;
        }
    }
private bool m_isBorderSheet = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is an Offpage Connector
/// </summary>
/// <author>Steve.Morrow</author>                                <date>02/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsOffpageConnector
    {
    get
        {
        return m_isOffpageConnector;
        }
    set
        {
        m_isOffpageConnector = value;
        }
    }
private bool m_isOffpageConnector = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a PSV
/// </summary>
/// <author>Steve.Morrow</author>                                <date>12/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsPSV
    {
    get
        {
        return m_isPSV;
        }
    set
        {
        m_isPSV = value;
        }
    }
private bool m_isPSV = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is an Instrument
/// </summary>
/// <author>Naveed.Khalid</author>                                <date>08/2014</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsInstrument
    {
    get
        {
        return m_isInstrument;
        }
    set
        {
        m_isInstrument = value;
        }
    }
private bool m_isInstrument = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Bubble Identification. 
/// This denotes a TieIn Or Specialty Item
/// </summary>
/// <author>Steve.Morrow</author>                                <date>01/2016</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool IsBubbleIdentification
    {
    get
        {
        return m_isBubbleIdentification;
        }
    set
        {
        m_isBubbleIdentification = value;
        }
    }
private bool m_isBubbleIdentification = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is an Instrument loop
/// </summary>
/// <author>Naveed.Khalid</author>                                <date>08/2014</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsInstrumentLoop
    {
    get
        {
        return m_isInstrumentLoop;
        }
    set
        {
        m_isInstrumentLoop = value;
        }
    }
private bool m_isInstrumentLoop = false;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is Non Intelligent InRun
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsNonIntelligentInRun
    {
    get
        {
        return m_isNonIntelligentInRun;
        }
    set
        {
        m_isNonIntelligentInRun = value;
        }
    }
private bool m_isNonIntelligentInRun = false;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Is a Spec Break
/// </summary>
/// <author>Qarib.Kazmi</author>                                <date>02/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     IsSpecBreak
    {
    get
        {
        return m_isSpecBreak;
        }
    set
        {
        m_isSpecBreak = value;
        }
    }
private bool m_isSpecBreak = false;
/*------------------------------------------------------------------------------------**/
/// <summary>
/// APPID Element Deleted
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     APPIDElementDeleted
    {
    get
        {
        return m_APPIDElementDeleted;
        }
    set
        {
        m_APPIDElementDeleted = value;
        }
    }
private bool m_APPIDElementDeleted = false;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets a value indicating whether this instance is parametric equipment.
/// </summary>
/// <value>
/// 	<c>true</c> if this instance is parametric equipment; otherwise, <c>false</c>.
/// </value>
/// <author>Steve.Morrow</author>                                    <date>1/4/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public bool IsParametricEquipment
    {
    get
        {
        return m_isParametricEquipment;
        }
    set
        {
        m_isParametricEquipment = value;
        }
    }
private bool m_isParametricEquipment = false;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets a value indicating whether this instance is equipment.
/// </summary>
/// <value>
/// 	<c>true</c> if this instance is equipment; otherwise, <c>false</c>.
/// </value>
/// <author>Steve.Morrow</author>                                    <date>1/4/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public bool IsEquipment
    {
    get
        {
        if (TagType.StartsWith ("AT_EQ") || TagType.StartsWith ("AT_MOT") ||
            LinkTab.StartsWith ("EQP_LNK "))
            return true;

        return false;
        }
    }

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Replace APPID Block with new OPPID Cell
/// </summary>
/// <remarks>
/// If this flag is true (set in the conversion schema on the class), then the old APPID cell is erased and replaced with the OPPID defined cell name (from the supplemental)
/// </remarks> 
/// <author>Steve.Morrow</author>                                <date>12/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     ReplaceWithNewOPPIDCell
    {
    get
        {
        return m_replaceWithNewOPPIDCell;
        }
    set
        {
        m_replaceWithNewOPPIDCell = value;
        }
    }
private bool m_replaceWithNewOPPIDCell = true;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Use alternate method to create the cell. If in this list, the addElementToPrimitive() method is used instead of createCellContainer
/// </summary>
/// <author>Steve.Morrow</author>                                <date>02/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public bool                     UseAlternateCellCreation
    {
    get
        {
        return m_useAlternateCellCreation;
        }
    set
        {
        m_useAlternateCellCreation = value;
        }
    }
private bool m_useAlternateCellCreation = false;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// OPPIDComponent
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public Bom.IComponent           OPPIDComponent
    {
    get
        {
        return m_oppidComponent;
        }
    set
        {
        m_oppidComponent = value;
        }
    }
private Bom.IComponent m_oppidComponent = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// InstrumentBubbleHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                InstrumentBubbleHandles
    {
    get
        {
        return m_instrumentBubbleHandles;
        }
    set
        {
        m_instrumentBubbleHandles = value;
        }
    }
private ArrayList m_instrumentBubbleHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// InstrumentBubbleLeaderHandles
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public ArrayList                InstrumentBubbleLeaderHandles
    {
    get
        {
        return m_instrumentBubbleLeaderHandles;
        }
    set
        {
        m_instrumentBubbleLeaderHandles = value;
        }
    }
private ArrayList m_instrumentBubbleLeaderHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// InstrumentBubbleCenterPoint
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2010</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public MsDgn.Point3d            InstrumentBubbleCenterPoint
    {
    get
        {
        return m_instrumentBubbleCenterPoint;
        }
    set
        {
        m_instrumentBubbleCenterPoint = value;
        }
    }
private MsDgn.Point3d m_instrumentBubbleCenterPoint = new MsDgn.Point3d ();

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Cell Origin
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IPoint             CellOrigin
    {
    get
        {
        return m_cellOrigin;
        }
    set
        {
        m_cellOrigin = value;
        }
    }
private BmfGe.IPoint m_cellOrigin = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Cell Scale
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IPoint             CellScale
    {
    get
        {
        return m_cellScale;
        }
    set
        {
        m_cellScale = value;
        }
    }
private BmfGe.IPoint m_cellScale = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Cell Range
/// </summary>
/// <author>Steve.Morrow</author>                                <date>09/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IRange             CellRange
    {
    get
        {
        return m_cellRange;
        }
    set
        {
        m_cellRange = value;
        }
    }
private BmfGe.IRange m_cellRange = null;


/*------------------------------------------------------------------------------------**/
/// <summary>
/// Cell Rotation Matrix
/// </summary>
/// <author>Steve.Morrow</author>                                <date>08/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IRotMatrix         CellRotationMatrix
    {
    get
        {
        return m_cellRotationMatrix;
        }
    set
        {
        m_cellRotationMatrix = value;
        }
    }
private BmfGe.IRotMatrix m_cellRotationMatrix = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ConnectionPoint CP1
/// </summary>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IPoint             ConnectionPtCP1
    {
    get
        {
        return m_connectionPtCP1;
        }
    set
        {
        m_connectionPtCP1 = value;
        }
    }
private BmfGe.IPoint m_connectionPtCP1 = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// ConnectionPoint CP2
/// </summary>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IPoint             ConnectionPtCP2
    {
    get
        {
        return m_connectionPtCP2;
        }
    set
        {
        m_connectionPtCP2 = value;
        }
    }
private BmfGe.IPoint m_connectionPtCP2 = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Direction CP1
/// </summary>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IVector            DirectionCP1
    {
    get
        {
        return m_directionCP1;
        }
    set
        {
        m_directionCP1 = value;
        }
    }
private BmfGe.IVector m_directionCP1 = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Direction CP2
/// </summary>
/// <author>Steve.Morrow</author>                                <date>10/2011</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public BmfGe.IVector            DirectionCP2
    {
    get
        {
        return m_directionCP2;
        }
    set
        {
        m_directionCP2 = value;
        }
    }
private BmfGe.IVector m_directionCP2 = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Actuator
/// </summary>
/// <author>Steve.Morrow</author>                                <date>01/2012</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public APPIDComponent           Actuator
    {
    get
        {
        return m_Actuator;
        }
    set
        {
        m_Actuator = value;
        }
    }
private APPIDComponent m_Actuator = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Component to be rotated by Angle specified in Conversion Mapping
/// </summary>
/// <author>Qarib.Kazmi</author>                                <date>12/2015</date> 
/*--------------+---------------+---------------+---------------+---------------+------*/
public double                     AngleToRotate
    {
    get
        {
        return m_angleToRotate;
        }
    set
        {
        m_angleToRotate = value;
        }
    }
private double m_angleToRotate = 0.0;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets the origin offset.
/// </summary>
/// <value>
/// The origin offset.
/// </value>
/// <author>Steve.Morrow</author>                               <date>12/22/2015</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public BmfGe.IPoint                 OriginOffset
    {
    get
        {
        return m_originOffset;
        }
    set
        {
        m_originOffset = value;
        }
    }
private BmfGe.IPoint m_originOffset = new BmfGe.Point (0,0,0);

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets the warning class.
/// </summary>
/// <value>
/// The warning class.
/// </value>
/// <author>Steve.Morrow</author>
/// <date>1/15/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public IECInstance WarningClassInstance
    {
    get
        {
        return m_warningClassInstance;
        }
    set
        {
        m_warningClassInstance = value;
        }
    }
private IECInstance m_warningClassInstance = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets the additional handles.
/// </summary>
/// <value>
/// The additional handles.
/// </value>
/// <author>Steve.Morrow</author>                                   <date>1/21/2016</date>
/// /*--------------+---------------+---------------+---------------+---------------+------*/ 
public ArrayList AdditionalHandles
    {
    get
        {
        return m_AdditionalHandles;
        }
    set
        {
        m_AdditionalHandles = value;
        }
    }
private ArrayList m_AdditionalHandles = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets the name of the APPID end condition.
/// </summary>
/// <value>
/// The name of the APPID end condition.
/// </value>
/// <author>Steve.Morrow</author>                               <date>1/27/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public string APPIDEndConditionName
    {
    get
        {
        return m_APPIDEndConditionName;
        }
    set
        {
        m_APPIDEndConditionName = value;
        }
    }
private string m_APPIDEndConditionName = null;

/*------------------------------------------------------------------------------------**/
/// <summary>
/// Gets or sets the reducer runs.
/// </summary>
/// <value>
/// The reducer runs.
/// </value>
/// <author>Steve.Morrow</author>                               <date>2/29/2016</date>
/*--------------+---------------+---------------+---------------+---------------+------*/ 
public ArrayList ReducerRuns
    {
    get
        {
        return m_reducerRuns;
        }
    set
        {
        m_reducerRuns = value;
        }
    }
private ArrayList m_reducerRuns = null;

}
}
