﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace PGSolutions.BetterRibbon.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("PGSolutions.BetterRibbon.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Analyze Current WB.
        /// </summary>
        internal static string AnalyzeLinksCurrent_Label {
            get {
                return ResourceManager.GetString("AnalyzeLinksCurrent_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Analyze Links in Current WB.
        /// </summary>
        internal static string AnalyzeLinksCurrent_ScreenTip {
            get {
                return ResourceManager.GetString("AnalyzeLinksCurrent_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Provides an analysis of all (non-Table) external links found in this WB in either cell formulas or named references.
        ///
        ///Should be idempotent, as the output worksheets (ie &quot;Links Errors&quot;, &quot;Linked Files&quot; and &quot;Links Analysis&quot;) are omitted from the analysis.
        /// </summary>
        internal static string AnalyzeLinksCurrent_SuperTip {
            get {
                return ResourceManager.GetString("AnalyzeLinksCurrent_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Analyze Listed WBs.
        /// </summary>
        internal static string AnalyzeLinksSelected_Label {
            get {
                return ResourceManager.GetString("AnalyzeLinksSelected_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Analyze Links in Listed WBs.
        /// </summary>
        internal static string AnalyzeLinksSelected_ScreenTip {
            get {
                return ResourceManager.GetString("AnalyzeLinksSelected_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Provides an analysis of all (non-Table) external links found in this WB in either cell formulas or named references.
        ///
        ///All valid WB names in the currently selected Range are analyzed.
        ///
        ///Always performed in a separate Excel instance to avoid file name conflicts with open workbooks..
        /// </summary>
        internal static string AnalyzeLinksSelected_SuperTip {
            get {
                return ResourceManager.GetString("AnalyzeLinksSelected_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Background Mode.
        /// </summary>
        internal static string BackgroundModeToggle_Label {
            get {
                return ResourceManager.GetString("BackgroundModeToggle_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Toggles Background-Mode .
        /// </summary>
        internal static string BackgroundModeToggle_ScreenTip {
            get {
                return ResourceManager.GetString("BackgroundModeToggle_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to When pressed, uses a separate background instance of Excel for analysis of a file list, instead of the default instance. 
        ///
        ///This is a beta feature, not yet fully tested..
        /// </summary>
        internal static string BackgroundModeToggle_SuperTip {
            get {
                return ResourceManager.GetString("BackgroundModeToggle_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to A Better Ribbon.
        /// </summary>
        internal static string BrandingButton_Label {
            get {
                return ResourceManager.GetString("BrandingButton_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Pretty (Darn) Good Solutions.
        /// </summary>
        internal static string BrandingButton_ScreenTip {
            get {
                return ResourceManager.GetString("BrandingButton_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Simple robust solutions to complex problems, by Pieter Geerkens
        ///
        ///https://github.com/pgeerkens/PGSolutions.BetterRibbon.
        /// </summary>
        internal static string BrandingButton_SuperTip {
            get {
                return ResourceManager.GetString("BrandingButton_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to PGSolutions.
        /// </summary>
        internal static string BrandingGroup_Label {
            get {
                return ResourceManager.GetString("BrandingGroup_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export Current.
        /// </summary>
        internal static string CurrentProjectButtonMS_Label {
            get {
                return ResourceManager.GetString("CurrentProjectButtonMS_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export VBA from Current Workbook.
        /// </summary>
        internal static string CurrentProjectButtonMS_ScreenTip {
            get {
                return ResourceManager.GetString("CurrentProjectButtonMS_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code from the current Excel workbook..
        /// </summary>
        internal static string CurrentProjectButtonMS_SuperTip {
            get {
                return ResourceManager.GetString("CurrentProjectButtonMS_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export Current.
        /// </summary>
        internal static string CurrentProjectButtonPG_Label {
            get {
                return ResourceManager.GetString("CurrentProjectButtonPG_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export VBA from Current Workbook.
        /// </summary>
        internal static string CurrentProjectButtonPG_ScreenTip {
            get {
                return ResourceManager.GetString("CurrentProjectButtonPG_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code from the current Excel workbook..
        /// </summary>
        internal static string CurrentProjectButtonPG_SuperTip {
            get {
                return ResourceManager.GetString("CurrentProjectButtonPG_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Controls Customizable in VBA.
        /// </summary>
        internal static string CustomizableGroup_Label {
            get {
                return ResourceManager.GetString("CustomizableGroup_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Application-Customizable Controls.
        /// </summary>
        internal static string CustomizableGroup_ScreenTip {
            get {
                return ResourceManager.GetString("CustomizableGroup_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to These controls can be dynamically customized, both in behaviour and appearance, within the VBA code for a workbook. The controls automatically deactivate when a workbook unaware of the customizations receives focus, and re-activate again when an aware workbook receives focus..
        /// </summary>
        internal static string CustomizableGroup_SuperTip {
            get {
                return ResourceManager.GetString("CustomizableGroup_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to External Links Analysis.
        /// </summary>
        internal static string LinksAnalysisGroup_Label {
            get {
                return ResourceManager.GetString("LinksAnalysisGroup_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to External Links Analyzer.
        /// </summary>
        internal static string LinksAnalysisGroup_ScreenTip {
            get {
                return ResourceManager.GetString("LinksAnalysisGroup_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Reports on all external links detected in either the current workbook, or the list of workbooks in column 1 of the currently selected Range.
        ///
        ///Both cell formulas and Named Ranges are searched..
        /// </summary>
        internal static string LinksAnalysisGroup_SuperTip {
            get {
                return ResourceManager.GetString("LinksAnalysisGroup_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap PGeerkens {
            get {
                object obj = ResourceManager.GetObject("PGeerkens", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;?xml version=&quot;1.0&quot; encoding=&quot;UTF-8&quot;?&gt;
        ///&lt;!-- Copyright 2018-2019 Pieter Geerkens --&gt;
        ///&lt;!-- When debugging, remember to check: Options -&gt; Advanced -&gt; General -&gt; Show_Add-In_user_interface_errors. --&gt;
        ///&lt;mso:customUI xmlns:mso=&quot;http://schemas.microsoft.com/office/2009/07/customui&quot;
        ///              xmlns:rd=&quot;BetterRibbon&quot;
        ///              onLoad=&quot;OnRibbonLoad&quot; loadImage=&quot;Ribbon_LoadImage&quot;&gt;
        ///  &lt;mso:ribbon&gt;
        ///    &lt;mso:tabs&gt;
        ///        &lt;mso:tab idMso=&quot;TabDeveloper&quot; &gt;
        ///            &lt;mso:group id=&quot;VbaExportGroupMS&quot; getVisib [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string Ribbon {
            get {
                return ResourceManager.GetString("Ribbon", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export Selected.
        /// </summary>
        internal static string SelectedProjectButtonMS_Label {
            get {
                return ResourceManager.GetString("SelectedProjectButtonMS_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export from Selected Workbook(s).
        /// </summary>
        internal static string SelectedProjectButtonMS_ScreenTip {
            get {
                return ResourceManager.GetString("SelectedProjectButtonMS_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code from selected Excel workbook(s)..
        /// </summary>
        internal static string SelectedProjectButtonMS_SuperTip {
            get {
                return ResourceManager.GetString("SelectedProjectButtonMS_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export Selected.
        /// </summary>
        internal static string SelectedProjectButtonPG_Label {
            get {
                return ResourceManager.GetString("SelectedProjectButtonPG_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Export from Selected Workbook(s).
        /// </summary>
        internal static string SelectedProjectButtonPG_ScreenTip {
            get {
                return ResourceManager.GetString("SelectedProjectButtonPG_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code from selected Excel workbook(s)..
        /// </summary>
        internal static string SelectedProjectButtonPG_SuperTip {
            get {
                return ResourceManager.GetString("SelectedProjectButtonPG_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Use &apos;SRC&apos; folder.
        /// </summary>
        internal static string UseSrcFolderToggleMS_Label {
            get {
                return ResourceManager.GetString("UseSrcFolderToggleMS_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Toggles use of folder-name SRC.
        /// </summary>
        internal static string UseSrcFolderToggleMS_ScreenTip {
            get {
                return ResourceManager.GetString("UseSrcFolderToggleMS_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to By default VBAExport is to a directory &apos;SRC&apos; sibling to the workbook. and Multi-Select is disabled..
        ///
        ///When this toggle is raised the export is instead to a directory eponymous with the workbook, suffixed by VBA. and Multi-Select is enabled..
        /// </summary>
        internal static string UseSrcFolderToggleMS_SuperTip {
            get {
                return ResourceManager.GetString("UseSrcFolderToggleMS_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Use &apos;SRC&apos; folder.
        /// </summary>
        internal static string UseSrcFolderTogglePG_Label {
            get {
                return ResourceManager.GetString("UseSrcFolderTogglePG_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Toggles use of folder-name SRC.
        /// </summary>
        internal static string UseSrcFolderTogglePG_ScreenTip {
            get {
                return ResourceManager.GetString("UseSrcFolderTogglePG_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to By default VBAExport is to a directory &apos;SRC&apos; sibling to the workbook. and Multi-Select is disabled..
        ///
        ///When this toggle is raised the export is instead to a directory eponymous with the workbook, suffixed by VBA. and Multi-Select is enabled..
        /// </summary>
        internal static string UseSrcFolderTogglePG_SuperTip {
            get {
                return ResourceManager.GetString("UseSrcFolderTogglePG_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to VBA by PGSolutions.
        /// </summary>
        internal static string VbaExportGroupMS_Label {
            get {
                return ResourceManager.GetString("VbaExportGroupMS_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to VBA Source Export Controls.
        /// </summary>
        internal static string VbaExportGroupMS_ScreenTip {
            get {
                return ResourceManager.GetString("VbaExportGroupMS_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code to a sibling directory of each workbook, by module type:
        ///
        ///Class modules as .CLS
        ///Standard modules as .VBA
        ///MSForm modules as .FRM
        ///
        ///Project references as .XML.
        /// </summary>
        internal static string VbaExportGroupMS_SuperTip {
            get {
                return ResourceManager.GetString("VbaExportGroupMS_SuperTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to VBA Source Code Export.
        /// </summary>
        internal static string VbaExportGroupPG_Label {
            get {
                return ResourceManager.GetString("VbaExportGroupPG_Label", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to VBA Source Export Controls.
        /// </summary>
        internal static string VbaExportGroupPG_ScreenTip {
            get {
                return ResourceManager.GetString("VbaExportGroupPG_ScreenTip", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to Exports all VBA code to a sibling directory of each workbook, by module type:
        ///
        ///Class modules as .CLS
        ///Standard modules as .VBA
        ///MSForm modules as .FRM
        ///
        ///Project references as .XML.
        /// </summary>
        internal static string VbaExportGroupPG_SuperTip {
            get {
                return ResourceManager.GetString("VbaExportGroupPG_SuperTip", resourceCulture);
            }
        }
    }
}
