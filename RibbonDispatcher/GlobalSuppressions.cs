////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;

// This file is used by Code Analysis to maintain SuppressMessage 
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given 
// a specific target and scoped to a namespace, type, member, etc.
//
// To add a suppression to this file, right-click the message in the 
// Code Analysis results, point to "Suppress Message", and click 
// "In Suppression File".
// You do not need to add suppressions to this file manually.

[assembly: SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures", Scope = "member",
        Target = "PGSolutions.RibbonDispatcher.Models.ControlModel`2.#.ctor(System.Func`2<System.String,PGSolutions.RibbonDispatcher.ComInterfaces.IActivatable`2<!0,!1>>,PGSolutions.RibbonDispatcher.ComInterfaces.IControlStrings,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Design","CA1006:DoNotNestGenericTypesInMemberSignatures",Scope = "member",
        Target = "PGSolutions.RibbonDispatcher.Models.ControlModel`2.#.ctor(System.Func`2<System.String,PGSolutions.RibbonDispatcher.ViewModels.IActivatable`2<!0,!1>>,PGSolutions.RibbonDispatcher.IControlStrings)")]
[assembly: SuppressMessage("Microsoft.Design","CA1006:DoNotNestGenericTypesInMemberSignatures",Scope = "member",
        Target = "PGSolutions.RibbonDispatcher.Models.AbstractSplitButtonModel`2.#.ctor(System.Func`2<System.String,PGSolutions.RibbonDispatcher.ViewModels.IActivatable`2<!0,!1>>,PGSolutions.RibbonDispatcher.IControlStrings2,PGSolutions.RibbonDispatcher.Models.MenuModel)")]
[assembly: SuppressMessage("Microsoft.Design","CA1006:DoNotNestGenericTypesInMemberSignatures",Scope = "member",
        Target = "PGSolutions.RibbonDispatcher.Models.ControlModel2`2.#.ctor(System.Func`2<System.String,PGSolutions.RibbonDispatcher.ViewModels.IActivatable`2<!0,!1>>,PGSolutions.RibbonDispatcher.IControlStrings2)")]

#region Irrelevant as all COM types passed by interface, not type
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.GroupModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.ToggleButtonVM")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.DropDownModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.DropDownVM")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.ToggleModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.SelectableItem")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.ButtonVM")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.ButtonModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.CheckBoxVM")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.SelectableItemModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.ComboBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.EditBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.LabelControlModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.MenuModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.SplitToggleButtonModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible", Scope = "type", Target = "PGSolutions.RibbonDispatcher.Models.SplitPressButtonModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.DynamicMenuModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.GalleryModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.MenuSeparatorModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticComboBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticDropDownModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1405:ComVisibleTypeBaseTypesShouldBeComVisible",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticGalleryModel")]
#endregion

#region All these types are factoyr created, then exported by interface.
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.ButtonModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.ComboBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.DropDownModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.EditBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.ResourceLoader")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.SelectableItemModel")]
[assembly: SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable", Scope = "type", Justification = "Public, Non-Creatable, class with exported Events.", Target = "PGSolutions.RibbonDispatcher.Models.ToggleModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.GalleryModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticComboBoxModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticDropDownModel")]
[assembly: SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable",Scope = "type",Target = "PGSolutions.RibbonDispatcher.Models.StaticGalleryModel")]
#endregion

#region COM Event Handlers are non-standard
[assembly: SuppressMessage("Microsoft.Design","CA1003:UseGenericEventHandlerInstances",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.ClickedEventHandler2")]
[assembly: SuppressMessage("Microsoft.Design","CA1003:UseGenericEventHandlerInstances",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.ToggledEventHandler2")]
[assembly: SuppressMessage("Microsoft.Design","CA1003:UseGenericEventHandlerInstances",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.SelectionMadeEventHandler2")]
[assembly: SuppressMessage("Microsoft.Design","CA1003:UseGenericEventHandlerInstances",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.EditedEventHandler2")]

[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IButtonModel.#Clicked")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IComboBoxModel.#Edited")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IDropDownModel.#SelectionMade")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IEditBoxModel.#Edited")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IGalleryModel.#SelectionMade")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IStaticComboBoxModel.#Edited")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IStaticDropDownModel.#SelectionMade")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IStaticGalleryModel.#SelectionMade")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IToggleModel.#Toggled")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractSelectableModel`2.#SelectionMade")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ViewModels.IDynamicMenuVM.#GetContent")]
[assembly: SuppressMessage("Microsoft.Design","CA1009:DeclareEventHandlersCorrectly",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ViewModels.IDynamicMenuVM.#ContentLoaded")]
#endregion

#region External Constraints on Event Signatures
[assembly: SuppressMessage("Microsoft.Design","CA1021:AvoidOutParameters",MessageId = "1#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ViewModels.IDynamicMenuVM.#OnGetContent(Microsoft.Office.Core.IRibbonControl,System.String&)")]

[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "1#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ViewModels.ContentEventHandler.#Invoke(Microsoft.Office.Core.IRibbonControl,System.String&)")]
[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "2#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "1#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "1#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.DynamicMenuModel.#OnGetContent(Microsoft.Office.Core.IRibbonControl,System.String&)")]

[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "2#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Design","CA1045:DoNotPassTypesByReference",MessageId = "1#",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]

[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Cancel",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "SaveAsUI",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Cancel",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "wb",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Success",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_AfterSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean)")]

[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "wb",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_Deactivate(Microsoft.Office.Interop.Excel.Workbook)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Cancel",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "SaveAsUI",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "wb",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_BeforeSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Success",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_AfterSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "Cancel",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]
[assembly: SuppressMessage("Microsoft.Usage","CA1801:ReviewUnusedParameters",MessageId = "wb",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_Close(Microsoft.Office.Interop.Excel.Workbook,System.Boolean&)")]
#endregion

#region "strings" is actually the most descriptive parameter name
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewGroupModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewButtonModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewToggleModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewEditBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewDropDownModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewStaticDropDownModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewComboBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewStaticComboBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewLabelControlModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewMenuModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewGalleryModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewStaticGalleryModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewMenuSeparatorModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewDynamicMenuModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewSplitToggleButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.ModelFactory.#NewSplitPressButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]

[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelServer.#GetSplitToggleButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelServer.#GetSplitPressButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewGroupModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewButtonModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewToggleModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewEditBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewDropDownModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewStaticDropDownModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewComboBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewStaticComboBoxModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewLabelControlModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewMenuModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewGalleryModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewStaticGalleryModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewMenuSeparatorModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewDynamicMenuModel(System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewSplitToggleButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Naming","CA1720:IdentifiersShouldNotContainTypeNames",MessageId = "strings",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ComInterfaces.IModelFactory.#NewSplitPressButtonModel(System.String,System.String,System.String,System.Boolean,System.Boolean)")]
#endregion

[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.IGroupVM")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.ITabVM")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.ILabelControlVM")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.IBoxControlVM")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.IButtonGroupVM")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.IBoxControlSource")]
[assembly: SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces",Scope = "type",Target = "PGSolutions.RibbonDispatcher.ViewModels.IButtonGroupSource")]

[assembly: SuppressMessage("Microsoft.Design","CA1011:ConsiderPassingBaseTypesAsParameters",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_Activate(Microsoft.Office.Interop.Excel.Workbook)")]
[assembly: SuppressMessage("Microsoft.Design","CA1011:ConsiderPassingBaseTypesAsParameters",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.AbstractCustomDispatcher.#Workbook_AfterSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Design","CA1011:ConsiderPassingBaseTypesAsParameters",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_Activate(Microsoft.Office.Interop.Excel.Workbook)")]
[assembly: SuppressMessage("Microsoft.Design","CA1011:ConsiderPassingBaseTypesAsParameters",Scope = "member",Target = "PGSolutions.RibbonDispatcher.Models.CustomDispatcher.#Workbook_AfterSave(Microsoft.Office.Interop.Excel.Workbook,System.Boolean)")]
[assembly: SuppressMessage("Microsoft.Design","CA1011:ConsiderPassingBaseTypesAsParameters",Scope = "member",Target = "PGSolutions.RibbonDispatcher.ViewModels.ViewModelFactory.#ParseXmlDoc(System.Xml.Linq.XElement)")]
