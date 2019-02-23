////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    //internal interface IRibbonViewModel: IRibbonViewModelCom {

    //}
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonViewModel)]
    public interface IRibbonViewModel {
        /// <summary>The Control ID of the Ribbon definition being dispatched by this instance.</summary>
        [Description("The Control ID of the Ribbon definition being dispatched by this instance.")]
        string Id { get; }

        /// <summary>.</summary>
        IRibbonUI RibbonUI { get; }

        /// <summary>Loads an image, making it accessible by name to ribbon controls via an 'image' tag.</summary>
        [Description("Loads an image, making it accessible by name to ribbon controls via an 'image' tag.")]
        object LoadImage(string ImageId);

        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        [Description("Call back for GetDescription events from ribbon elements.")]
        string GetDescription(IRibbonControl Control);

        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        [Description("Call back for GetEnabled events from ribbon elements.")]
        bool GetEnabled(IRibbonControl Control);

        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        [Description("Call back for GetKeyTip events from ribbon elements.")]
        string GetKeyTip(IRibbonControl Control);

        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        [Description("Call back for GetLabel events from ribbon elements.")]
        string GetLabel(IRibbonControl Control);

        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        [Description("Call back for GetScreenTip events from ribbon elements.")]
        string GetScreenTip(IRibbonControl Control);

        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        [Description("Call back for GetSize events from ribbon elements.")]
        bool   GetSize(IRibbonControl Control);

        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        [Description("Call back for GetSuperTip events from ribbon elements.")]
        string GetSuperTip(IRibbonControl Control);

        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        [Description("Call back for GetVisible events from ribbon elements.")]
        bool GetVisible(IRibbonControl Control);

        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        [Description("Call back for GetImage events from ribbon elements.")]
        object GetImage(IRibbonControl Control);

        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        [Description("Call back for GetShowImage events from ribbon elements.")]
        bool GetShowImage(IRibbonControl Control);

        /// <summary>Call back for GetShowLabe l events from ribbon elements.</summary>
        [Description("Call back for GetShowLabe l events from ribbon elements.")]
        bool GetShowLabel(IRibbonControl Control);

        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        [Description("Call back for GetPressed events from the checkBox and toggleButton ribbon elements.")]
        bool GetPressed(IRibbonControl Control);

        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        [Description("Call back for OnAction events from the checkBox and toggleButton ribbon elements.")]
        void OnActionToggle(IRibbonControl Control, bool IsPressed);

        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        [Description("Call back for OnAction events from the button ribbon elements.")]
        void OnAction(IRibbonControl Control);

        /// <summary>Returns the ID of the currently selected list item.</summary>
        /// <remarks>This callback is typically used only for list controls with a dynamic list.</remarks>
        [Description("Returns the ID of the currently selected list item.\nThis callback is typically used only for list controls with a dynamic list.")]
        string GetSelectedItemId(IRibbonControl Control);

        /// <summary>Returns the index of the currently selected list item.</summary>
        /// <remarks>This callback is typically used only for list controls with a static list.</remarks>
        [Description("Returns the index of the currently selected list item.\nThis callback is typically used only for list controls with a static list.")]
        int GetSelectedItemIndex(IRibbonControl Control);

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [Description("Call back for OnAction events from the drop-down ribbon elements.")]
        void OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [Description("Call back for ItemCount events from the drop-down ribbon elements.")]
        int GetItemCount(IRibbonControl Control);

        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemID events from the drop-down ribbon elements.")]
        string GetItemId(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemLabel events from the drop-down ribbon elements.")]
        string GetItemLabel(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemScreenTip events from the drop-down ribbon elements.")]
        string GetItemScreenTip(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        string GetItemSuperTip(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemImage events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemImage events from the drop-down ribbon elements.")]
        object GetItemImage(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemShowImage events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemShowImage events from the drop-down ribbon elements.")]
        bool GetItemShowImage(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemShowLabel events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemShowLabel events from the drop-down ribbon elements.")]
        bool GetItemShowLabel(IRibbonControl Control, int Index);
    }
}
