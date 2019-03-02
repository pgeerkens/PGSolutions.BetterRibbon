////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The complete set of Ribbon Callbacks supported by this implementation.</summary>
    [Description("The complete set of Ribbon Callbacks supported by this implementation.")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ICallbackDispatcher)]
    public interface ICallbackDispatcher {
        /// <summary>Loads an image, making it accessible by name to ribbon controls via an 'image' tag.</summary>
        [DispId( 1), Description("Loads an image, making it accessible by name to ribbon controls via an 'image' tag.")]
        object LoadImage(string ImageId);

        #region IDescriptionableVM implementation
        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        [DispId( 2), Description("Call back for GetDescription events from ribbon elements.")]
        string GetDescription(IRibbonControl Control);
        #endregion

        #region IControlVM implementation
        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        [DispId( 3), Description("Call back for GetEnabled events from ribbon elements.")]
        bool GetEnabled(IRibbonControl Control);

        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        [DispId( 4), Description("Call back for GetKeyTip events from ribbon elements.")]
        string GetKeyTip(IRibbonControl Control);

        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        [DispId( 5), Description("Call back for GetLabel events from ribbon elements.")]
        string GetLabel(IRibbonControl Control);

        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        [DispId( 6), Description("Call back for GetScreenTip events from ribbon elements.")]
        string GetScreenTip(IRibbonControl Control);

        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        [DispId( 7), Description("Call back for GetSuperTip events from ribbon elements.")]
        string GetSuperTip(IRibbonControl Control);

        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        [DispId( 8), Description("Call back for GetVisible events from ribbon elements.")]
        bool GetVisible(IRibbonControl Control);
        #endregion

        #region ISizeableVM implementation
        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        [DispId( 9), Description("Call back for GetSize events from ribbon elements.")]
        bool GetSize(IRibbonControl Control);
        #endregion

        #region IImageableVM implementation
        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        [DispId(10), Description("Call back for GetImage events from ribbon elements.")]
        object GetImage(IRibbonControl Control);

        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        [DispId(11), Description("Call back for GetShowImage events from ribbon elements.")]
        bool GetShowImage(IRibbonControl Control);

        /// <summary>Call back for GetShowLabe l events from ribbon elements.</summary>
        [DispId(12), Description("Call back for GetShowLabe l events from ribbon elements.")]
        bool GetShowLabel(IRibbonControl Control);
        #endregion

        #region IClickableVM implementation
        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        [DispId(13), Description("Call back for OnAction events from the button ribbon elements.")]
        void OnAction(IRibbonControl Control);
        #endregion

        #region IToggleableVM implementation
        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(14), Description("Call back for GetPressed events from the checkBox and toggleButton ribbon elements.")]
        bool GetPressed(IRibbonControl Control);

        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(15), Description("Call back for OnAction events from the checkBox and toggleButton ribbon elements.")]
        void OnActionToggle(IRibbonControl Control, bool IsPressed);
        #endregion

        #region ISelectableVM implementation - DropDown & ComboBox
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(16), Description("Call back for ItemCount events from the drop-down ribbon elements.")]
        int GetItemCount(IRibbonControl Control);

        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(17), Description("Call back for GetItemID events from the drop-down ribbon elements.")]
        string GetItemId(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemImage events from the drop-down ribbon elements.</summary>
        [DispId(18), Description("Call back for GetItemImage events from the drop-down ribbon elements.")]
        object GetItemImage(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(19), Description("Call back for GetItemLabel events from the drop-down ribbon elements.")]
        string GetItemLabel(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(20), Description("Call back for GetItemScreenTip events from the drop-down ribbon elements.")]
        string GetItemScreenTip(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(21), Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        string GetItemSuperTip(IRibbonControl Control, int Index);
        #endregion

        #region ISelectable2VM implementation - DropDown
        /// <summary>Returns the ID of the currently selected list item.</summary>
        /// <remarks>This callback is typically used only for list controls with a dynamic list.</remarks>
        [DispId(22), Description("Returns the ID of the currently selected list item.\nThis callback is typically used only for list controls with a dynamic list.")]
        string GetSelectedItemId(IRibbonControl control);

        /// <summary>Returns the index of the currently selected list item.</summary>
        /// <remarks>This callback is typically used only for list controls with a static list.</remarks>
        [DispId(23), Description("Returns the index of the currently selected list item.\nThis callback is typically used only for list controls with a static list.")]
        int GetSelectedItemIndex(IRibbonControl control);

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(24), Description("Call back for OnAction events from the drop-down ribbon elements.")]
        void OnActionSelected(IRibbonControl control, string selectedId, int selectedIndex);
        #endregion

        #region IEditableVM implementation - EditBox & ComboBox
        /// <summary>Gets the text to display in the edit box.</summary>
        /// <param name="control"></param>
        [DispId(25), Description("Gets the text to display in the edit box.")]
        string GetText(IRibbonControl control);

        /// <summary>Called when the value in the edit box is changed and committed by the user.</summary>
        /// <param name="control"></param>
        [DispId(26), Description("Called when the value in the edit box is changed and committed by the user.")]
        void OnTextChanged(IRibbonControl control, string text);
        #endregion

        #region IDynamicMenuVM implementation
        /// <summary>Gets an XML string that contains the contents of this dynamic menu.</summary>
        /// <param name="control"></param>
        [DispId(27), Description("Gets an XML string that contains the contents of this dynamic menu.")]
        string GetContent(IRibbonControl control);
        #endregion

        #region GallerySizeVM implementation
        /// <summary>Asks for the height of items, in pixels.</summary>
        /// <param name="control"></param>
        [DispId(28), Description("Asks for the height of items, in pixels.")]
        int GetItemHeight(IRibbonControl control);

        /// <summary>Asks for the width of items, in pixels.</summary>
        /// <param name="control"></param>
        [DispId(29), Description("Asks for the width of items, in pixels.")]
        int GetItemWidth(IRibbonControl control);
        #endregion

        #region IMenuSeparatorVM implementation
        /// <summary>For a menu separator, gets the text to be displayed .</summary>
        /// <param name="control"></param>
        [DispId(30), Description("For a menu separator, gets the text to be displayed ")]
        string GetTitle(IRibbonControl control);
        #endregion
    }
}
