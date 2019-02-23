////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;
namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    public interface IEditBox : IRibbonControlVM, IEditable {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //[Description("Returns the unique (within this ribbon) identifier for this control.")]
        //new string Id           { get; }
        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description  { get; }
        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip       { get; }
        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label        { get; }
        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip    { get; }
        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip     { get; }

        ///// <summary>Gets or sets whether the control is enabled.</summary>
        //[Description("Gets or sets whether the control is enabled.")]
        //new bool IsEnabled      { get; }
        ///// <summary>Gets or sets whether the control is visible.</summary>
        //[Description("Gets or sets whether the control is visible.")]
        //new bool IsVisible      { get; }

        ///// <inheritdoc/>
        //new void Invalidate();
    }

    public interface IComboBox: IDropDownVM, IEditBox {

    }
}
