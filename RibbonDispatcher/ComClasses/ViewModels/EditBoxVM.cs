////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    public class EditBoxVM : AbstractControlVM<IEditBoxSource>, IEditBox,
            IActivatable<IEditBoxSource,EditBoxVM> {
        internal EditBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new EditBoxVM Attach(IEditBoxSource source) => Attach<EditBoxVM>(source);

        public override void Detach() {
            EditCommitted = null;
            base.Detach();
        }
        #endregion

        #region IEditable implementation
        public event EventHandler EditCommitted;

        public int    MaxLength     => Source?.MaxLength ?? 10;
        public string Text          => Source?.Text ?? "";
        public string SizeString    => Source?.Text ?? "mnmnmnmnmn";

        public void OnEditCOmmitted(object sender, EventArgs e)
        => EditCommitted?.Invoke(this, EventArgs.Empty);
        #endregion

    }

    public interface IEditBoxSource : IRibbonCommonSource {
        event EventHandler EditCommitted;

        int    MaxLength    { get; }
        string Text         { get; }
        string SizeString   { get; }
    }
    public interface IEditBox : IRibbonControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string Id           { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description  { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip       { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        new string Label        { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip    { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip     { get; }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [Description("Gets or sets whether the control is enabled.")]
        new bool IsEnabled      { get; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [Description("Gets or sets whether the control is visible.")]
        new bool IsVisible      { get; }

        /// <inheritdoc/>
        new void Invalidate();
    }
}
