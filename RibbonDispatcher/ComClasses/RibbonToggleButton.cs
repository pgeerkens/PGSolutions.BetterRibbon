////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    [Description("The ViewModel for Ribbon ToggleButton objects")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggle))]
    [Guid(Guids.RibbonToggleButton)]
    public class RibbonToggleButton : RibbonCheckBox, IRibbonToggle,
        IActivatable<RibbonToggleButton, IRibbonToggleSource>, ISizeable, IImageable {
        internal RibbonToggleButton(string itemId) : base(itemId) { }

        #region IActivatable implementation
        RibbonToggleButton IActivatable<RibbonToggleButton, IRibbonToggleSource>.Attach(IRibbonToggleSource source)
        => Attach<RibbonToggleButton>(source);
        #endregion

        #region IToggleable implementation
        /// <inheritdoc/>>
        public override string Label => IsPressed || string.IsNullOrEmpty(AlternateLabel)
                                     ? base.Label ?? Id
                                     : AlternateLabel;
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public override bool IsSizeable => true;
        /// <inheritdoc/>
        public override bool IsLarge => Source?.IsLarge ?? true;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override bool IsImageable => true;

        /// <inheritdoc/>
        public override object Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public override bool ShowImage => Source?.ShowImage ?? true;

        /// <inheritdoc/>
        public override bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
