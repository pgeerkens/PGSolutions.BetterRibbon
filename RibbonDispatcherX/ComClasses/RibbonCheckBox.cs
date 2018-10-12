////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for Ribbon CheckBox objects.</summary>
    [Description("The ViewModel for Ribbon CheckBox objects.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonCheckBox))]
    [Guid(Guids.RibbonCheckBox)]
    public class RibbonCheckBox : RibbonCommon, IRibbonCheckBox, IActivatableControl<IRibbonCommon, bool>,
        IToggleableMixin {
        internal RibbonCheckBox(string itemId, IResourceManager mgr, bool visible, bool enabled)
            : base(itemId, mgr, visible, enabled) {
        }

        #region IToggleable implementation
        private bool _isAttached    = false;
        private bool _enableVisible = true;

        public override bool IsEnabled => base.IsEnabled && _isAttached;
        public override bool IsVisible => base.IsEnabled && _enableVisible;

        public IRibbonCheckBox Attach(Func<bool> getter) {
            _isAttached = true;
            _enableVisible = true;
            this.SetGetter(getter);
            return this;
        }

        public void Detach() => Detach(true);
        public void Detach(bool enableVisible) {
            _enableVisible = enableVisible;
            _isAttached = false;
            SetLanguageStrings(RibbonTextLanguageControl.Empty);
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, bool>.Attach(Func<bool> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, bool>.Detach() => Detach();
        #endregion

        #region Publish IToggleableMixin to class default interface
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public bool IsPressed           => this.GetPressed();

        /// <summary>TODO</summary>
        public override string Label    => this.GetLabel();

        /// <summary>TODO</summary>
        public void OnToggled(bool IsPressed) => Toggled?.Invoke(IsPressed);

        /// <summary>TODO</summary>
        IRibbonTextLanguageControl IToggleableMixin.LanguageStrings => LanguageStrings;
        #endregion
    }
}
