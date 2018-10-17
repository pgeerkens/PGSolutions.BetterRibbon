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
        internal RibbonCheckBox(string itemId, IRibbonControlStrings strings, bool visible, bool enabled
        ) : base(itemId, strings, visible, enabled) { }

        #region IActivatable implementation
        private bool _isAttached    = false;

        public override bool IsEnabled => base.IsEnabled && _isAttached;
        public override bool IsVisible => (base.IsVisible && _isAttached)
                                       || (ShowWhenInactive);

        public bool ShowWhenInactive { get; set; } //= true;

        public IRibbonCheckBox Attach(Func<bool> getter) {
            _isAttached = true;
            this.SetGetter(getter);
            return this;
        }

        public void Detach() {
            _isAttached = false;
            this.SetGetter(()=>false);
            SetLanguageStrings(RibbonControlStrings.Empty);
            Invalidate();
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
        IRibbonControlStrings IToggleableMixin.LanguageStrings => Strings;
        #endregion
    }
}
