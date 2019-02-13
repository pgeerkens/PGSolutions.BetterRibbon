////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class AbstractRibbonGroupViewModel : RibbonCommon, IRibbonGroup, IActivatableControl<IRibbonCommon, bool>, IToggleable {
        protected AbstractRibbonGroupViewModel(IRibbonFactory factory, string itemId, bool isVisible, bool isEnabled)
        : base (itemId, null, isVisible, isEnabled)
        => Factory = factory;

        internal AbstractRibbonGroupViewModel(string itemId, IRibbonControlStrings strings, bool visible, bool enabled)
        : base(itemId, strings, visible, enabled) { }

        protected IRibbonFactory Factory { get; }

        protected static string NoImage => "MacroSecurity";

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        public IRibbonGroup Attach(Func<bool> getter) {
            base.Attach();
            Getter = getter;
            return this;
        }

        public override void Detach() {
            Toggled = null;
            Getter = () => false;
            base.Detach();
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, bool>.Attach(Func<bool> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, bool>.Detach() => Detach();
        #endregion

        #region IToggleable implementation
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <inheritdoc/>>
        public bool IsPressed => Getter?.Invoke() ?? false;

        /// <inheritdoc/>>
        public virtual void OnToggled(object sender, bool isPressed) => Toggled?.Invoke(this, isPressed);

        /// <summary>TODO</summary>
        private Func<bool> Getter { get; set; }
        #endregion
    }
}
