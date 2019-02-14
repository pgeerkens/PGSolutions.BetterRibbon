////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public class RibbonGroupViewModel : RibbonCommon, IRibbonGroup, IActivatableControl<IRibbonCommon,bool> {
        public RibbonGroupViewModel(IRibbonFactory factory, string itemId, IRibbonControlStrings strings, bool visible, bool enabled)
        : base(itemId, strings, visible, enabled) {
            Factory = factory;
            AdaptorControls = new Dictionary<string, IActivatable>();
            Add(this);
        }

        public IRibbonFactory Factory { get; }

        protected static string NoImage => "MacroSecurity";

        public RibbonGroupViewModel Add(IActivatable control) {
            if (control == null) return null;
            AdaptorControls.Add(new KeyValuePair<string, IActivatable>(control.Id, control));
            return this;
        }

        public void DetachControls() {
            foreach (var ctrl in AdaptorControls) if(ctrl.Value != this) ctrl.Value?.Detach();
        }

        protected IDictionary<string, IActivatable> AdaptorControls { get; }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        public IRibbonGroup Attach(Func<bool> showInactiveGetter) {
            base.Attach();
            ShowInactiveGetter = showInactiveGetter;
            return this;
        }

        public override void Detach() {
            foreach (var c in AdaptorControls) c.Value.Detach();
            ShowInactiveGetter = () => false;
            base.Detach();
        }

        public override void Invalidate() {
            foreach (var ctrl in AdaptorControls) { if (ctrl.Value != this) ctrl.Value.Invalidate(); }

            base.Invalidate();
        }

        /// <inheritdoc/>>
        public bool ShowInactive => ShowInactiveGetter?.Invoke() ?? false;

        /// <inheritdoc/>>
        public virtual void SetShowInactive(bool showInactive) {
            foreach (var ctrl in AdaptorControls) {
                ctrl.Value.ShowActiveOnly = !showInactive;
                ctrl.Value.Invalidate();
            }
        }

        public TControl GetControl<TControl>(string controlId) where TControl : RibbonCommon
        => AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as TControl;

        private Func<bool> ShowInactiveGetter { get; set; }

        /// <inheritdoc/>>
        IRibbonCommon IActivatableControl<IRibbonCommon,bool>.Attach(Func<bool> getter) =>
            Attach(getter) as IRibbonCommon;

        /// <inheritdoc/>>
        void IActivatableControl<IRibbonCommon,bool>.Detach() => Detach();
        #endregion
    }
}
