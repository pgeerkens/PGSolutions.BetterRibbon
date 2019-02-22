////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    public class GroupVM : AbstractControlVM<IRibbonCommonSource>, IRibbonGroup,
            IActivatable<IRibbonCommonSource,GroupVM> {
        public GroupVM(IRibbonFactory factory, string itemId)
        : base(itemId) {
            Factory = factory;
            Controls = new KeyedControls();
            Add<IRibbonCommonSource>(this);
        }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        GroupVM IActivatable<IRibbonCommonSource,GroupVM>.Attach(IRibbonCommonSource source)
        => Attach<GroupVM>(source);

        public override void Detach() {
            Invalidate(ctrl => {
                ctrl.Detach();
                ctrl.SetShowInactive(false);
            });
            base.Detach();
        }

        internal IRibbonFactory Factory { get; }

        protected KeyedCollection<string,IActivatable> Controls { get; }

        public override void Invalidate() => Invalidate(null);

        internal void Invalidate(Action<IActivatable> action) {
            foreach (var ctrl in Controls) {
                if (ctrl != this) {
                    action?.Invoke(ctrl);
                    ctrl.Invalidate(); 
                }
            }
            base.Invalidate();
        }

        public TControl GetControl<TControl>(string controlId) where TControl : class,IRibbonControlVM
        => Controls.FirstOrDefault(ctl => ctl.Id == controlId) as TControl;

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        public GroupVM Add<TSource>(IActivatable control)
        where TSource:IRibbonCommonSource {
            if (control == null) return null;
            Controls.Add(control);
            return this;
        }

        protected static string NoImage => "MacroSecurity";

        public new void SetShowInactive(bool showInactive) {
            foreach (var vm in Controls) { 
                if (vm != this) { vm.SetShowInactive(showInactive); }
            }
            base.SetShowInactive(showInactive);
        }

        private class KeyedControls : KeyedCollection<string,IActivatable> {
            public KeyedControls() :base() { }
            protected override string GetKeyForItem(IActivatable control) => control?.Id;
        }
    }
}
