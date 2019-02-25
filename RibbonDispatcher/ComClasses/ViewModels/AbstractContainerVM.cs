////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal abstract class AbstractContainerVM<TSource>: AbstractControlVM<TSource>, IContainerControl
        where TSource : IControlSource {
        protected AbstractContainerVM(IViewModelFactory factory, string itemId)
        : base(itemId) {
            Factory = factory;
            Controls = new KeyedControls();
        }

        internal IViewModelFactory Factory { get; }

        protected KeyedCollection<string, IActivatable> Controls { get; }

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        public void Add(IActivatable control) {
            if (control == null) return;
            Controls.Add(control);
        }

        public new void SetShowInactive(bool showInactive) {
            foreach (var vm in Controls) {
                if (vm != this) { vm.SetShowInactive(showInactive); }
            }
            base.SetShowInactive(showInactive);
        }

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

        public override void Detach() {
            Invalidate(ctrl => {
                ctrl.Detach();
                ctrl.SetShowInactive(false);
            });
            base.Detach();
        }

        IEnumerator<IActivatable> IEnumerable<IActivatable>.GetEnumerator() => Controls.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => Controls.GetEnumerator();
    }
}
