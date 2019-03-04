////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public abstract class AbstractContainerVM<TSource,TVM>: AbstractControlVM<TSource,TVM>, IContainerControl
        where TSource : IControlSource where TVM:class,IControlVM {
        protected AbstractContainerVM(string itemId) : this(itemId, new KeyedControls()) { }
        protected AbstractContainerVM(string itemId, IEnumerable<IControlVM> controls) : base(itemId)
        => Controls = new KeyedControls(controls);

        protected KeyedControls Controls { get; set; }

        public TControl GetControl<TControl>(string controlId) where TControl : class, IControlVM
        => Controls.Item<TControl>(controlId);

        public void PurgeChildren() {
            foreach(var child in Controls) {
                if (child is IContainerControl container) container.PurgeChildren();
                child.Detach();
            }
            Controls.Clear();
        }

        [SuppressMessage("Microsoft.Design", "CA1004:GenericMethodsShouldProvideTypeParameter")]
        public void Add(IControlVM control) {
            if (control == null) return;
            Controls.Add(control);
        }

        public new      void SetShowInactive(bool showInactive) {
            foreach (var vm in Controls) {
                if (vm != this) { vm.SetShowInactive(showInactive); }
            }
            base.SetShowInactive(showInactive);
        }

        public override void Invalidate() => Invalidate(null);

        public          void Invalidate(Action<IControlVM> action) {
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

        public IEnumerator<IControlVM> GetEnumerator() => Controls.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => Controls.GetEnumerator();
    }
}
