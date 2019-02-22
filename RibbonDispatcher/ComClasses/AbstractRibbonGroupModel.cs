﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonGroupModel : IRibbonCommonSource, IInvalidate {
        protected AbstractRibbonGroupModel(GroupVM viewModel) {
            ViewModel = (viewModel as IActivatable<IRibbonCommonSource,GroupVM>)
                      ?.Attach(this);
            Strings   = ViewModel?.Factory.GetStrings(ViewModel.Id);
        }

        public bool     IsEnabled    { get; set; } = true;
        public bool     IsVisible    { get; set; } = true;
        public bool     ShowInactive { get; private set; } = true;
        public IStrings Strings      { get; private set; }

        protected GroupVM ViewModel { get; }

        public void Invalidate() => Invalidate(null);

        public virtual void Invalidate(Action<IActivatable> action) => ViewModel?.Invalidate(action);

        /// <summary>Set ShowInactive for al- child controls of this ViewModel - even the unattached.</summary>
        /// <param name="showInactive">The <see cref="bool"/> value to be set</param>
        public void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.Invalidate(c => c.SetShowInactive(ShowInactive));
        }

        public void DetachControls() => ViewModel?.Detach();
    }
}
