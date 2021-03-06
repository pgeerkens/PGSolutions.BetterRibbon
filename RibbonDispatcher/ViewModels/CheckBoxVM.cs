﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon CheckBoxVM objects.</summary>
    internal class CheckBoxVM : AbstractControl2VM<IToggleSource,IToggleVM>, IToggleVM,
        IActivatable<IToggleSource, IToggleVM>, IToggleableVM {
        public CheckBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public override IToggleVM Attach(IToggleSource source) => Attach<CheckBoxVM>(source);

        /// <inheritdoc/>
        public override void Detach() { Toggled = null; base.Detach(); }
        #endregion

        #region IToggleable implementation
        /// <inheritdoc/>
        public event ToggledEventHandler Toggled;

        /// <inheritdoc/>>
        public bool IsPressed => Source?.IsPressed ?? false;

        /// <inheritdoc/>>
        public virtual void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control,isPressed);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public virtual bool IsLarge => false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public virtual IImageObject Image => null;

        /// <inheritdoc/>
        public virtual bool ShowImage => false;

        /// <inheritdoc/>
        public virtual bool ShowLabel => true;
        #endregion
    }
}
