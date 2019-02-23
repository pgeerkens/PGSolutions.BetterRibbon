////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for Ribbon CheckBoxVM objects.</summary>
    internal class CheckBoxVM : AbstractControlVM<IToggleSource>, IToggleControlVM,
        IActivatable<IToggleSource, IToggleControlVM>, IToggleableVM {
        internal CheckBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new IToggleControlVM Attach(IToggleSource source) => Attach<CheckBoxVM>(source);

        public override void Detach() {
            Toggled = null;
            base.Detach();
        }
        #endregion

        #region IToggleable implementation
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <inheritdoc/>>
        public bool IsPressed => Source?.IsPressed ?? false;

        /// <inheritdoc/>>
        public virtual void OnToggled(IRibbonControl control, bool isPressed)
        => Toggled?.Invoke(control,isPressed);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public virtual bool IsSizeable => false;

        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        public virtual bool IsLarge => false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public virtual bool IsImageable => false;
        /// <inheritdoc/>
        public virtual ImageObject Image => null;

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowImage => false;

        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowLabel => true;
        #endregion
    }
}
