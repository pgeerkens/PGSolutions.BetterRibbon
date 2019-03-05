////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings2 = IControlStrings2;

    /// <summary>The COM visible Model for Ribbon Split (Press) Button controls.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Split (Press) Button controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvent))]
    [ComDefaultInterface(typeof(ISplitPressButtonModel))]
    [Guid(Guids.SplitPressButtonModel)]
    public class SplitPressButtonModel: AbstractSplitButtonModel<IButtonSource,ISplitPressButtonVM>,
            ISplitPressButtonModel, IButtonSource {
        internal SplitPressButtonModel(Func<string,SplitPressButtonVM> funcViewModel,
                IStrings2 strings, ButtonModel button, MenuModel menu)
        : base(funcViewModel, strings, menu)
        => _buttonModel = button;

        public ISplitPressButtonModel Attach(string controlId) {
            base.Attach(controlId, this);
            if (ViewModel != null) {
                _buttonModel.Attach(ViewModel.ButtonVM.ControlId);
                _buttonModel.ViewModel.Clicked += OnClicked;
            }
            return this;
        }

        public override void Detach() { ButtonModel.Detach(); base.Detach(); }

        #region Pressable implementation
        public event ClickedEventHandler Clicked;

        public IButtonModel ButtonModel => _buttonModel; private ButtonModel _buttonModel { get; }

        private void OnClicked(IRibbonControl control) => Clicked?.Invoke(control);
        #endregion

        public ISplitPressButtonModel SetImage(IImageObject image) {Image = image; return this; }
    }
}
