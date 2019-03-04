////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon Menu controls.</summary>
    [Description("The COM visible Model for Ribbon Menu controls.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMenuModel))]
    [Guid(Guids.MenuModel)]
    public class MenuModel: ControlModel<IMenuSource,IMenuVM>, IMenuModel,
            IMenuSource {
        internal MenuModel(Func<string,MenuVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings)
        { }

        /// <inheritdoc/>
        public IMenuModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            return this;
        }

        public new IControlStrings2 Strings => base.Strings as IControlStrings2;

        public bool         IsLarge   { get; set; } = true;

        #region IImageable implementation
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IMenuModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion
    }
}
