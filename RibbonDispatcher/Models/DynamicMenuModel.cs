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
    [ComDefaultInterface(typeof(IDynamicMenuModel))]
    [Guid(Guids.DynamicMenuModel)]
    public class DynamicMenuModel: ControlModel<IDynamicMenuSource,IDynamicMenuVM>, IDynamicMenuModel,
            IDynamicMenuSource {
        internal DynamicMenuModel(Func<string,DynamicMenuVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings)
        { }

        /// <inheritdoc/>
        public IDynamicMenuModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            return this;
        }

        public new IControlStrings2 Strings => base.Strings as IControlStrings2;

        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        #region IImageable implementation
        public IDynamicMenuModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion

        public string       Content   { get; }
    }
}
