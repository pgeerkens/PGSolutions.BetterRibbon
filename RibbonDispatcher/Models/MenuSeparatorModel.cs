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
    /// <summary>The COM visible Model for Ribbon Label controls.</summary>
    [Description("The COM visible Model for Ribbon Menu Separator controls.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMenuSeparatorModel))]
    [Guid(Guids.MenuSeparatorModel)]
    public class MenuSeparatorModel: ControlModel<IMenuSeparatorSource, IMenuSeparatorVM>,
            IMenuSeparatorModel, IMenuSeparatorSource {
        internal MenuSeparatorModel(Func<string, MenuSeparatorVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings) { }

        public string Title { get; set; }

        public IMenuSeparatorModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            return this;
        }
    }
}
