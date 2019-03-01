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
    [Description("The COM visible Model for Ribbon Label controls.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILabelControlModel))]
    [Guid(Guids.LabelControlModel)]
    public class LabelControlModel: ControlModel<ILabelControlSource, ILabelControlVM>,
            ILabelControlModel, ILabelControlSource {
        internal LabelControlModel(Func<string, LabelControlVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings) { }

        public bool        IsLarge   { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public ILabelControlModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            return this;
        }
    }
}
