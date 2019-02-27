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
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Label controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILabelModel))]
    [Guid(Guids.LabelModel)]
    public class LabelModel: ControlModel<ILabelSource, ILabelVM>,
            ILabelModel, ILabelSource {
        internal LabelModel(Func<string, LabelVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings)
        { }

        public bool        IsLarge   { get; set; } = true;
        public bool        ShowLabel { get; set; } = true;

        public ILabelModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            ViewModel?.Invalidate();
            return this;
        }
    }
}
