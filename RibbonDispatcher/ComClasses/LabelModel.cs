﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The COM visible Model for Ribbon Label controls.</summary>
    [Description("The COM visible Model for Ribbon Label controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILabelModel))]
    [Guid(Guids.LabelModel)]
    public class LabelModel: ControlModel<ILabelSource, ILabelVM>,
            ILabelModel, ILabelSource {
        internal LabelModel(Func<string, LabelVM> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
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