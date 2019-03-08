////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    internal delegate void ChangedEventHandler(object sender, ControlChangedEventArgs e);

    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    internal delegate void PurgedEventHandler(object sender, ControlPurgedEventArgs e);

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    internal class ControlChangedEventArgs : EventArgs, IControlChangedEventArgs {
        /// <summary>TODO</summary>
        public ControlChangedEventArgs(IControlVM control) => Control = control;
        /// <summary>TODO</summary>
        public IControlVM Control { get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    internal class ControlPurgedEventArgs: EventArgs {
        /// <summary>TODO</summary>
        public ControlPurgedEventArgs(string controlId) => ControlId = controlId;
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Performance","CA1811:AvoidUncalledPrivateCode")]
        public string ControlId { get; }
    }
}
