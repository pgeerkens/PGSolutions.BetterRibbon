////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

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
        public ControlChangedEventArgs(string controlId) => ControlId = controlId;
        /// <summary>TODO</summary>
        public string ControlId { get; }
    }

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    internal class ControlPurgedEventArgs: EventArgs {
        /// <summary>TODO</summary>
        public ControlPurgedEventArgs(string controlId) => ControlId = controlId;
        /// <summary>TODO</summary>
        public string ControlId { get; }
    }
}
