////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [CLSCompliant(true)]
    public class BoxControlVM: AbstractContainerVM<IBoxControlSource,IBoxControlVM>, IBoxControlVM,
             IActivatable<IBoxControlSource,IBoxControlVM> {
        internal protected BoxControlVM(string controlId, KeyedControls controls) : base(controlId, controls) { }
        internal protected BoxControlVM(string controlId) : base(controlId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IBoxControlVM Attach(IBoxControlSource source) => Attach<BoxControlVM>(source);

        protected override bool DefaultShowInactive { get => true; set { } }

        public string    ControlId => base.ControlId;
    }
}
