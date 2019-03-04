////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [CLSCompliant(true)]
    public class ButtonGroupVM: AbstractContainerVM<IButtonGroupSource,IButtonGroupVM>, IButtonGroupVM,
             IActivatable<IButtonGroupSource,IButtonGroupVM> {
    //    internal protected BoxControlVM(string controlId, KeyedControls controls) : base(controlId, controls) { }
        internal protected ButtonGroupVM(string controlId, IEnumerable<IControlVM> controls) : base(controlId,controls) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IButtonGroupVM Attach(IButtonGroupSource source) => Attach<ButtonGroupVM>(source);

        protected override bool DefaultShowInactive { get => true; set { } }
    }
}
