////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [CLSCompliant(true)]
    public class GroupVM : AbstractContainerVM<IControlSource,IGroupVM>, IGroupVM, 
            IActivatable<IControlSource,IGroupVM> {
        //internal protected GroupVM(string controlId, KeyedControls controls) : base(controlId, controls) { }
        internal protected GroupVM(string controlId, IEnumerable<IControlVM> controls) : base(controlId, controls) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IGroupVM Attach(IControlSource source) => Attach<GroupVM>(source);

        protected override bool DefaultShowInactive { get => true; set { } }
    }
}
