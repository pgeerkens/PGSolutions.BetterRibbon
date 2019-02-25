////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class LabelVM: AbstractControlVM<IControlSource>, ILabelVM,
             IActivatable<IControlSource, LabelVM> {
        public LabelVM(string itemId) : base(itemId) { }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        LabelVM IActivatable<IControlSource, LabelVM>.Attach(IControlSource source)
        => Attach<LabelVM>(source);

        public override string Description
        => throw new InvalidOperationException("Attribute Description not supported on a Label.");
    }
}
