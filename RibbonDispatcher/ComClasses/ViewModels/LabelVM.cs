////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class LabelVM: AbstractControlVM<ILabelSource>, ILabelVM,
             IActivatable<ILabelSource,ILabelVM>, ISizeableVM {
        public LabelVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public new ILabelVM Attach(ILabelSource source) => Attach<LabelVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        public override string Description
        => throw new InvalidOperationException("Attribute Description not supported on a Label.");
    }
}
