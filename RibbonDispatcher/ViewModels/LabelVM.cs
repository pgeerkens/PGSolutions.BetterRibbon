////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    internal class LabelVM: AbstractControlVM<ILabelSource,ILabelVM>, ILabelVM,
             IActivatable<ILabelSource,ILabelVM>, ISizeableVM {
        public LabelVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override ILabelVM Attach(ILabelSource source) => Attach<LabelVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        protected override bool DefaultShowInactive { get => true; set { } }
    }
}
