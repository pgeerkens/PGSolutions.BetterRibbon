////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    internal class LabelControlVM: AbstractControlVM<ILabelControlSource,ILabelControlVM>, ILabelControlVM,
             IActivatable<ILabelControlSource,ILabelControlVM>, ISizeableVM {
        public LabelControlVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override ILabelControlVM Attach(ILabelControlSource source) => Attach<LabelControlVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        protected override bool DefaultShowInactive { get => true; set { } }
    }
}
