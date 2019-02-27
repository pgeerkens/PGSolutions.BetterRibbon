////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the common link between <see cref="ControlModel{TSource, TCtrl}"/> objects created from VBA and C#.></summary>
    internal static partial class ModelFactoryExtensions {

        /// <summary>.</summary>
        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel: ControlModel<TSource, TVM> where TSource: IControlSource where TVM: IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}
