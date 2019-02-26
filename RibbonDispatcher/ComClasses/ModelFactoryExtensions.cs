////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings  = IControlStrings;
    using IStrings2 = IControlStrings2;

    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the common link between <see cref="ControlModel{TSource, TCtrl}"/> objects created from VBA and C#.></summary>
    internal static partial class ModelFactoryExtensions {
        public static IStrings GetStrings(this IViewModelFactory vm, string controlId)
        => vm.ResourceManager.GetControlStrings(controlId);

        public static IStrings2 GetStrings2(this IViewModelFactory vm, string controlId)
        => vm.ResourceManager.GetControlStrings2(controlId);

        public static object LoadImage(this IViewModelFactory vm, string imageId)
        => vm.ResourceManager.GetImage(imageId);

        public static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel: ControlModel<TSource, TVM> where TSource: IControlSource where TVM: IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}
