////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>These extension methods on <see cref="ViewModelFactory"/> are the common link between <see cref="ControlModel{TSource, TCtrl}"/> objects created from VBA and C#.></summary>
    public static partial class ModelFactoryExtensions {
        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactory NewModelFactory(this AbstractDispatcher dispatcher)
            => dispatcher?.NewModelFactory(dispatcher?.ResourceLoader)
                ?? throw new ArgumentNullException(nameof(dispatcher));

        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelFactory NewModelFactory(this AbstractDispatcher dispatcher, IResourceLoader resourceLoader)
            => new ModelFactory(dispatcher?.ViewModelFactory, resourceLoader)
                ?? throw new ArgumentNullException(nameof(dispatcher));

        /// <summary>Returns a new instance of an <see cref="IModelFactory"/>.</summary>
        /// <param name="model"></param>
        public static IModelServer NewModelServer(this AbstractDispatcher dispatcher, IResourceLoader resourceLoader)
            => new ModelFactory(dispatcher?.ViewModelFactory, resourceLoader)
                ?? throw new ArgumentNullException(nameof(dispatcher));

        /// <summary>.</summary>
        internal static TModel InitializeModel<TSource, TVM, TModel>(this TModel model)
            where TModel: class,IControlSource where TSource: IControlSource where TVM: IControlVM {

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}
