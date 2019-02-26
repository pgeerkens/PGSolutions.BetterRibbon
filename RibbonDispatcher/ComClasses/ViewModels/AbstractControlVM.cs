////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    internal abstract class AbstractControlVM<TSource>: IControlVM, IActivatable<TSource,IControlVM>
        where TSource: IControlSource {
        /// <summary>TODO</summary>
        protected AbstractControlVM(string itemId) => Id = itemId;

        #region Common Control implementation
        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        public string Id             { get; }
        /// <inheritdoc/>
        public virtual string KeyTip         => Strings?.KeyTip ?? "";
        /// <inheritdoc/>
        public virtual string Label          => Strings?.Label ?? Id;
        /// <inheritdoc/>
        public virtual string ScreenTip      => Strings?.ScreenTip ?? $"{Id} ScreenTip";
        /// <inheritdoc/>
        public virtual string SuperTip       => Strings?.SuperTip ?? $"{Id} SuperTip";

        /// <inheritdoc/>
        protected virtual IControlStrings Strings => Source?.Strings;

        /// <inheritdoc/>
        public bool IsEnabled        => Source?.IsEnabled ?? false;

        /// <inheritdoc/>
        public bool IsVisible        => Source?.IsVisible ?? ShowInactive;
        #endregion

        #region IActivatable implementation
        protected TSource Source { get; private set; }

        protected bool IsAttached => Source != null;

        /// <inheritdoc/>
        protected virtual T Attach<T>(TSource source) where T: AbstractControlVM<TSource> {
            Source = source;
            Invalidate();
            return this as T;
        }

        public IControlVM Attach(TSource source) => Attach<AbstractControlVM<TSource>>(source);

        /// <inheritdoc/>
        public virtual void Detach() {
            Source = default;
            Invalidate();
        }

        /// <inheritdoc/>
        public bool ShowInactive => DefaultShowInactive;

        /// <inheritdoc/>
        public void SetShowInactive(bool showInactive) {
            DefaultShowInactive = showInactive; Invalidate();
        }
        protected virtual bool DefaultShowInactive { get; set; } = false;
        #endregion

        /// <inheritdoc/>
        public virtual void Invalidate() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));
    }
}
