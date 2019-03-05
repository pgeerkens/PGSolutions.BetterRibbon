////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public abstract class AbstractControlVM<TSource,TVM>: IControlVM, IActivatable<TSource,TVM>
        where TSource: IControlSource where TVM:class,IControlVM {
        /// <summary>TODO</summary>
        protected AbstractControlVM(string controlId) =>  ControlId = controlId;

        #region Common Control implementation
        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <summary>Raised after the control has been purged from the ViewModel and can no longer service Ribbon callbacks.</summary>
        internal event PurgedEventHandler Purged;

        public void OnPurged(IContainerControl sender) {
            Source = default;
            Purged?.Invoke(sender, new ControlPurgedEventArgs(ControlId));
        }

        /// <inheritdoc/>
        public         string ControlId { get; }
        /// <inheritdoc/>
        public virtual string KeyTip    => Source?.KeyTip ?? "";
        /// <inheritdoc/>
        public virtual string Label     => Source?.Label ?? ControlId;
        /// <inheritdoc/>
        public virtual string ScreenTip => Source?.ScreenTip ?? $"{ControlId} ScreenTip";
        /// <inheritdoc/>
        public virtual string SuperTip  => Source?.SuperTip ?? $"{ControlId} SuperTip";

        /// <inheritdoc/>
        public bool IsEnabled => Source?.IsEnabled ?? false;

        /// <inheritdoc/>
        public bool IsVisible => Source?.IsVisible ?? ShowInactive;
        #endregion

        #region IActivatable implementation
        protected TSource Source { get; private set; }

        protected bool IsAttached => Source != null;

        /// <inheritdoc/>
        protected virtual T Attach<T>(TSource source) where T: AbstractControlVM<TSource,TVM> {
            Source = source;
            Invalidate();
            return this as T;
        }

        public abstract TVM Attach(TSource source);

        /// <inheritdoc/>
        public virtual void Detach() => Source = default;

        /// <inheritdoc/>
        public bool ShowInactive => DefaultShowInactive;

        /// <inheritdoc/>
        public void SetShowInactive(bool showInactive) {
            DefaultShowInactive = showInactive; Invalidate();
        }
        protected virtual bool DefaultShowInactive { get; set; } = false;
        #endregion

        /// <inheritdoc/>
        public virtual void Invalidate() => Changed?.Invoke(this, new ControlChangedEventArgs(this));
    }

    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public abstract class AbstractControl2VM<TSource,TVM>: AbstractControlVM<TSource,TVM>, IControlVM,
            IActivatable<TSource,TVM>
        where TSource: IControlSource2 where TVM:class,IControlVM {
        /// <summary>TODO</summary>
        protected AbstractControl2VM(string controlId) : base(controlId) { }

        public virtual string Description => Source?.Description ?? $"{ControlId} Description";
    }
}
