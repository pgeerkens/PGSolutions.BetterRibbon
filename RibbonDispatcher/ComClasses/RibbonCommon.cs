////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCommon))]
    [Guid(Guids.RibbonCommon)]
    public abstract class RibbonCommon<TSource> : IRibbonCommon, IActivatable<IRibbonCommon,TSource>
        where TSource : IRibbonCommonSource {
        /// <summary>TODO</summary>
        protected RibbonCommon(string itemId) => Id = itemId;

        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        public string Id        { get; }
        /// <inheritdoc/>
        [Description("Returns the KeyTip string for this control.")]
        public string KeyTip => Strings?.KeyTip ?? "";
        /// <inheritdoc/>
        [Description("Returns the Label string for this control.")]
        public virtual string Label => Strings?.Label ?? Id;
        /// <inheritdoc/>
        [Description("Returns the screenTip string for this control.")]
        public string ScreenTip => Strings?.ScreenTip ?? $"{Id} ScreenTip";
        /// <inheritdoc/>
        [Description("Returns the SuperTip string for this control.")]
        public string SuperTip => Strings?.SuperTip ?? $"{Id} SuperTip";
        /// <inheritdoc/>
        [Description("Returns the SuperTip string for this control.")]
        public string AlternateLabel => Strings?.AlternateLabel ?? $"{Id} Alternate";
        /// <inheritdoc/>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        public string Description => Strings?.Description ?? $"{Id} Description";

        /// <inheritdoc/>
        protected virtual IRibbonControlStrings Strings => Source?.Strings;

        /// <inheritdoc/>
        public bool IsEnabled => Source?.IsEnabled ?? false;

        /// <inheritdoc/>
        public bool IsVisible => Source?.IsVisible ?? ShowInactive;

        /// <inheritdoc/>
        public bool ShowInactive => Source?.ShowInactive ?? _defaultShowInactive;

        #region IActivatable implementation
        protected TSource Source { get; private set; }

        protected bool IsAttached => Source != null;

        public virtual T Attach<T>(TSource source) where T:RibbonCommon<TSource> {
            Source = source;
            Invalidate();
            return this as T;
        }

        IRibbonCommon IActivatable<IRibbonCommon,TSource>.Attach(TSource source)
        => Attach<RibbonCommon<TSource>>(source);

        public virtual void Detach() {
            Source = default;
            Invalidate();
        }
        #endregion

        public void SetShowInactive(bool showInactive) {
            _defaultShowInactive = showInactive; Invalidate();
        } bool _defaultShowInactive = false;

        /// <inheritdoc/>
        public virtual void Invalidate() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));
    }
}
