////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCommon))]
    [Guid(Guids.RibbonCommon)]
    public abstract class RibbonCommon : IRibbonCommon, IActivatableControl<IRibbonCommon> {
        /// <summary>TODO</summary>
        protected RibbonCommon(string itemId, IRibbonControlStrings strings, bool visible, bool enabled) {
            Id       = itemId;
            Strings  = strings;
            _visible = visible;
            _enabled = enabled;
        }

        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        public         string Id          { get; }
        /// <inheritdoc/>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        public virtual string Description => Strings?.Description ?? "";
        /// <inheritdoc/>
        [Description("Returns the KeyTip string for this control.")]
        public virtual string KeyTip      => Strings?.KeyTip ?? "";
        /// <inheritdoc/>
        [Description("Returns the Label string for this control.")]
        public virtual string Label       => Strings?.Label ?? Id;
        /// <inheritdoc/>
        [Description("Returns the screenTip string for this control.")]
        public virtual string ScreenTip   => Strings?.ScreenTip ?? Id;
        /// <inheritdoc/>
        [Description("Returns the SuperTip string for this control.")]
        public virtual string SuperTip    => Strings?.SuperTip ?? "";
        /// <inheritdoc/>
        [Description("Returns the SuperTip string for this control.")]
        public virtual string AlternateLabel => Strings?.AlternateLabel ?? "";

        /// <inheritdoc/>
        protected IRibbonControlStrings Strings { get; private set; }

        #region IActivatable implementation
        public bool ShowWhenInactive { get; set; } = false;

        public virtual IRibbonCommon Attach() {
            IsAttached = true;
            return this;
        }

        public virtual T Attach<T>() where T : RibbonCommon {
            IsAttached = true;
            return this as T;
        }

        public virtual void Detach() {
            IsAttached = false;
            SetLanguageStrings(RibbonControlStrings.Empty);
            Invalidate();
        }

        /// <inheritdoc/>
        public virtual bool IsEnabled {
            get => _enabled  && IsAttached;
            set { _enabled = value; Invalidate(); }
        }
        private bool _enabled;

        /// <inheritdoc/>
        public virtual bool IsVisible {
            get => (_visible && IsAttached)  || ShowWhenInactive;
            set { _visible = value; Invalidate(); }
        }

        protected bool IsAttached { get; private set; } = false;

        private bool _visible;
        #endregion

        /// <inheritdoc/>
        public void SetLanguageStrings(IRibbonControlStrings strings) {
            Strings = strings;
            Invalidate();
        }

        public void SetLanguageStrings() => SetLanguageStrings(RibbonControlStrings.Default(Id));

        /// <inheritdoc/>
        public virtual void Invalidate() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));
    }
}
