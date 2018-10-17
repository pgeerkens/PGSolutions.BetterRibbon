////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
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
    public abstract class RibbonCommon : IRibbonCommon {
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
        protected IRibbonControlStrings Strings { get; private set; }

        /// <inheritdoc/>
        public virtual bool IsEnabled {
            get => _enabled;
            set { _enabled = value; OnChanged(); }
        }
        private bool _enabled;

        /// <inheritdoc/>
        public virtual bool IsVisible {
            get => _visible;
            set { _visible = value; OnChanged(); }
        }
        private bool _visible;

        /// <inheritdoc/>
        public void SetLanguageStrings(IRibbonControlStrings strings) {
            Strings = strings;
            OnChanged();
        }

        public void SetLanguageStrings() => SetLanguageStrings(RibbonControlStrings.Default(Id));

        /// <inheritdoc/>
        public virtual void OnChanged() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));

        /// <inheritdoc/>
        public void Invalidate() => OnChanged();

        //private static LanguageStrings GetLanguageStrings(string controlId, IResourceManager mgr)
        //    => mgr.GetControlStrings(controlId);
    }
}
