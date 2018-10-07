using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher2013.AbstractCOM;
using LanguageStrings = PGSolutions.RibbonDispatcher2013.AbstractCOM.IRibbonTextLanguageControl;

namespace PGSolutions.RibbonDispatcher2013.ConcreteCOM {

    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCommon))]
    [Guid(Guids.RibbonCommon)]
    public abstract class RibbonCommon : IRibbonCommon {
        /// <summary>TODO</summary>
        protected RibbonCommon(string itemId, IResourceManager resourceManager, bool visible, bool enabled) {
            Id               = itemId;
            LanguageStrings  = GetLanguageStrings(itemId, resourceManager);
            _visible         = visible;
            _enabled         = enabled;
        }

        /// <summary>TODO</summary>
        internal event ChangedEventHandler Changed;

        /// <inheritdoc/>
        public         string Id          { get; }
        /// <inheritdoc/>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        public virtual string Description => LanguageStrings?.Description ?? "";
        /// <inheritdoc/>
        [Description("Returns the KeyTip string for this control.")]
        public virtual string KeyTip      => LanguageStrings?.KeyTip ?? "";
        /// <inheritdoc/>
        [Description("Returns the Label string for this control.")]
        public virtual string Label       => LanguageStrings?.Label ?? Id;
        /// <inheritdoc/>
        [Description("Returns the screenTip string for this control.")]
        public virtual string ScreenTip   => LanguageStrings?.ScreenTip ?? Id;
        /// <inheritdoc/>
        [Description("Returns the SuperTip string for this control.")]
        public virtual string SuperTip    => LanguageStrings?.SuperTip ?? "";

        /// <inheritdoc/>
        protected LanguageStrings LanguageStrings { get; private set; }

        /// <inheritdoc/>
        public bool IsEnabled {
            get => _enabled;
            set { _enabled = value; OnChanged(); }
        }
        private bool _enabled;

        /// <inheritdoc/>
        public bool IsVisible {
            get => _visible;
            set { _visible = value; OnChanged(); }
        }
        private bool _visible;

        /// <inheritdoc/>
        public void SetLanguageStrings(LanguageStrings languageStrings) {
            LanguageStrings = languageStrings;
            OnChanged();
        }

        /// <inheritdoc/>
        public void OnChanged() => Changed?.Invoke(this, new ControlChangedEventArgs(Id));

        private static LanguageStrings GetLanguageStrings(string controlId, IResourceManager mgr)
            => mgr.GetControlStrings(controlId);
    }
}
