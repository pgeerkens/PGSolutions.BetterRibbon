using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.AbstractCOM;
using LanguageStrings = PGSolutions.RibbonDispatcher.AbstractCOM.IRibbonTextLanguageControl;

namespace PGSolutions.RibbonDispatcher.Concrete {

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

    //    [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)")]
        private static LanguageStrings GetLanguageStrings(string controlId, IResourceManager mgr)
            => new RibbonTextLanguageControl(
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_Label"))          ?? controlId,
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_ScreenTip"))      ?? controlId + " ScreenTip",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_SuperTip"))       ?? controlId + " SuperTip",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_KeyTip"))         ?? "",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_AlternateLabel")) ?? controlId + " Alternate",
                    mgr.GetCurrentUIString(Invariant($"{controlId ?? ""}_Description"))    ?? controlId + " Description");
        /// <summary>TODO</summary>
        private static string Invariant(string formattable) => string.Format(formattable, CultureInfo.InvariantCulture);
    }
}
