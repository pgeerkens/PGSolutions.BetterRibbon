////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using System.ComponentModel;
using PGSolutions.RibbonDispatcher.ControlMixins;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IClickable<T> : IRibbonControlModel<T> {
        void OnClicked();
    }
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IToggleable<T> : IRibbonControlModel<T> {
        void OnToggled(bool isPressed);
    }
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControlModel<T> {
        V SetViewModel<V>(T viewModel);
        T ViewModel { get; }
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonAdapter<T> {
        T SetProxy(T proxy);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringLoader {
        IRibbonTextLanguageControl GetStrings(string ControlId);
    //    IRibbonTextLanguageControl GetStrings(string ControlId, CultureInfo cultureInfo);
    }

    /// <summary>The ViewModel for RibbonButton adapters.</summary>
    [Description("The ViewModel for Ribbon Button adapters.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonAdapter<IClickable<IRibbonButton>>))]
    [Guid(Guids.RibbonButtonAdaptor)]
    public class RibbonButtonAdaptor : RibbonButton, IRibbonAdapter<IClickable<IRibbonButton>> {
        internal RibbonButtonAdaptor(string itemId, IResourceManager mgr, bool visible, bool enabled,
                RdControlSize size, ImageObject image, bool showImage, bool showLabel)
            : base(itemId, mgr, visible, enabled, size, image, showImage, showLabel) {
        }

        public override bool IsVisible => base.IsVisible && Proxy != null;

        private IClickable<IRibbonButton> Proxy { get; set; }

        public IClickable<IRibbonButton> SetProxy(IClickable<IRibbonButton> proxy) {

            return Proxy = proxy.SetViewModel<IClickable<IRibbonButton>>(this);
        }

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public override void OnClicked() => Proxy.OnClicked();
    }

    /// <summary>The ViewModel for RibbonButton adapters.</summary>
    [Description("The ViewModel for Ribbon Button adapters.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonAdapter<IToggleable<IRibbonToggleButton>>))]
    [Guid(Guids.RibbonToggleButtonAdaptor)]
    public class RibbonToggleButtonAdaptor : RibbonToggleButton, IToggleableMixin, IRibbonAdapter<IToggleable<IRibbonToggleButton>> {
        internal RibbonToggleButtonAdaptor(string itemId, IResourceManager mgr, bool visible, bool enabled,
                RdControlSize size, ImageObject image, bool showImage, bool showLabel)
            : base(itemId, mgr, visible, enabled, size, image, showImage, showLabel) {
        }

        public override bool IsVisible => base.IsVisible && Proxy != null;

        private IToggleable<IRibbonToggleButton> Proxy { get; set; }

        public IToggleable<IRibbonToggleButton> SetProxy(IToggleable<IRibbonToggleButton> proxy) {

            return Proxy = proxy.SetViewModel<IToggleable<IRibbonToggleButton>>(this);
        }

        /// <summary>TODO</summary>
        public override bool IsPressed {
            get => this.GetPressed();
            set => this.SetPressed(value);
        }

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public override void OnToggled(bool isPressed) => Proxy.OnToggled(isPressed);
    }
}
