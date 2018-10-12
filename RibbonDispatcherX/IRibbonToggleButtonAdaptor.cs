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
    /// <summary>The ViewModel for RibbonButton adapters.</summary>
    [Description("The ViewModel for Ribbon Button adapters.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggleButtonAdaptor))]
    [Guid(Guids.RibbonToggleButtonAdaptor)]
    public class RibbonToggleButtonAdaptor : RibbonToggleButton, IToggleableMixin, IRibbonToggleButtonAdaptor {
        internal RibbonToggleButtonAdaptor(string itemId, IResourceManager mgr, bool visible, bool enabled,
                RdControlSize size, ImageObject image, bool showImage, bool showLabel)
            : base(itemId, mgr, visible, enabled, size, image, showImage, showLabel) {
        }

        public override bool IsVisible => base.IsVisible && Proxy != null;

        private IToggleableRibbonToggleButton Proxy { get; set; }

        public IToggleableRibbonToggleButton SetProxy(IToggleableRibbonToggleButton proxy) {
            proxy.SetViewModel(this);
            return Proxy = proxy;
        }

        /// <summary>TODO</summary>
        public override bool IsPressed => this.GetPressed();

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public override void OnToggled(bool isPressed) => Proxy.OnToggled(isPressed);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IToggleableRibbonToggleButton {
        IRibbonToggleButton ViewModel { get; }
        void OnToggled(bool isPressed);
        void SetViewModel(IRibbonToggleButton viewModel);
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonToggleButtonAdaptor {
        IToggleableRibbonToggleButton SetProxy(IToggleableRibbonToggleButton proxy);
    }
}
