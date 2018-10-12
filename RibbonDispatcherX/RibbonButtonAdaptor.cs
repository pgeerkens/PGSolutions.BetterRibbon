////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using System.ComponentModel;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComClasses
{
    /// <summary>The ViewModel for RibbonButton adapters.</summary>
    [Description("The ViewModel for Ribbon Button adapters.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickableRibbonButton))]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid(Guids.RibbonButtonAdaptor)]
    public class RibbonButtonAdaptor : RibbonButton, IRibbonButton, ActivatableControl<IRibbonCommon> {
        internal RibbonButtonAdaptor(string itemId, IResourceManager mgr, bool visible, bool enabled,
                RdControlSize size, ImageObject image, bool showImage, bool showLabel)
            : base(itemId, mgr, visible, enabled, size, image, showImage, showLabel) {
        }

        private bool _isAttached = false;

        public override bool IsEnabled => base.IsEnabled && _isAttached;

        public IRibbonButton Attach(IRibbonTextLanguageControl strings) {
            SetLanguageStrings(strings);
            _isAttached = true;
            return this;
        }

        public void Detach() {
            _isAttached = false;
            SetLanguageStrings(RibbonTextLanguageControl.Empty);
            SetImageMso("MacroSecurity");
        }

        IRibbonCommon ActivatableControl<IRibbonCommon>.Attach(IRibbonTextLanguageControl strings) =>
            Attach(strings) as IRibbonCommon;
        void ActivatableControl<IRibbonCommon>.Detach() => Detach();
    }

    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IClickableRibbonButton {
        void OnClicked();
    }

    public interface ActivatableControl<TCtl> where TCtl:IRibbonCommon {
        TCtl Attach(IRibbonTextLanguageControl strings);
        void Detach();
    }
}
