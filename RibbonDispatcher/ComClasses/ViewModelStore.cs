////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Description("")]    
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IViewModelStore))]
    [Guid(Guids.ViewModelStore)]
    public sealed class ViewModelStore : IViewModelStore {
        internal ViewModelStore() {

        }

        private IReadOnlyDictionary<string, IActivatable> AdaptorControls { get; }

        IRibbonGroup IViewModelStore.AttachGroup(string controlId, IRibbonControlStrings strings)
            => AttachGroup(controlId,strings);
        IRibbonButton IViewModelStore.AttachButton(string controlId, IRibbonControlStrings strings)
            => AttachButton(controlId,strings);
        IRibbonToggle IViewModelStore.AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) => AttachToggle(controlId,strings,source);
        IRibbonToggle IViewModelStore.AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) => AttachCheckBox(controlId,strings,source);
        IRibbonDropDown IViewModelStore.AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source) => AttachDropDown(controlId,strings,source);

        internal RibbonGroup AttachGroup(string controlId, IRibbonControlStrings strings) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonGroup;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach();
            return ctrl;
        }

        internal RibbonButton AttachButton(string controlId, IRibbonControlStrings strings) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach();
            return ctrl;
        }

        internal RibbonToggleButton AttachToggle(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonToggleButton;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            ctrl?.Invalidate();
            return ctrl;
        }

        internal RibbonCheckBox AttachCheckBox(string controlId, IRibbonControlStrings strings,
                IBooleanSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonCheckBox;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }

        internal RibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings,
                IIntegerSource source) {
            var ctrl = AdaptorControls.FirstOrDefault(kv => kv.Key == controlId).Value as RibbonDropDown;
            ctrl?.SetLanguageStrings(strings ?? RibbonControlStrings.Default(controlId));
            ctrl?.Attach(source.Getter);
            return ctrl;
        }
    }
}
