////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {

    /// <summary></summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Description("")]    
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    //[ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonGroupModel))]
    [Guid(Guids.RibbonGroupModel)]
    public sealed class RibbonGroupModel : IRibbonGroupModel {
        internal RibbonGroupModel(ViewModelStore viewModelStore)
            => ViewModelStore = viewModelStore;

        public  IRibbonGroup   ViewModel      => _viewModel;
        private RibbonGroup    _viewModel     { get; set; }
        private ViewModelStore ViewModelStore { get; }

        public void Attach(string controlId, IRibbonControlStrings strings) =>
            _viewModel = ViewModelStore.AttachGroup(controlId, strings);
    }

    /// <summary></summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Description("")]    
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButtonModel))]
    [Guid(Guids.RibbonButtonModel)]
    public sealed class RibbonButtonModel : IRibbonButtonModel {
        internal RibbonButtonModel(Func<string,IRibbonControlStrings,RibbonButton> factory)
            => Factory = factory;

        public event ClickedEventHandler Clicked;

        public IRibbonButton   ViewModel      => _viewModel;
        private RibbonButton   _viewModel     { get; set; }
        private ViewModelStore ViewModelStore { get; }
        private Func<string,IRibbonControlStrings,RibbonButton> Factory { get; }

        public void OnClicked(object sender) => Clicked(sender);

        public void Attach(string controlId, IRibbonControlStrings strings) {
            _viewModel = Factory(controlId, strings);
            _viewModel.Clicked += OnClicked;
        }
    }

    /// <summary></summary>
    [SuppressMessage( "Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable" )]
    [Description("")]    
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggleModel))]
    [Guid(Guids.RibbonToggleModel)]
    public sealed class RibbonToggleModel : IRibbonToggleModel, IBooleanSource {
        internal RibbonToggleModel(ViewModelStore viewModelStore)
            => ViewModelStore = viewModelStore;

        public event ToggledEventHandler Toggled;

        public  IRibbonToggle       ViewModel      => _viewModel;
        private RibbonToggleButton  _viewModel     { get; set; }
        private ViewModelStore      ViewModelStore { get; }

        public bool IsPressed { 
            get => _isPressed;
            set { _isPressed = value; ViewModel.Invalidate(); }
        } private bool _isPressed;

        private void OnToggled(object sender, bool isPressed) => Toggled(sender,isPressed);

        public void Attach(string controlId, IRibbonControlStrings strings) {
            _viewModel = ViewModelStore.AttachToggle(controlId, strings, this);
            _viewModel.Toggled += OnToggled;
        }

        bool IBooleanSource.Getter() => IsPressed;
    }

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

namespace PGSolutions.RibbonDispatcher.ComInterfaces { 
    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IViewModelStore)]
    public interface IViewModelStore {
        IRibbonGroup    AttachGroup(string controlId, IRibbonControlStrings strings);
        IRibbonButton   AttachButton(string controlId, IRibbonControlStrings strings);
        IRibbonToggle   AttachToggle(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonToggle   AttachCheckBox(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings, IIntegerSource source);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonGroupModel)]
    public interface IRibbonGroupModel {
        IRibbonGroup ViewModel { get; }

        void Attach(string controlId, IRibbonControlStrings strings);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonButtonModel)]
    public interface IRibbonButtonModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event ClickedEventHandler Clicked;

        IRibbonButton ViewModel { get; }

        void Attach(string controlId, IRibbonControlStrings strings);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonToggleModel)]
    public interface IRibbonToggleModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event ToggledEventHandler Toggled;

        IRibbonToggle ViewModel { get; }
        bool IsPressed { get; set; }

        void Attach(string controlId, IRibbonControlStrings strings);
    }
}
