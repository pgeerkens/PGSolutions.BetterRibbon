////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ITextEditEvents))]
    [ComDefaultInterface(typeof(IEditBoxModel))]
    [Guid(Guids.EditBoxModel)]
    public class EditBoxModel : RibbonControlModel<IEditBoxSource, EditBoxVM>,
            IEditBoxModel, IEditBoxSource {
        public EditBoxModel(Func<string, EditBoxVM> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event EditedEventHandler Edited;

        public string Text       { get; } = "";

        public IEditBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited+= OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnEdited(object sender, string text) => Edited?.Invoke(sender,text);
    }
}
namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.ITextEditEvents)]
    public interface ITextEditEvents {
        /// <summary>Fired when the associated control is clicked by the user.</summary>
        [Description("Fired when the associated control is clicked by the user.")]
        void Edited(object sender, string text);
    }

    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IEditBoxModel)]
    public interface IEditBoxModel: IRibbonCommonSource {
        /// <summary>Gets the {IRibbonControlStrings} for this control.</summary>
        new IRibbonControlStrings Strings {
            [Description("Gets the {IRibbonControlStrings} for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        new bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        new bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IEditBoxModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }
}
