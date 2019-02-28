////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon static ComboBox controls.</summary>
    [Description("The COM visible Model for Ribbon static ComboBox controls")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent))]
    [ComDefaultInterface(typeof(IStaticComboBoxModel))]
    [Guid(Guids.StaticComboBoxModel)]
    public sealed class StaticComboBoxModel: ControlModel<IStaticComboBoxSource,IStaticComboBoxVM>, IStaticComboBoxModel,
            IStaticComboBoxSource {
        internal StaticComboBoxModel(Func<string, StaticComboBoxVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings) { }

        public IStaticComboBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited += OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items => ViewModel.Items;

        public int FindId(string id)
        => Items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text { get; set; } = "";

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control, text);
        #endregion
    }
}
