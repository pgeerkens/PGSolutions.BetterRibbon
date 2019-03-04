﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.BetterRibbon.Properties;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;
using System.Xml.Linq;

namespace PGSolutions.BetterRibbon {
    using Dictionary = Dictionary<string,ViewModelFactory>;

    /// <summary>The concrete instantiation of <see cref="AbstractDispatcher"/> for <see cref="ThisAddIn"/>.</summary>
    /// <remarks>
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly;
    /// <see cref="IRibbonExtensibility"/>.
    /// 
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface - MUST be COM-visible")]
    [CLSCompliant(false)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    [ComVisible(true)]
    public sealed class Dispatcher: AbstractDispatcher, IRibbonExtensibility {
        internal Dispatcher() : base() { }

        /// </inheritdoc>
        protected override string          RibbonXml      => Resources.Ribbon;

        /// </inheritdoc>
        protected override IResourceLoader ResourceLoader { get; } = new MyResourceManager();

        private            XDocument       RibbonXmlDoc   { get; } = XDocument.Parse(Resources.Ribbon);

        private            Dictionary      Factories      { get; } = new Dictionary();

        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            SaveCurrent(":");

            base.OnRibbonLoad(ribbonUI);
        }

        private void SetCurrentWorkbook(string workbookName) {
            if (Factories.TryGetValue(workbookName, out var factory))
                SetViewModelFactory(factory);
            else
                SetViewModelFactory(Factories[":"]);
            RibbonUI?.Invalidate();
        }

        public override void RegisterWorkbook(string workbookName) {
            if ( ! Factories.TryGetValue(workbookName,out var factory)) {
                factory = ViewModelFactory.ParseXmlDoc(RibbonXmlDoc.Root);
                Factories.Add(workbookName, factory);
            }
            SetCurrentWorkbook(workbookName);
        }

        public void SaveCurrent(string workbookName) {
            if ( ! Factories.ContainsKey(workbookName)) Factories.Add(workbookName, ViewModelFactory);
            RegisterWorkbook(workbookName);
        }

        public void FloatCurrent(string workbookName) {
            if (Factories.ContainsKey(workbookName)) Factories.Remove(workbookName);
        }
    }
}
