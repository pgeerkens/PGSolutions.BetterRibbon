////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

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

        /// <summary>.</summary>
        protected override string RibbonXml    => Resources.Ribbon;

        private XDocument         RibbonXmlDoc { get; } = XDocument.Parse(Resources.Ribbon);

        private Dictionary        Factories    { get; } = new Dictionary();

        /// <summary>The <see cref="IResourceLoader"/> for common shared resources.</summary>
        private IResourceLoader ResourceLoader { get; } = new MyResourceManager();

        /// inheritdoc/>
        public override string GetCustomUI(string RibbonID) {
            SetViewModelFactory(RibbonXml.ParseXmlTabs());

            return base.GetCustomUI(RibbonID);
        }

        public override void RegisterWorkbook(string workbookName) {
            if ( ! Factories.TryGetValue(workbookName,out var factory)) {
                factory = RibbonXmlDoc.ParseXmlTabs();
                Factories.Add(workbookName, factory);
            }
            SetViewModelFactory(factory);
            RibbonUI.Invalidate();
        }

        /// <inheritdoc/>
        public override object LoadImage(string ImageId) => ResourceLoader.GetImage(ImageId);
    }
}
