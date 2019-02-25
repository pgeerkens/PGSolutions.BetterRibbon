////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly; <see cref="IRibbonExtensibility"/>.
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    public sealed class BetterRibbonViewModel: AbstractDispatcher, IRibbonExtensibility {
        internal BetterRibbonViewModel(string controlId) : base(controlId, new MyResourceManager()) { }

        /// <inheritdoc/>
        protected override string RibbonXml => Resources.Ribbon;
    }
}
