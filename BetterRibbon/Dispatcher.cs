////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.BetterRibbon.Properties;

namespace PGSolutions.BetterRibbon {

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
    public sealed class CustomDispatcher: AbstractCustomDispatcher, IRibbonExtensibility {
        internal CustomDispatcher() : base() { }

        /// <inheritdoc/>
        protected override string          RibbonXml      => Resources.Ribbon;

        /// <inheritdoc/>
        public    override IResourceLoader ResourceLoader { get; } = new MyResourceManager();
    }
}
