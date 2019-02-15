using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISelectableItem))]
    [Guid(Guids.SelectableItem)]
    public class SelectableItem : RibbonCommon<ISelectableItemSource>,
            IActivatable<SelectableItem, ISelectableItemSource>, ISelectableItem, IImageable {
        /// <summary>TODO</summary>
        internal SelectableItem(string ItemId) : base(ItemId) { }

        #region IActivatable implementation
        SelectableItem IActivatable<SelectableItem, ISelectableItemSource>.Attach(ISelectableItemSource source)
        => Attach<SelectableItem>(source);

        public override void Detach() => base.Detach();
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public bool IsImageable => true;

        /// <inheritdoc/>
        public object Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? true;

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
