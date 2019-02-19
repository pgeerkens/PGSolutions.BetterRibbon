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
    public class SelectableItem : RibbonCommon<ISelectableItemSource>, ISelectableItem,
            IActivatable<ISelectableItemSource,SelectableItem>, IClickable, IImageable {
        /// <summary>TODO</summary>
        internal SelectableItem(string ItemId) : base(ItemId) { }

        #region IActivatable implementation
        SelectableItem IActivatable<ISelectableItemSource,SelectableItem>.Attach(ISelectableItemSource source)
        => Attach<SelectableItem>(source);

        public override void Detach() => base.Detach();
        #endregion

        #region IClickable implementation
        /// <summary>The Clicked event source for COM clients</summary>
        public event EventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked() => Clicked?.Invoke(this,EventArgs.Empty);

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked(object sender, EventArgs e) => Clicked?.Invoke(this,e);
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public bool IsImageable => true;

        /// <inheritdoc/>
        public ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? true;

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
