using System;
using System.ComponentModel;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RedInk.OutlookAddIn
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class Globals
    {
        private static ThisAddIn _thisAddIn;
        private static Factory _factory;
        private static ThisRibbonCollection _ribbons;

        internal static ThisAddIn ThisAddIn
        {
            get => _thisAddIn;
            set
            {
                if (_thisAddIn == null)
                {
                    _thisAddIn = value;
                }
                else if (!ReferenceEquals(_thisAddIn, value))
                {
                    throw new NotSupportedException("ThisAddIn has already been initialized.");
                }
            }
        }

        internal static Factory Factory
        {
            get => _factory;
            set
            {
                if (_factory == null)
                {
                    _factory = value;
                }
                else if (!ReferenceEquals(_factory, value))
                {
                    throw new NotSupportedException("Factory has already been initialized.");
                }
            }
        }

        internal static ThisRibbonCollection Ribbons
        {
            get
            {
                if (_ribbons == null)
                {
                    _ribbons = new ThisRibbonCollection(_factory?.GetRibbonFactory());
                }

                return _ribbons;
            }
        }
    }

    [EditorBrowsable(EditorBrowsableState.Never)]
    internal sealed partial class ThisRibbonCollection : RibbonCollectionBase
    {
        internal ThisRibbonCollection(RibbonFactory factory)
            : base(factory)
        {
        }

        internal ThisRibbonCollection this[Outlook.Inspector inspector]
        {
            get => GetRibbonContextCollection<ThisRibbonCollection>(inspector);
        }

        internal ThisRibbonCollection this[Outlook.Explorer explorer]
        {
            get => GetRibbonContextCollection<ThisRibbonCollection>(explorer);
        }
    }
}
