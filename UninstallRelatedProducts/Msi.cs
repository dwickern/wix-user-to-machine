using System.ComponentModel;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace UninstallRelatedProducts
{
    /// <summary>
    /// Wrappers to native MSI functionality
    /// </summary>
    public static class Msi
    {
        /// <summary>
        /// Gets all of the product codes which are related to the specified upgrade code.
        /// </summary>
        /// <param name="upgradeCode">Upgrade code to query</param>
        /// <returns></returns>
        public static IEnumerable<Guid> GetRelatedProducts(Guid upgradeCode)
        {
            // Upgrade code must be in registry format, enclosed in curly braces
            var upgradeCodeString = upgradeCode.ToString("B");
            var buffer = new StringBuilder(39);
            var i = 0;

            while (true)
            {
                var result = Native.MsiEnumRelatedProducts(upgradeCodeString, 0, i++, buffer);
                switch (result)
                {
                    case Native.ERROR_SUCCESS:
                        yield return Guid.Parse(buffer.ToString());
                        break;
                    case Native.ERROR_NO_MORE_ITEMS:
                        yield break;
                    default:
                        throw new Win32Exception((int)result);
                }
            }
        }

        public static void Uninstall(Guid productCode, bool silent)
        {
            if (silent)
            {
                // HACK should set the UI level back when finished uninstalling
                var hwnd = IntPtr.Zero;
                Native.MsiSetInternalUI(Native.INSTALLUILEVEL_NONE, ref hwnd);
            }

            // Product code must be in registry format, enclosed in curly braces
            var productCodeString = productCode.ToString("B");
            var result = Native.MsiConfigureProduct(productCodeString, Native.INSTALLLEVEL_DEFAULT, Native.INSTALLSTATE_ABSENT);
            if (result != Native.ERROR_SUCCESS)
                throw new Win32Exception((int)result);
        }

        /// <summary>
        /// P/Invoke definitions
        /// </summary>
        private static class Native
        {
            /// <summary>
            /// The <see cref="MsiEnumRelatedProducts"/> function enumerates products with a specified upgrade code.
            /// This function lists the currently installed and advertised products that have the specified
            /// UpgradeCode property in their Property table.
            /// </summary>
            /// <param name="lpUpgradeCode">
            /// The null-terminated string specifying the upgrade code of related products that the installer is to enumerate.
            /// </param>
            /// <param name="dwReserved">
            /// This parameter is reserved and must be 0.
            /// </param>
            /// <param name="iProductIndex">
            /// The zero-based index into the registered products.
            /// </param>
            /// <param name="lpProductBuf">
            /// A buffer to receive the product code GUID. This buffer must be 39 characters long.
            /// The first 38 characters are for the GUID, and the last character is for the terminating null character.
            /// </param>
            /// <returns>
            /// <list type="table">
            /// <listheader>
            ///     <term>Value</term>
            ///     <description>Meaning</description>
            /// </listheader>
            /// <item>
            ///     <term>ERROR_BAD_CONFIGURATION</term>
            ///     <description>The configuration data is corrupt.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_INVALID_PARAMETER</term>
            ///     <description>An invalid parameter was passed to the function.</description>
            /// </item>
            /// <item>
            ///     <term><see cref="ERROR_NO_MORE_ITEMS"/></term>
            ///     <description>There are no products to return.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_NOT_ENOUGH_MEMORY</term>
            ///     <description>The system does not have enough memory to complete the operation.
            ///     Available starting with Windows Server 2003.</description>
            /// </item>
            /// <item>
            ///     <term><see cref="ERROR_SUCCESS"/></term>
            ///     <description>A value was enumerated.</description>
            /// </item>
            /// </list>
            /// </returns>
            [DllImport("msi.dll", CharSet = CharSet.Auto)]
            internal static extern UInt32 MsiEnumRelatedProducts(string lpUpgradeCode, int dwReserved, int iProductIndex, StringBuilder lpProductBuf);

            /// <summary>
            /// The <see cref="MsiConfigureProduct"/> function installs or uninstalls a product.
            /// </summary>
            /// <param name="szProduct">
            /// Specifies the product code for the product to be configured.
            /// </param>
            /// <param name="iInstallLevel">
            /// Specifies how much of the product should be installed when installing the product to its default state.
            /// The iInstallLevel parameters are ignored, and all features are installed, if the eInstallState parameter
            /// is set to any value other than INSTALLSTATE_DEFAULT.
            /// </param>
            /// <param name="eInstallState">
            /// Specifies the installation state for the product.
            /// </param>
            /// <returns>
            /// <list type="table">
            /// <listheader>
            ///     <term>Value</term>
            ///     <description>Meaning</description>
            /// </listheader>
            /// <item>
            ///     <term>ERROR_INVALID_PARAMETER</term>
            ///     <description>An invalid parameter is passed to the function.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_SUCCESS</term>
            ///     <description>The function succeeds.</description>
            /// </item>
            /// <item>
            ///     <term>An error that relates to an action</term>
            ///     <description>For more information, see <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa376931(v=vs.85).aspx">Error Codes.</see></description>
            /// </item>
            /// <item>
            ///     <term><see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa369284(v=vs.85).aspx">Initialization Error</see></term>
            ///     <description>An error that relates to initialization.</description>
            /// </item>
            /// </list>
            /// </returns>
            [DllImport("msi.dll", CharSet = CharSet.Auto)]
            internal static extern UInt32 MsiConfigureProduct(string szProduct, Int32 iInstallLevel, Int32 eInstallState);

            /// <summary>
            /// The <see cref="MsiSetInternalUI"/> function enables the installer's internal user interface.
            /// Then this user interface is used for all subsequent calls to user-interface-generating installer
            /// functions in this process. For more information, see <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa372391(v=vs.85).aspx">User Interface Levels</see>.
            /// </summary>
            /// <param name="dwUILevel">
            /// Specifies the level of complexity of the user interface.
            /// </param>
            /// <param name="phWnd">
            /// Pointer to a window. This window becomes the owner of any user interface created.
            /// A pointer to the previous owner of the user interface is returned.
            /// If this parameter is null, the owner of the user interface does not change.</param>
            /// <returns>
            /// The previous user interface level is returned.
            /// If an invalid dwUILevel is passed, then INSTALLUILEVEL_NOCHANGE is returned.
            /// </returns>
            [DllImport("msi.dll", SetLastError = true)]
            internal static extern UInt32 MsiSetInternalUI(UInt32 dwUILevel, ref IntPtr phWnd);

            ///// <summary>
            ///// The INSTALLUI_HANDLER function prototype defines a callback function that the installer calls
            ///// for progress notification and error messages. For more information on the usage of this function
            ///// prototype, a sample code snippet is available in <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa368786(v=vs.85).aspx">Handling Progress Messages Using MsiSetExternalUI</see>.
            ///// </summary>
            ///// <param name="pvContext">
            ///// Pointer to an application context passed to the MsiSetExternalUI function.
            ///// This parameter can be used for error checking.
            ///// </param>
            ///// <param name="iMessageType">
            ///// Specifies a combination of one message box style, one message box icon type, one default button,
            ///// and one installation message type.
            ///// </param>
            ///// <param name="szMessage">
            ///// Specifies the message text.
            ///// </param>
            ///// <returns>
            ///// <list type="table">
            ///// <listheader>
            /////     <term>Value</term>
            /////     <description>Meaning</description>
            ///// </listheader>
            ///// <item>
            /////     <term>-1</term>
            /////     <description>An error was found in the message handler.
            /////     Windows Installer ignores this returned value.</description>
            ///// </item>
            ///// <item>
            /////     <term>0</term>
            /////     <description>No action was taken.</description>
            ///// </item>
            ///// </list>
            ///// The following return values map to the buttons specified by the message box style:
            ///// <para>IDOK IDCANCEL IDABORT IDRETRY IDIGNORE IDYES IDNO</para>
            ///// </returns>
            //internal delegate int InstalluiHandler(IntPtr pvContext, UInt32 iMessageType, [MarshalAs(UnmanagedType.LPTStr)] string szMessage);

            /// <summary>
            /// The operation completed successfully.
            /// </summary>
            public const UInt32 ERROR_SUCCESS = 0x00000000;

            /// <summary>
            /// No more data is available.
            /// </summary>
            public const UInt32 ERROR_NO_MORE_ITEMS = 0x00000103;

            /// <summary>
            /// Completely silent installation.
            /// </summary>
            public const UInt32 INSTALLUILEVEL_NONE = 2;

            /// <summary>
            /// The authored default features are installed.
            /// </summary>
            public const Int32 INSTALLLEVEL_DEFAULT = 0;

            /// <summary>
            /// The product is uninstalled.
            /// </summary>
            public const Int32 INSTALLSTATE_ABSENT = 2;
        }
    }
}