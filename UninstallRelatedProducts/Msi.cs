using System.ComponentModel;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Reflection;
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

        /// <summary>
        /// Uninstalls the specified product
        /// </summary>
        /// <param name="productCode">Product code of the product to uninstall</param>
        /// <param name="silent">Whether to suppress the UI</param>
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
        /// Gets the version of the specified product
        /// </summary>
        /// <param name="productCode">Product code of the product to query</param>
        /// <returns>Product version, never null</returns>
        public static Version GetVersion(Guid productCode)
        {
            var version = GetInfo(productCode, Native.INSTALLPROPERTY_VERSIONSTRING);
            return Version.Parse(version);
        }

        /// <summary>
        /// Gets whether the specified product was installed for all users
        /// </summary>
        /// <param name="productCode">Product code of the product to query</param>
        /// <returns>
        /// true if the product was installed for all users (per-machine); otherwise false.
        /// </returns>
        public static bool IsAllUsers(Guid productCode)
        {
            var assignmentType = GetInfo(productCode, Native.INSTALLPROPERTY_ASSIGNMENTTYPE);
            switch (assignmentType)
            {
                case "0": return false;
                case "1": return true;
                default: throw new InvalidOperationException("Invalid assignment type value: " + assignmentType);
            }
        }

        /// <summary>
        /// Gets an MSI property for the specified product
        /// </summary>
        /// <param name="productCode">Product code of the product to query</param>
        /// <param name="property">MSI property to retrieve</param>
        /// <returns>Property value, never null</returns>
        static string GetInfo(Guid productCode, string property)
        {
            // Product code must be in registry format, enclosed in curly braces
            var productCodeString = productCode.ToString("B");

            var len = 256;
            var buffer = new StringBuilder(len);
            var result = Native.MsiGetProductInfo(productCodeString, property, buffer, ref len);

            if (result == Native.ERROR_MORE_DATA)
            {
                // Buffer wasn't big enough; resize and try again
                ++len;
                buffer = new StringBuilder(len);
                result = Native.MsiGetProductInfo(productCodeString, property, buffer, ref len);
            }

            if (result != Native.ERROR_SUCCESS)
                throw new Win32Exception((int)result);

            return buffer.ToString();
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
            [DllImport("msi.dll")]
            internal static extern UInt32 MsiSetInternalUI(UInt32 dwUILevel, ref IntPtr phWnd);

            /// <summary>
            /// The <see cref="MsiGetProductInfo"/> function returns product information for published and installed products.
            /// </summary>
            /// <param name="szProduct">
            /// Specifies the product code for the product.
            /// </param>
            /// <param name="szProperty">
            /// Specifies the property to be retrieved.
            /// 
            /// <para>The <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa371225(v=vs.85).aspx">Required Properties</see>
            /// are guaranteed to be available, but other properties are available only if that property is set. For more information,
            /// see <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa370889(v=vs.85).aspx">Properties</see>.</para>
            /// </param>
            /// <param name="lpValueBuf">
            /// Pointer to a buffer that receives the property value. This parameter can be null.
            /// </param>
            /// <param name="pcchValueBuf">
            /// Pointer to a variable that specifies the size, in characters, of the buffer pointed to by the lpValueBuf parameter.
            /// On input, this is the full size of the buffer, including a space for a terminating null character. If the buffer
            /// passed in is too small, the count returned does not include the terminating null character.
            /// 
            /// <para>If lpValueBuf is null, pcchValueBuf can be null. In this case, the function checks that the property is registered
            /// correctly with the product.</para>
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
            ///     <term>ERROR_MORE_DATA</term>
            ///     <description>A buffer is too small to hold the requested data.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_SUCCESS</term>
            ///     <description>The function completed successfully.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_UNKNOWN_PRODUCT</term>
            ///     <description>The product is unadvertised or uninstalled.</description>
            /// </item>
            /// <item>
            ///     <term>ERROR_UNKNOWN_PROPERTY</term>
            ///     <description>The property is unrecognized.
            ///     <para>Note: The <see cref="MsiGetProductInfo"/> function returns ERROR_UNKNOWN_PROPERTY if
            ///     the application being queried is advertised and not installed.</para></description>
            /// </item>
            /// </list>
            /// </returns>
            [DllImport("msi.dll", CharSet = CharSet.Auto)]
            internal static extern UInt32 MsiGetProductInfo(string szProduct, string szProperty, [Out] StringBuilder lpValueBuf, ref Int32 pcchValueBuf);

            /// <summary>
            /// The operation completed successfully.
            /// </summary>
            public const UInt32 ERROR_SUCCESS = 0x00000000;

            /// <summary>
            /// More data is available.
            /// </summary>
            public const UInt32 ERROR_MORE_DATA = 0x000000EA;

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

            /// <summary>
            /// Product version. For more information, see the <see cref="http://msdn.microsoft.com/en-us/library/windows/desktop/aa370859(v=vs.85).aspx">ProductVersion</see> property.
            /// </summary>
            public const string INSTALLPROPERTY_VERSIONSTRING = "VersionString";

            /// <summary>
            /// Equals 0 (zero) if the product is advertised or installed per-user.
            /// <para>Equals 1 (one) if the product is advertised or installed per-machine for all users.</para>
            /// </summary>
            public const string INSTALLPROPERTY_ASSIGNMENTTYPE = "AssignmentType";
        }
    }
}