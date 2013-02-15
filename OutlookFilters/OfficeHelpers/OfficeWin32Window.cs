using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Interop;

namespace OutlookFilters.OfficeHelpers
{
    ///<summary>
    /// This class retrieves the IWin32Window from the current active Office window.
    /// This could be used to set the parent for Windows Forms and MessageBoxes.
    ///</summary>
    public class OfficeWin32Window : IWin32Window
    {
        #region DLLImport - user32
        ///<summary>
        /// Retrieves a handle to the top-level window whose class name and window name
        /// match the specified strings. This function does not search child windows. This
        /// function does not perform a case-sensitive search.
        ///</summary>
        ///<param name="lpClassName">The classname of the window (use Spy++)</param>
        ///<param name="lpWindowName">The The window name (the window's title).</param>
        ///<returns>Returns a valid window handle; NULL otherwise.</returns>
        [DllImport("user32")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        #endregion

        #region IWin32Window Implementation
        public IntPtr Handle { get; private set; }
        #endregion


        #region Constructors
        public OfficeWin32Window(object windowObject)
        {
            if (windowObject != null)
            {
                string caption = windowObject.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, null, windowObject, null).ToString();

                Handle = FindWindow("rctrl_renwnd32\0", caption);
            }
            else
            {
                Handle = IntPtr.Zero;
            }
        }
        #endregion
    }
}
