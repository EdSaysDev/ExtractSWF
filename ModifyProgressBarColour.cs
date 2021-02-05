// Licensed under CC-BY-SA 4.0. From https://stackoverflow.com/a/9753302/ by https://stackoverflow.com/users/1032613/user1032613

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ProgressBarControlUtils
{

    // from https://stackoverflow.com/a/9753302/
    public static class ModifyProgressBarColour
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr w, IntPtr l);
        public static void SetState(this ProgressBar pBar, int state)
        {
            SendMessage(pBar.Handle, 1040, (IntPtr)state, IntPtr.Zero);
        }

        public static void SetState(this ProgressBar pBar, ProgressBarState state)
        {
            SetState(pBar, (int)state);
        }
    }

    public enum ProgressBarState : int
    {
        Normal = 1,
        Error = 2,
        Paused = 3
    }
}