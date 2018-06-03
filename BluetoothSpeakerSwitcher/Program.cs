using System.Diagnostics;
using AutoIt;

namespace BluetoothSpeakerSwitcher
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1) return;
            if(!int.TryParse(args[0], out var dropDown)) return;
            new Process
            {
                StartInfo =
                {
                    FileName = @"C:\Windows\System32\control.exe",
                    Arguments = "/name Microsoft.Sound /page playback"
                }
            }.Start();
            AutoItX.WinWait("Sound");
            var winHandle = AutoItX.WinGetHandle("Sound");
            var listView = AutoItX.ControlGetHandle(winHandle, "[CLASS:SysListView32; INSTANCE:1]");
            var btnSetDefault = AutoItX.ControlGetHandle(winHandle, "[CLASS:Button; INSTANCE:2]");
            var btnOk = AutoItX.ControlGetHandle(winHandle, "[CLASS:Button; INSTANCE:4]");
            AutoItX.ControlFocus(winHandle, listView);
            AutoItX.ControlSend(winHandle, listView, "{UP}");
            for (var i = 0; i < dropDown; i++)
            {
                AutoItX.ControlSend(winHandle, listView, "{DOWN}");
            }
            AutoItX.ControlClick(winHandle, btnSetDefault);
            AutoItX.ControlClick(winHandle, btnOk);
        }
    }
}
