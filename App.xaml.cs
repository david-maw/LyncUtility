using System;
using System.Threading;
using System.Windows;

namespace LyncUtility
{
   /// <summary>
   /// Interaction logic for App.xaml
   /// </summary>
   public partial class App : Application
   {
      /// <summary>An event name only we would use.</summary>
      private const string UniqueEventName = "LyncUtility-{DFF34688-ED19-4298-81B0-9F8EEF60FFAD}";

      /// <summary>The event wait handle.</summary>
      private EventWaitHandle eventWaitHandle;

      /// <summary>The app startup method, called via the Startup event when the App is initiated.</summary>
      /// <param name="sender">The sender (same as 'this').</param>
      /// <param name="e">The startup event args - command line parameters and not much else.</param>
      private void AppOnStartup(object sender, StartupEventArgs e)
      {
         bool isOwned;
         eventWaitHandle = new EventWaitHandle(false, EventResetMode.AutoReset, UniqueEventName, out isOwned);

         if (isOwned) // This means we are the first instance to run
         {
            // Spawn a thread which will be waiting for an event from any subsequent instances
            var thread = new Thread(
                () =>
                {
                   while (eventWaitHandle.WaitOne())
                   {
                      Current.Dispatcher.BeginInvoke(
                          (Action)(() => ((MainWindow)Current.MainWindow).BringToForeground()));
                   }
                });

            // It is important mark the new thread as a background thread otherwise it will prevent the app from exiting.
            thread.IsBackground = true;
            thread.Start();
         }
         else
         {
            // Notify the other instance so it can bring itself to the foreground.
            eventWaitHandle.Set();
            // Terminate this instance immediately (do not even initialize the main window).
            Environment.Exit(1);
         }
      }
   }
}
