using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace LyncUtility
{
   public partial class MainWindow : Window
   {
      #region Fields
      // Current dispatcher reference for changes in the user interface.
      private Dispatcher dispatcher;
      private LyncClient lyncClient;
      System.Windows.Forms.Timer clock = new System.Windows.Forms.Timer();
      DateTime endTime = DateTime.MinValue;
      ContactAvailability currentAvailability = 0, futureAvailability = 0;
      #endregion
      #region Constructor
      public MainWindow()
      {
         InitializeComponent();

         //Save the current dispatcher to use it for changes in the user interface.
         dispatcher = Dispatcher.CurrentDispatcher;
         // Create a timer, but do not start it yet
         clock.Interval = 1000; // Milliseconds
         clock.Tick += OnTick;

         // We want to use a notify icon in the system tray, but WPF doesn't have one so we will use the one from Windows Forms instead
         var systrayContextMenu = new System.Windows.Forms.ContextMenu();
         systrayContextMenu.MenuItems.Add("E&xit", (s, o) => { this.Close(); });
         systrayContextMenu.MenuItems.Add("&Open", ActivateForm);

         // Create the systray icon
         var notifyIcon = new System.Windows.Forms.NotifyIcon();
         // The Icon property sets the icon that will appear in the systray for this application.
         using (Stream stream = Application.GetResourceStream(new Uri("LyncUtility.ico", UriKind.Relative)).Stream)
         {
            notifyIcon.Icon = new Icon(stream);
         }

         // The ContextMenu property sets the menu that will appear when the systray icon is right clicked.
         notifyIcon.ContextMenu = systrayContextMenu;

         // Set the text that will be displayed when the mouse hovers over the systray icon.
         notifyIcon.Text = "LyncUtility";
         notifyIcon.Visible = true;

         // Allow the double click on the systray icon to activate the form.
         notifyIcon.DoubleClick += ActivateForm;
      }
      #endregion
      #region Handlers for user interface events
      /// <summary>
      /// Handler for the Loaded event of the Window.
      /// Used to initialize the values shown in the user interface, get the Lync client instance
      /// and start listening for events of changes in the client state.
      /// </summary>
      private void Window_Loaded(object sender, RoutedEventArgs e)
      {

         //Set up event for changes in the state of the client
         try
         {
            lyncClient = LyncClient.GetClient();
         }
         catch (ClientNotFoundException clientNotFoundException)
         {
            Trace.WriteLine(clientNotFoundException);
            return;
         }
         catch (NotStartedByUserException notStartedByUserException)
         {
            Console.Out.WriteLine(notStartedByUserException);
            return;
         }
         catch (LyncClientException lyncClientException)
         {
            Console.Out.WriteLine(lyncClientException);
            return;
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Error: " + systemException);
               return;
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }

         lyncClient.StateChanged +=
             new EventHandler<ClientStateChangedEventArgs>(Client_StateChanged);

         //Update the user interface
         UpdateUserInterface(lyncClient.State);
      }

      /// <summary>
      /// Updates the user interface
      /// </summary>
      /// <param name="currentState"></param>
      private void UpdateUserInterface(ClientState currentState)
      {
         //Update the client state in the user interface
         clientStateTextBox.Text = currentState.ToString();

         if (currentState == ClientState.SignedIn)
         {
            //Listen for events of changes of the contact's information
            lyncClient.Self.Contact.ContactInformationChanged +=
                new EventHandler<ContactInformationChangedEventArgs>(SelfContact_ContactInformationChanged);

            GetAvailability();

            //Update the SignInOut button content
            signInOutButton.Content = "Sign Out";
            futureStatePanel.Visibility = Visibility.Visible;
         }
         else
         {
            //Update the SignInOut button content
            signInOutButton.Content = "Sign In";
            futureStatePanel.Visibility = Visibility.Collapsed;
         }
      }

      /// <summary>
      /// Called periodically whenever there's a future state pending
      /// </summary>
      /// <param name="sender"></param>
      /// <param name="e"></param>
      private void OnTick(object sender, EventArgs e)
      {
         TimeSpan timeLeft = endTime - DateTime.Now;
         if (timeLeft > TimeSpan.FromSeconds(1))
         {
            // Just a boring tick, just update the time left
            timeLeftLabel.Content = timeLeft.ToString(@"hh\:mm\:ss");
         }
         else
         {
            try
            {
               if (lyncClient.State == ClientState.SignedOut)
               {
                  Trace.WriteLine("OnTick: Signed out, signing in");
                  //Sign in If the current client state is Signed Out
                  lyncClient.BeginSignIn(null, null, null, SignInCallback, null);
               }
               else if (lyncClient.State == ClientState.SignedIn)
               {
                  Trace.WriteLine("OnTick: Timed out, signed in, setting availability");
                  SetAvailability(futureAvailability);
               }
               else
               {
                  // We've no idea what's going on so just wait a bit and hope it sorts itself out
                  Trace.WriteLine("OnTick: Timed out, waiting a bit because lync client state=" + lyncClient.State);
                  endTime = DateTime.Now + TimeSpan.FromSeconds(5);
                  return;
               }
            }
            catch (LyncClientException lyncClientException)
            {
               Trace.WriteLine("Lync Exception: "+lyncClientException);
            }
            catch (SystemException systemException)
            {
               if (IsLyncException(systemException))
               {
                  // Log the exception thrown by the Lync Model API.
                  Trace.WriteLine("Lync Exception: " + systemException);
               }
               else
               {
                  // Rethrow the SystemException which did not come from the Lync Model API.
                  throw;
               }
            }
            Trace.WriteLine("OnTick: Stopping countdown timer");
            timeLeftLabel.Content = "none";
            clock.Stop();
         }
      }

      private void ActivateForm(object Sender, EventArgs e)
      {
         // Show the form when the user double clicks on the notify icon.
         BringToForeground();
      }

      /// <summary>Brings main window to foreground.</summary>
      public void BringToForeground()
      {
         // Set the WindowState to normal if the form is not visible.
         if (this.WindowState == WindowState.Minimized || this.Visibility == Visibility.Hidden)
         {
            this.ShowInTaskbar = true;
            this.Show();
            this.WindowState = WindowState.Normal;
         }

         // Activate the form and bring it to the foreground.
         this.Activate();
         this.Topmost = true;
         this.Topmost = false;
         this.Focus();
      }

      /// <summary>
      /// Handler for the Click event of the SignInOut Button. Used to sign in or out Lync depending on the current client state.
      /// </summary>
      private void SignInOutButton_Click(object sender, RoutedEventArgs e)
      {
         try
         {
            if (lyncClient.State == ClientState.SignedIn)
            {
               //Sign out If the current client state is Signed In
               lyncClient.BeginSignOut(SignOutCallback, null);
            }
            else if (lyncClient.State == ClientState.SignedOut)
            {
               //Sign in If the current client state is Signed Out
               lyncClient.BeginSignIn(null, null, null, SignInCallback, null);
            }
         }
         catch (LyncClientException lyncClientException)
         {
            Trace.WriteLine(lyncClientException);
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Error: " + systemException);
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }
      }
      private void OnBecomeAvailable(object sender, RoutedEventArgs e)
      {
         futureAvailability = ContactAvailability.Free;
      }

      private void OnBecomeBusy(object sender, RoutedEventArgs e)
      {
         futureAvailability = ContactAvailability.Busy;
      }

      private void OnBecomeDoNotDisturb(object sender, RoutedEventArgs e)
      {
         futureAvailability = ContactAvailability.DoNotDisturb;
      }

      private void OnBecomeAway(object sender, RoutedEventArgs e)
      {
         futureAvailability = ContactAvailability.Away;
      }

      private void OnBecomeAutomatic(object sender, RoutedEventArgs e)
      {
         futureAvailability = ContactAvailability.None;
      }
      private void OnStartButtonClick(object sender, RoutedEventArgs e)
      {
         float delayHours;
         if (!float.TryParse(delayTime.Text, out delayHours))
         {
            delayHours = 1;
            delayTime.Text = delayHours.ToString();
         }

         endTime = DateTime.Now + TimeSpan.FromHours(delayHours);
         clock.Start();
      }
      #endregion
      #region Handlers for Lync events
      /// <summary>
      /// Handler for the ContactInformationChanged event of the contact. Used to update the contact's information in the user interface.
      /// </summary>
      private void SelfContact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
      {
         //Only update the contact information in the user interface if the client is signed in.
         //Ignore other states including transitions (e.g. signing in or out).
         if (lyncClient.State == ClientState.SignedIn)
         {
            //Get from Lync only the contact information that changed.

            if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
            {
               //Use the current dispatcher to read the contact's availability on the UI thread
               dispatcher.BeginInvoke(new Action(GetAvailability));
            }
         }
      }

      /// <summary>
      /// Handler for the StateChanged event of the contact. Used to update the user interface with the new client state.
      /// </summary>
      private void Client_StateChanged(object sender, ClientStateChangedEventArgs e)
      {
         //Use the current dispatcher to update the user interface with the new client state.
         dispatcher.BeginInvoke(new Action<ClientState>(UpdateUserInterface), e.NewState);
      }
      #endregion
      #region Callbacks
      /// <summary>
      /// Callback invoked when LyncClient.BeginSignIn is completed
      /// </summary>
      /// <param name="result">The status of the asynchronous operation</param>
      private void SignInCallback(IAsyncResult result)
      {
         try
         {
            lyncClient.EndSignIn(result);
         }
         catch (LyncClientException e)
         {
            Trace.WriteLine(e);
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Error: " + systemException);
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }

      }

      /// <summary>
      /// Callback invoked when LyncClient.BeginSignOut is completed
      /// </summary>
      /// <param name="result">The status of the asynchronous operation</param>
      private void SignOutCallback(IAsyncResult result)
      {
         try
         {
            lyncClient.EndSignOut(result);
         }
         catch (LyncClientException e)
         {
            Trace.WriteLine(e);
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Error: " + systemException);
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }

      }

      /// <summary>
      /// Callback invoked when Self.BeginPublishContactInformation is completed
      /// </summary>
      /// <param name="result">The status of the asynchronous operation</param>
      private void PublishContactInformationCallback(IAsyncResult result)
      {
         lyncClient.Self.EndPublishContactInformation(result);
      }
      #endregion
      #region Lync Utilities
      /// <summary>
      /// Gets the contact's current availability value from Lync and updates the corresponding elements in the user interface
      /// </summary>
      private void GetAvailability()
      {
         //Get the current availability value from Lync
         try
         {
            currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
         }
         catch (LyncClientException e)
         {
            Trace.WriteLine(e);
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Error: " + systemException);
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }


         if (currentAvailability != 0)
         {
            //React to changed availability
         }
      }
      private void SetAvailability(ContactAvailability NewAvailability)
      {
         Trace.WriteLine("SetAvailability("+NewAvailability.ToString()+")");
         //Add the availability to the contact information items to be published
         Dictionary<PublishableContactInformationType, object> newInformation =
             new Dictionary<PublishableContactInformationType, object>();
         newInformation.Add(PublishableContactInformationType.Availability, NewAvailability);

         //Publish the new availability value
         try
         {
            lyncClient.Self.BeginPublishContactInformation(newInformation, PublishContactInformationCallback, null);
         }
         catch (LyncClientException lyncClientException)
         {
            Trace.WriteLine(lyncClientException);
         }
         catch (SystemException systemException)
         {
            if (IsLyncException(systemException))
            {
               // Log the exception thrown by the Lync Model API.
               Trace.WriteLine("Lync Exception: " + systemException);
            }
            else
            {
               // Rethrow the SystemException which did not come from the Lync Model API.
               throw;
            }
         }
      }

      private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
      {
         if (this.WindowState == WindowState.Normal)
         {// Treat close as minimize
            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = false;
            e.Cancel = true;
            return;
         }
         if (endTime > DateTime.Now)
         {
            var result = MessageBox.Show("Future action is scheduled, do you want to close and ignore it", "LyncUtility", MessageBoxButton.YesNoCancel);
            if (result == MessageBoxResult.Yes)
               return;
            else
               e.Cancel = true;
         }
      }

      private void exitButton_Click(object sender, RoutedEventArgs e)
      {
         App.Current.Shutdown();
      }

      /// <summary>
      /// Identify if a particular SystemException is one of the exceptions which may be thrown
      /// by the Lync Model API.
      /// </summary>
      /// <param name="ex"></param>
      /// <returns></returns>
      private bool IsLyncException(SystemException ex)
      {
         return
             ex is NotImplementedException ||
             ex is ArgumentException ||
             ex is NullReferenceException ||
             ex is NotSupportedException ||
             ex is ArgumentOutOfRangeException ||
             ex is IndexOutOfRangeException ||
             ex is InvalidOperationException ||
             ex is TypeLoadException ||
             ex is TypeInitializationException ||
             ex is InvalidComObjectException ||
             ex is InvalidCastException;
      }
      #endregion
   }
}
