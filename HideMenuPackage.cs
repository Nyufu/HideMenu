using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Task = System.Threading.Tasks.Task;

namespace HideMenu
{
	/// <summary>
	/// This is the class that implements the package exposed by this assembly.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the
	/// IVsPackage interface and uses the registration attributes defined in the framework to
	/// register itself and its components with the shell. These attributes tell the pkgdef creation
	/// utility what data to put into .pkgdef file.
	/// </para>
	/// <para>
	/// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
	/// </para>
	/// </remarks>
	[Guid(PackageGuidString)]
	[PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	[ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
	public sealed class HideMenuPackage : AsyncPackage
	{
		/// <summary>
		/// HideMenuPackage GUID string.
		/// </summary>
		public const string PackageGuidString = "95a39d58-048c-4dac-96df-16385d4fa22d";

		#region Package Members

		private Window _mainWindow;

		private FrameworkElement _titleBar;
		public FrameworkElement TitleBar
		{
			get => _titleBar;
			set
			{
				_titleBar = value;
				UpdateElementHeight(_titleBar);
				AddElementHandlers(_titleBar);
			}
		}

		private FrameworkElement _menuBar;
		public FrameworkElement MenuBar
		{
			get => _menuBar;
			set
			{
				_menuBar = value;
				UpdateElementHeight(_menuBar);
				AddElementHandlers(_menuBar);
			}
		}

		private bool _isVisible;
		public bool IsVisible
		{
			get => _isVisible;
			set
			{
				if (_isVisible == value)
				{
					return;
				}

				_isVisible = value;
				UpdateElementHeight(_titleBar);
				UpdateElementHeight(_menuBar);
			}
		}

		void UpdateElementHeight(FrameworkElement element)
		{
			if (IsVisible)
			{
				element.ClearValue(FrameworkElement.HeightProperty);
			}
			else
			{
				element.Height = 0;
			}
		}

		void AddElementHandlers(FrameworkElement element)
		{
			element.IsKeyboardFocusWithinChanged += OnContainerFocusChanged;
		}

		private void OnContainerFocusChanged(object sender, DependencyPropertyChangedEventArgs e)
		{
			IsVisible = IsAggregateFocusInMenuContainer();
		}

		private void PopupLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
		{
			if (IsVisible && !IsAggregateFocusInMenuContainer())
			{
				IsVisible = false;
			}
		}

		private bool IsAggregateFocusInMenuContainer()
		{
			if (TitleBar.IsKeyboardFocusWithin || MenuBar.IsKeyboardFocusWithin)
			{
				return true;
			}

			for (DependencyObject sourceElement = (DependencyObject)Keyboard.FocusedElement; sourceElement != null; sourceElement = sourceElement.GetVisualOrLogicalParent())
			{
				if (sourceElement == TitleBar || sourceElement == MenuBar)
				{
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		/// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
		/// <param name="progress">A provider for progress updates.</param>
		/// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
		protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
		{
			await base.InitializeAsync(cancellationToken, progress);
			// When initialized asynchronously, the current thread may be a background thread at this point.
			// Do any initialization that requires the UI thread after switching to the UI thread.
			await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

			EventManager.RegisterClassHandler(typeof(UIElement), UIElement.LostKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(PopupLostKeyboardFocus));

			_mainWindow = Application.Current.MainWindow;
			_mainWindow.LayoutUpdated += DetectLayoutElements;
		}

		private void DetectLayoutElements(object sender, EventArgs e)
		{
			if (TitleBar == null)
			{
				MainWindowTitleBar titleBar = _mainWindow.FindDescendants<MainWindowTitleBar>().FirstOrDefault();

				if (titleBar != null)
				{
					TitleBar = titleBar;
				}
			}

			if (TitleBar != null)
			{
				if (MenuBar == null)
				{
					Grid rootGrid = (Grid)TitleBar.Parent;

					if (rootGrid != null)
					{
						foreach (object child in rootGrid.Children)
						{
							if (child.GetType().Name == "DockPanel" && ((DockPanel)child).Name.Length == 0)
							{
								MenuBar = (DockPanel)child;
								_mainWindow.LayoutUpdated -= DetectLayoutElements;
								break;
							}
						}
					}
				}
				else
				{
					_mainWindow.LayoutUpdated -= DetectLayoutElements;
				}
			}
		}

		#endregion
	}
}
