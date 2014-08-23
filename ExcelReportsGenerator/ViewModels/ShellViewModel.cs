#region

using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

using BenMVVM;

using ExcelReportsGenerator.Common;
using ExcelReportsGenerator.Common.DependenyInjection;
using ExcelReportsGenerator.Common.EventArgsHelper;
using ExcelReportsGenerator.Models;
using ExcelReportsGenerator.Properties;

#endregion

namespace ExcelReportsGenerator.ViewModels
{
  /// <summary>
  ///   The shell view model.
  /// </summary>
  internal class ShellViewModel : ViewModelBase
  {
    #region Fields

    /// <summary>
    ///   The container catalog.
    /// </summary>
    private readonly ContainerCatalog _containerCatalog = new ContainerCatalog();

    /// <summary>
    ///   The _navigate items.
    /// </summary>
    private ObservableCollection<TreeViewModel> _navigateItems;

    /// <summary>
    ///   The _selected navigation item.
    /// </summary>
    private TreeViewModel _selectedNavigationItem;

    /// <summary>
    ///   The _shell content.
    /// </summary>
    private TabControlModel _selectedTabControl;

    /// <summary>
    ///   The _selected tab index.
    /// </summary>
    private int _selectedTabIndex;

    /// <summary>
    ///   The _shell contents
    /// </summary>
    private ObservableCollection<TabControlModel> _shellContents;

    #endregion

    #region Constructors and Destructors

    /// <summary>
    ///   Initializes a new instance of the <see cref="ShellViewModel" /> class.
    /// </summary>
    public ShellViewModel()
    {
      this._containerCatalog.Register<IContainerCatalog>(this._containerCatalog);

      var defaultSelectedNavigation = new TreeViewModel
                                        {
                                          Title = "Excel Reports", 
                                          IsSelected = true, 
                                          ImageSource =
                                            "/ExcelReportsGenerator;component/Resources/Images/ExelIcon.jpg",
                                          DllKey = "ExcelReportsGenerator.ViewModels.ExcelReportViewModel"
                                        };

      this.NavigateItems = new ObservableCollection<TreeViewModel>
                             {
                               defaultSelectedNavigation, 
                               new TreeViewModel
                                 {
                                   Title = "Text Reports", 
                                   ImageSource =
                                     "/ExcelReportsGenerator;component/Resources/Images/File-Text-icon.png", 
                                   DllKey =
                                     "ExcelReportsGenerator.ViewModels.TextReportViewModel"
                                 }
                             };
      this.SelectedNavigationItem = defaultSelectedNavigation;

      this.NotifyPropertyChanged(() => this.NavigateItems);
    }

    #endregion

    #region Public Properties

    /// <summary>
    ///   Gets the close command.
    /// </summary>
    /// <value>
    ///   The close command.
    /// </value>
    public ICommand CloseCommand
    {
      get
      {
        return new RelayCommand(param => this.CloseApplication());
      }
    }

    /// <summary>
    ///   Gets or sets the navigate items.
    /// </summary>
    /// <value>
    ///   The navigate items.
    /// </value>
    public ObservableCollection<TreeViewModel> NavigateItems
    {
      get
      {
        return this._navigateItems;
      }

      set
      {
        this._navigateItems = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets the open command.
    /// </summary>
    /// <value>
    ///   The open command.
    /// </value>
    public ICommand OpenCommand
    {
      get
      {
        return new RelayCommand(param => this.OpenReportHandler());
      }
    }

    /// <summary>
    ///   Gets the report generate command.
    /// </summary>
    /// <value>
    ///   The report generate command.
    /// </value>
    public ICommand ReportGenerateCommand
    {
      get
      {
        return new RelayCommand(param => this.ReportGenerateHandler());
      }
    }

    /// <summary>
    ///   Gets or sets the selected navigation item.
    /// </summary>
    /// <value>
    ///   The selected navigation item.
    /// </value>
    public TreeViewModel SelectedNavigationItem
    {
      get
      {
        return this._selectedNavigationItem;
      }

      set
      {
        this._selectedNavigationItem = value;
        this.SelectedNavigationChangeHandler();
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets the selected navigation item command.
    /// </summary>
    /// <value>
    ///   The selected navigation item command.
    /// </value>
    public ICommand SelectedNavigationItemCommand
    {
      get
      {
        return new RelayCommand(param => this.SelectedNavigationChangeHandler());
      }
    }

    /// <summary>
    ///   Gets or sets the content of the shell.
    /// </summary>
    /// <value>
    ///   The content of the shell.
    /// </value>
    public TabControlModel SelectedTabControl
    {
      get
      {
        return this._selectedTabControl;
      }

      set
      {
        if (this._selectedTabControl == null)
        {
          this._selectedTabControl = this._containerCatalog.ResolveDependencies(value);
        }

        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the index of the selected tab.
    /// </summary>
    /// <value>
    ///   The index of the selected tab.
    /// </value>
    public int SelectedTabIndex
    {
      get
      {
        return this._selectedTabIndex;
      }

      set
      {
        this._selectedTabIndex = value;
        this.NotifyPropertyChanged();
      }
    }

    /// <summary>
    ///   Gets or sets the shell contents.
    /// </summary>
    public ObservableCollection<TabControlModel> ShellContents
    {
      get
      {
        return this._shellContents;
      }

      set
      {
        this._shellContents = value;
        this.NotifyPropertyChanged();
      }
    }

    #endregion

    #region Methods

    /// <summary>
    ///   Closes the application.
    /// </summary>
    private void CloseApplication()
    {
      Application.Current.Shutdown();
    }

    /// <summary>
    /// Gets the new type.
    /// </summary>
    /// <param name="typeName">
    /// Name of the type.
    /// </param>
    /// <returns>
    /// Returns the instance of the object using reflection.
    /// </returns>
    private object GetNewType(string typeName)
    {
      if (string.IsNullOrEmpty(typeName))
      {
        throw new ArgumentNullException("typeName", Resources.NullTypeName);
      }

      Type type = Type.GetType(typeName, true);
      object newInstance = Activator.CreateInstance(type);
      return newInstance;
    }

    /// <summary>
    ///   Opens the report handler.
    /// </summary>
    private void OpenReportHandler()
    {
      if (this.SelectedTabControl == null)
      {
        return;
      }

      var shell = this.SelectedTabControl.ShellContent as IShellViewModel;

      if (shell == null)
      {
        return;
      }

      shell.FileBrowserHandler();
    }

    /// <summary>
    ///   Reports the generate handler.
    /// </summary>
    private void ReportGenerateHandler()
    {
      if (this.SelectedTabControl == null)
      {
        return;
      }

      var shell = this.SelectedTabControl.ShellContent as IShellViewModel;

      if (shell == null)
      {
        return;
      }

      shell.ExcelReportGenerator();
    }

    /// <summary>
    ///   Selects the navigation change handler.
    ///   Uses Reflection to bind the Shell Content with the correct View Model.
    /// </summary>
    private void SelectedNavigationChangeHandler()
    {
      if (this.ShellContents == null)
      {
        return;
      }

      if (this.SelectedNavigationItem == null)
      {
        return;
      }

      var shellContentList = this.ShellContents.ToList();

      if (shellContentList.Exists(navItem => navItem.DllKey == this.SelectedNavigationItem.DllKey))
      {
        this.SelectedTabIndex =
          shellContentList.FindIndex(navItem => navItem.DllKey == this.SelectedNavigationItem.DllKey);
        return;
      }

      var shellContent = new TabControlModel
                           {
                             ShellContent = this.GetNewType(this.SelectedNavigationItem.DllKey), 
                             HeaderName = this.SelectedNavigationItem.Title, 
                             IsActive = true, 
                             ImageSource = this.SelectedNavigationItem.ImageSource,
                             DllKey = this.SelectedNavigationItem.DllKey
                           };

      shellContent.OnCloseTab += this.ShellContentOnCloseTab;
      this.ShellContents.Add(shellContent);

      shellContentList = this.ShellContents.ToList();
      this.SelectedTabIndex = shellContentList.FindLastIndex(navItem => navItem.DllKey == this.SelectedNavigationItem.DllKey);
    }

    /// <summary>
    /// Handles the OnCloseTab event of the shellContent control.
    /// </summary>
    /// <param name="sender">
    /// The source of the event.
    /// </param>
    /// <param name="eventArgs">
    /// The <see cref="Common.EventArgsHelper.CloseTabEventArgs"/> instance containing the event data.
    /// </param>
    private void ShellContentOnCloseTab(object sender, CloseTabEventArgs eventArgs)
    {
      if (eventArgs == null)
      {
        return;
      }

      if (this.ShellContents == null)
      {
        return;
      }

      switch (eventArgs.CloseTab)
      {
        case CloseTab.Close:
          this.ShellContents.Remove(sender as TabControlModel);
          break;
        case CloseTab.CloseAllButThis:
          break;
        case CloseTab.CloseAll:
          break;
      }

      this.SelectedTabIndex = 0;

      var selectedNode = this.NavigateItems.FirstOrDefault(node => node.DllKey == "ExcelReportsGenerator.ViewModels.ExcelReportViewModel");

      if (selectedNode == null)
      {
        return;
      }

      selectedNode.IsSelected = true;
      this.SelectedNavigationItem = selectedNode;
    }

    #endregion
  }
}